using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjCompileResult(
    byte[] File,
    PresentationProgramResult Program,
    IReadOnlyList<Diagnostic> Diagnostics);

/// <summary>
/// Lowers one validated, source-free PPJ program directly into the native C#
/// Presentation compiler IR. JavaScript never materializes a Presentation
/// object model on this path; the protobuf model is an internal writer IR.
/// </summary>
internal static class PpjAuthoredPresentationCompiler
{
    private const double EmuPerPoint = 12_700d;
    private const double CustomPathUnitsPerPoint = 1_000d;

    internal static PpjCompileResult ValidateOnly(
        PresentationProgramRequest request,
        PpjValidationResult validation)
    {
        var program = validation.Program!;
        var assets = ValidateAssets(program, request.Assets);
        var receipt = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            SourceBound = false,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        if (request.IncludeNodeMap)
            receipt.NodeMapJson = ByteString.CopyFrom(validation.Expansion.NodeMapJson);
        receipt.Assets.Add(assets.Select(asset => asset.Clone()));
        return new([], receipt, []);
    }

    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        EffectiveCodecLimits limits)
    {
        var validation = PpjProgramValidator.Validate(request.ProgramJson.Span);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        var program = validation.Program!;
        if (program.Source is not null)
            throw new CodecException(
                "ppj.sourceBoundCompileRequired",
                "A source-bound PPJ must be compiled against its exact validated source PPTX.",
                "$.source");

        var assets = ValidateAssets(program, request.Assets);
        var catalog = new Catalog(program.Root);
        var envelope = BuildEnvelope(program, validation.Expansion!, catalog, assets);
        var exported = PptxCodec.Export(envelope, limits);
        var file = PpjEmbeddedProgramCodec.Embed(
            exported.File,
            request.ProgramJson.Span,
            validation,
            envelope.Presentation,
            assets,
            limits);
        var fileSha256 = Sha256(file);
        var receipt = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            OutputSha256 = fileSha256,
            SourceBound = false,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        if (request.IncludeNodeMap)
            receipt.NodeMapJson = ByteString.CopyFrom(validation.Expansion.NodeMapJson);
        receipt.Assets.Add(assets.Select(asset => asset.Clone()));
        receipt.ChangedNodeIds.Add(validation.Expansion.Nodes.Select(node => node.Id));
        return new(file, receipt, exported.Diagnostics);
    }

    private static ArtifactEnvelope BuildEnvelope(
        PpjProgramModel program,
        PpjExpansionResult expansion,
        Catalog catalog,
        IReadOnlyList<Asset> assets)
    {
        var presentation = new PresentationArtifact
        {
            Id = program.Meta.Id,
            Name = program.Meta.Title,
            SlideWidthEmu = Emu(program.Design.Width),
            SlideHeightEmu = Emu(program.Design.Height),
            AuthoredTheme = catalog.Theme,
        };

        var expandedByPage = expansion.Pages.ToDictionary(page => page.Id, StringComparer.Ordinal);
        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var expanded = expandedByPage[page.Id];
            var slide = new PresentationSlide
            {
                Id = page.Id,
                Name = DisplayName(page.Name, page.Role, page.Id),
            };
            if (page.Raw.TryGetProperty("hidden", out var hidden)) slide.Hidden = hidden.GetBoolean();
            if (page.Raw.TryGetProperty("background", out var background))
                slide.Background = BuildBackground(background, catalog);
            if (page.Raw.TryGetProperty("notes", out var notes))
                slide.SpeakerNotes = BuildNotes(notes, catalog);

            for (var elementIndex = 0; elementIndex < expanded.Elements.Count; elementIndex++)
                slide.Elements.Add(BuildElement(expanded.Elements[elementIndex], expanded.ElementJson[elementIndex], catalog));

            foreach (var animation in page.Animations)
                slide.Animations.Add(BuildAnimation(animation, expanded.Elements));
            ApplyTransition(slide, page, pageIndex == 0 ? null : presentation.Slides[pageIndex - 1]);
            presentation.Slides.Add(slide);
        }

        AddSections(presentation, program);
        AddCustomShows(presentation, program);
        AddComments(presentation, program);

        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Presentation = presentation,
        };
        envelope.Assets.Add(assets.Select(asset => asset.Clone()));
        return envelope;
    }

    private static PresentationElement BuildElement(PpjElementModel element, JsonElement raw, Catalog catalog)
    {
        RejectProperties(raw, element.Id, "hidden", "locked");
        var output = new PresentationElement
        {
            Id = element.Id,
            Name = DisplayName(element.Name, element.Role, element.Id),
        };
        switch (element)
        {
            case PpjTextElementModel:
                output.Shape = BuildTextShape(element, raw, catalog, "textbox");
                break;
            case PpjShapeElementModel shape:
                output.Shape = BuildShape(shape, raw, catalog);
                break;
            case PpjImageElementModel image:
                output.Image = BuildImage(image, raw, catalog);
                break;
            case PpjChartElementModel chart:
                output.Chart = BuildChart(chart, raw, catalog);
                break;
            case PpjTableElementModel table:
                output.Table = BuildTable(table, raw, catalog);
                break;
            case PpjConnectorElementModel connector:
                output.Connector = BuildConnector(connector, raw, catalog);
                break;
            case PpjGroupElementModel group:
                output.Group = BuildGroup(group, raw, catalog);
                break;
            case PpjPlaceholderElementModel placeholder:
                output.Shape = BuildPlaceholder(placeholder, raw, catalog);
                break;
            case PpjMediaElementModel:
                throw Unsupported(element.Id, "media authoring requires a typed native media compiler");
            case PpjSmartArtElementModel:
                throw Unsupported(element.Id, "source-free SmartArt authoring is not yet a PPJ compiler capability");
            case PpjOleElementModel:
                throw Unsupported(element.Id, "source-free OLE authoring is not yet a PPJ compiler capability");
            case PpjOpaqueElementModel:
                throw Unsupported(element.Id, "opaque elements require a source-bound PPJ compile");
            default:
                throw Unsupported(element.Id, $"expanded element type {element.Type} cannot be authored");
        }
        return output;
    }

    private static PresentationShape BuildTextShape(
        PpjElementModel element,
        JsonElement raw,
        Catalog catalog,
        string geometry)
    {
        var shape = ShapeFrame(element.Frame, geometry);
        var namedStyle = catalog.TextStyle(OptionalString(raw, "styleRef"));
        var inlineStyle = Property(raw, "style");
        shape.TextBody = raw.TryGetProperty("text", out var text)
            ? BuildTextBody(text, namedStyle, inlineStyle, catalog)
            : EmptyTextBody(namedStyle, inlineStyle);
        shape.Text = Flatten(shape.TextBody);
        if (raw.TryGetProperty("fill", out var fill)) ApplyTextBoxFill(shape, fill, catalog);
        if (raw.TryGetProperty("stroke", out var stroke)) ApplyLine(shape, stroke, catalog);
        else shape.LineStyle = "none";
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static PresentationShape BuildShape(PpjShapeElementModel element, JsonElement raw, Catalog catalog)
    {
        var geometry = element.GeometryKind == "custom" ? "custom" : element.GeometryPreset ?? "rect";
        var shape = ShapeFrame(element.Frame, geometry);
        var namedStyle = catalog.ShapeStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id, raw.TryGetProperty("text", out _));
        if (raw.TryGetProperty("text", out var text))
        {
            shape.TextBody = BuildTextBody(text, null, Property(raw, "textStyle"), catalog);
            shape.Text = Flatten(shape.TextBody);
        }
        if (geometry == "custom") ApplyCustomGeometry(shape, raw.GetProperty("geometry"), element.Id);
        else if (raw.GetProperty("geometry").TryGetProperty("adjustments", out var adjustments) && adjustments.GetArrayLength() > 0)
            throw Unsupported(element.Id, "preset-geometry adjustments are not yet compiler-owned");
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static PresentationImage BuildImage(PpjImageElementModel element, JsonElement raw, Catalog catalog)
    {
        var image = new PresentationImage
        {
            AssetId = catalog.NativeAssetId(element.AssetId),
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
        };
        ApplyImageCrop(image, element, raw, catalog);
        if (raw.TryGetProperty("opacity", out var opacity))
            image.OpacityThousandthPercent = Opacity(opacity.GetDouble());
        if (raw.TryGetProperty("mask", out var mask))
        {
            if (mask.GetProperty("kind").GetString() != "preset")
                throw Unsupported(element.Id, "custom image masks require the native custom-geometry picture compiler");
            if (mask.TryGetProperty("adjustments", out var adjustments) && adjustments.GetArrayLength() > 0)
                throw Unsupported(element.Id, "adjusted image-mask geometry is not yet compiler-owned");
            image.MaskPreset = mask.GetProperty("preset").GetString()!;
        }
        if (raw.TryGetProperty("border", out var border))
        {
            var color = catalog.Color(border.GetProperty("color"));
            image.Border = new PresentationImageBorder
            {
                ColorRgb = color.Rgb,
                WidthEmu = Emu(border.GetProperty("width").GetDouble()),
                Style = LineStyle(OptionalString(border, "dash")),
                Cap = OptionalString(border, "cap") ?? string.Empty,
                Join = OptionalString(border, "join") ?? string.Empty,
            };
            var borderOpacity = OptionalDouble(border, "opacity") ?? color.Alpha;
            if (borderOpacity < 1) image.Border.OpacityThousandthPercent = Opacity(borderOpacity);
        }
        if (raw.TryGetProperty("shadow", out var shadow))
        {
            var color = catalog.Color(shadow.GetProperty("color"));
            image.Shadow = new PresentationShadow
            {
                ColorRgb = color.Rgb,
                BlurRadiusEmu = Emu(shadow.GetProperty("blur").GetDouble()),
                DistanceEmu = Emu(shadow.GetProperty("distance").GetDouble()),
                DirectionAngle60000 = Angle(shadow.GetProperty("angle").GetDouble()),
                OpacityThousandthPercent = Opacity(OptionalDouble(shadow, "opacity") ?? color.Alpha),
            };
        }
        if (element.Frame.Rotation != 0 || element.Frame.FlipH || element.Frame.FlipV)
        {
            image.Transform = new PresentationImageTransform();
            if (element.Frame.Rotation != 0) image.Transform.RotationAngle60000 = Angle(element.Frame.Rotation);
            if (element.Frame.FlipH) image.Transform.FlipHorizontal = true;
            if (element.Frame.FlipV) image.Transform.FlipVertical = true;
        }
        ApplyAccessibility(image, element.Accessibility);
        return image;
    }

    private static PresentationChart BuildChart(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        RejectUnsupportedFrameTransform(element.Id, element.Frame, "chart");
        if (element.Data.Series.Any(series => series.Values.Any(value => value is null)))
            throw Unsupported(element.Id, "null chart values require a missing-value-aware native chart cache");
        if (element.ChartType is "radar" or "waterfall")
            throw Unsupported(element.Id, $"{element.ChartType} chart authoring is not yet compiler-owned");
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "rich chart-title formatting is not yet compiler-owned");

        var chart = new PresentationChart
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Type = ChartType(element.ChartType),
            Title = element.Title is null ? string.Empty : Flatten(element.Title),
            BarDirection = element.ChartType == "bar" ? "bar" : element.ChartType == "column" ? "column" : string.Empty,
        };
        chart.Categories.Add(element.Data.Categories.Select(CategoryText));
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyChartStyle(chart, namedStyle, inlineStyle, catalog, element.Id);

        var rawXAxis = Property(raw, "xAxis");
        var rawYAxis = Property(raw, "yAxis");
        if (rawXAxis is not null || rawYAxis is not null)
        {
            if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
                throw Unsupported(element.Id, "pie and doughnut charts cannot define axes");
            if (FirstProperty(inlineStyle, namedStyle, "showCategoryAxis") is not null ||
                FirstProperty(inlineStyle, namedStyle, "showValueAxis") is not null)
                throw Unsupported(element.Id, "structured axes cannot be combined with legacy axis visibility fields");
            chart.XAxis = rawXAxis is { } xAxis ? BuildChartAxis(xAxis) : new SpreadsheetChartAxisArtifact();
            chart.YAxis = rawYAxis is { } yAxis ? BuildChartAxis(yAxis) : new SpreadsheetChartAxisArtifact();
        }
        var rawSecondaryXAxis = Property(raw, "secondaryXAxis");
        var rawSecondaryYAxis = Property(raw, "secondaryYAxis");
        if (rawSecondaryXAxis is not null || rawSecondaryYAxis is not null)
        {
            if (chart.Type != SpreadsheetChartType.Combo)
                throw Unsupported(element.Id, "secondary axes require a combo chart");
            chart.SecondaryXAxis = rawSecondaryXAxis is { } xAxis
                ? BuildChartAxis(xAxis)
                : new SpreadsheetChartAxisArtifact();
            chart.SecondaryYAxis = rawSecondaryYAxis is { } yAxis
                ? BuildChartAxis(yAxis)
                : new SpreadsheetChartAxisArtifact();
        }

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        for (var index = 0; index < element.Data.Series.Count; index++)
        {
            var source = element.Data.Series[index];
            if (element.ChartType != "combo") RejectProperties(seriesJson[index], element.Id, "chartType", "axis");
            var effectiveType = ChartType(source.ChartType ?? element.ChartType);
            var series = BuildSeries(source, seriesJson[index], catalog, effectiveType);
            if (chart.Type == SpreadsheetChartType.Combo)
            {
                if (source.ChartType is "bar" or "column")
                {
                    var direction = source.ChartType == "bar" ? "bar" : "column";
                    if (chart.BarDirection.Length > 0 && chart.BarDirection != direction)
                        throw Unsupported(element.Id, "combo charts cannot mix horizontal bars and vertical columns");
                    chart.BarDirection = direction;
                }
                chart.ComboSeries.Add(new PresentationComboSeriesArtifact
                {
                    Type = ChartType(source.ChartType!),
                    AxisGroup = source.Axis == "secondary"
                        ? PresentationChartAxisGroup.Secondary
                        : PresentationChartAxisGroup.Primary,
                    Series = series,
                });
            }
            else chart.Series.Add(series);
        }
        ApplyAccessibility(chart, element.Accessibility);
        return chart;
    }

    private static SpreadsheetChartSeriesArtifact BuildSeries(
        PpjChartSeriesModel source,
        JsonElement raw,
        Catalog catalog,
        SpreadsheetChartType chartType)
    {
        if (raw.TryGetProperty("fill", out _) && raw.TryGetProperty("color", out _))
            throw Unsupported(source.Id, "chart-series color and fill are aliases and cannot both be present");
        var series = new SpreadsheetChartSeriesArtifact { Name = source.Name };
        series.Values.Add(source.Values.Select(value => value!.Value));
        series.XValues.Add(source.XValues);
        series.BubbleSizes.Add(source.BubbleSizes);
        if (raw.TryGetProperty("fill", out var fill))
        {
            var fillColor = FillColor(fill, catalog) ??
                throw Unsupported(source.Id, "explicit no-fill chart series are not yet compiler-owned");
            if (fillColor.Opacity != 1)
                throw Unsupported(source.Id, "chart-series fill opacity is not yet compiler-owned");
            series.Fill = new SpreadsheetColor { Rgb = fillColor.Rgb };
        }
        else if (raw.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            if (resolved.Alpha != 1)
                throw Unsupported(source.Id, "chart-series color opacity is not yet compiler-owned");
            series.Fill = new SpreadsheetColor { Rgb = resolved.Rgb };
        }
        if (raw.TryGetProperty("stroke", out var stroke))
            series.Line = BuildChartLine(stroke, catalog);
        if (raw.TryGetProperty("marker", out var marker))
            series.Marker = BuildChartMarker(marker, catalog);
        if (raw.TryGetProperty("trendlines", out var trendlines))
        {
            if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
                throw Unsupported(source.Id, "trendlines require a bar, column, or line series");
            series.Trendlines.Add(trendlines.EnumerateArray().Select(item => BuildChartTrendline(item, catalog)));
        }
        if (raw.TryGetProperty("errorBars", out var errorBars))
        {
            if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
                throw Unsupported(source.Id, "error bars require a bar, column, or line series");
            series.ErrorBars = BuildChartErrorBars(errorBars, catalog);
        }
        return series;
    }

    private static SpreadsheetChartAxisArtifact BuildChartAxis(JsonElement source)
    {
        var axis = new SpreadsheetChartAxisArtifact
        {
            Title = OptionalString(source, "title") ?? string.Empty,
            NumberFormatCode = OptionalString(source, "numberFormat") ?? string.Empty,
        };
        if (source.TryGetProperty("tickLabelInterval", out var tickLabelInterval))
            axis.TickLabelInterval = checked((uint)tickLabelInterval.GetInt32());
        if (source.TryGetProperty("min", out var minimum)) axis.Minimum = minimum.GetDouble();
        if (source.TryGetProperty("max", out var maximum)) axis.Maximum = maximum.GetDouble();
        if (source.TryGetProperty("majorUnit", out var majorUnit)) axis.MajorUnit = majorUnit.GetDouble();
        if (source.TryGetProperty("visible", out var visible)) axis.Visible = visible.GetBoolean();
        if (source.TryGetProperty("textStyle", out var textStyle))
        {
            axis.TextStyle = new SpreadsheetChartTextStyleArtifact();
            if (textStyle.TryGetProperty("fontSize", out var fontSize))
                axis.TextStyle.FontSizePoints = fontSize.GetDouble();
        }
        return axis;
    }

    private static SpreadsheetChartMarkerArtifact BuildChartMarker(JsonElement source, Catalog catalog)
    {
        if (source.ValueKind == JsonValueKind.String)
            return new SpreadsheetChartMarkerArtifact { Symbol = Marker(source.GetString()!) };
        var marker = new SpreadsheetChartMarkerArtifact
        {
            Symbol = Marker(OptionalString(source, "symbol") ?? "circle"),
        };
        if (source.TryGetProperty("size", out var size)) marker.Size = checked((uint)size.GetInt32());
        if (source.TryGetProperty("fill", out var fill))
        {
            var color = catalog.Color(fill);
            if (color.Alpha != 1) throw Unsupported("chart marker", "fill opacity is not compiler-owned");
            marker.Fill = new SpreadsheetColor { Rgb = color.Rgb };
        }
        if (source.TryGetProperty("stroke", out var stroke)) marker.Line = BuildChartLine(stroke, catalog);
        return marker;
    }

    private static SpreadsheetChartTrendlineArtifact BuildChartTrendline(JsonElement source, Catalog catalog)
    {
        var trendline = new SpreadsheetChartTrendlineArtifact
        {
            Type = TrendlineType(source.GetProperty("type").GetString()!),
            Name = OptionalString(source, "name") ?? string.Empty,
            DisplayEquation = source.TryGetProperty("displayEquation", out var equation) && equation.GetBoolean(),
            DisplayRSquared = source.TryGetProperty("displayRSquared", out var rSquared) && rSquared.GetBoolean(),
        };
        if (source.TryGetProperty("order", out var order)) trendline.PolynomialOrder = checked((uint)order.GetInt32());
        if (source.TryGetProperty("period", out var period)) trendline.Period = checked((uint)period.GetInt32());
        if (source.TryGetProperty("forward", out var forward)) trendline.Forward = forward.GetDouble();
        if (source.TryGetProperty("backward", out var backward)) trendline.Backward = backward.GetDouble();
        if (source.TryGetProperty("intercept", out var intercept)) trendline.Intercept = intercept.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke)) trendline.Line = BuildChartLine(stroke, catalog);
        return trendline;
    }

    private static SpreadsheetChartErrorBarsArtifact BuildChartErrorBars(JsonElement source, Catalog catalog)
    {
        var errorBars = new SpreadsheetChartErrorBarsArtifact
        {
            Direction = ErrorBarDirection(OptionalString(source, "direction") ?? "y"),
            Type = ErrorBarType(OptionalString(source, "type") ?? "both"),
            ValueType = ErrorBarValueType(source.GetProperty("valueType").GetString()!),
            NoEndCap = source.TryGetProperty("noEndCap", out var noEndCap) && noEndCap.GetBoolean(),
        };
        if (source.TryGetProperty("value", out var value)) errorBars.Value = value.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke)) errorBars.Line = BuildChartLine(stroke, catalog);
        return errorBars;
    }

    private static SpreadsheetChartLineStyleArtifact BuildChartLine(JsonElement source, Catalog catalog)
    {
        var color = catalog.Color(source.GetProperty("color"));
        var line = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = color.Rgb },
            DashStyle = ChartDash(OptionalString(source, "dash")),
            Cap = OptionalString(source, "cap") ?? string.Empty,
            Join = OptionalString(source, "join") ?? string.Empty,
        };
        if (source.TryGetProperty("width", out var width)) line.WidthPoints = width.GetDouble();
        var opacity = OptionalDouble(source, "opacity") ?? color.Alpha;
        if (opacity < 1) line.OpacityThousandthPercent = Opacity(opacity);
        return line;
    }

    private static PresentationTable BuildTable(PpjTableElementModel element, JsonElement raw, Catalog catalog)
    {
        RejectUnsupportedFrameTransform(element.Id, element.Frame, "table");
        var table = new PresentationTable
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
        };
        var namedStyle = catalog.TableStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        var defaultTextStyle = FirstProperty(inlineStyle, namedStyle, "defaultTextStyle");
        var defaultCellFill = FirstProperty(inlineStyle, namedStyle, "defaultCellFill");
        table.ColumnWidthsEmu.Add(ScaledExtents(
            element.Columns.Select(column => column.Width).ToArray(),
            table.WidthEmu));
        var defaultHeight = element.Frame.Height / element.Rows.Count;
        var rowHeights = ScaledExtents(
            element.Rows.Select(row => row.Height ?? defaultHeight).ToArray(),
            table.HeightEmu);
        var occupied = new bool[element.Rows.Count, element.Columns.Count];
        for (var rowIndex = 0; rowIndex < element.Rows.Count; rowIndex++)
        {
            var source = element.Rows[rowIndex];
            var row = new PresentationTableRow { HeightEmu = rowHeights[rowIndex] };
            for (var columnIndex = 0; columnIndex < element.Columns.Count; columnIndex++)
                row.Cells.Add(new PresentationTableCell());
            var cursor = 0;
            for (var sourceCellIndex = 0; sourceCellIndex < source.Cells.Count; sourceCellIndex++)
            {
                var cell = source.Cells[sourceCellIndex];
                while (cursor < element.Columns.Count && occupied[rowIndex, cursor]) cursor++;
                if (cursor >= element.Columns.Count)
                    throw Unsupported(element.Id, "table span expansion exceeded the declared physical grid");
                var rawCell = raw.GetProperty("rows")[rowIndex].GetProperty("cells")[sourceCellIndex];
                var targetCell = row.Cells[cursor];
                targetCell.TextBody = BuildTextBody(
                    rawCell.GetProperty("text"),
                    defaultTextStyle,
                    Property(rawCell, "textStyle"),
                    catalog);
                targetCell.Text = PptxTextCodec.Flatten(targetCell.TextBody);
                if ((Property(rawCell, "fill") ?? defaultCellFill) is { } cellFill)
                    targetCell.Fill = BuildTableCellFill(cellFill, catalog, element.Id);
                if (Property(rawCell, "borders") is { } borders)
                    targetCell.Borders = BuildTableCellBorders(borders, catalog);
                if (cell.RowSpan > 1 || cell.ColumnSpan > 1)
                {
                    table.MergeRanges.Add(new PresentationTableMergeRange
                    {
                        StartRow = checked((uint)rowIndex),
                        EndRow = checked((uint)(rowIndex + cell.RowSpan - 1)),
                        StartColumn = checked((uint)cursor),
                        EndColumn = checked((uint)(cursor + cell.ColumnSpan - 1)),
                    });
                    for (var rowSpan = 0; rowSpan < cell.RowSpan; rowSpan++)
                        for (var columnSpan = 0; columnSpan < cell.ColumnSpan; columnSpan++)
                            occupied[rowIndex + rowSpan, cursor + columnSpan] = true;
                }
                else occupied[rowIndex, cursor] = true;
                cursor += cell.ColumnSpan;
            }
            table.Rows.Add(row);
        }
        if (FirstProperty(inlineStyle, namedStyle, "headerRows") is { } headerRows)
        {
            if (headerRows.GetInt32() > 1)
                throw Unsupported(element.Id, "native table authoring currently supports at most one header row");
            table.FirstRow = headerRows.GetInt32() == 1;
        }
        if (FirstProperty(inlineStyle, namedStyle, "bandedRows") is { } bandedRows)
            table.BandedRows = bandedRows.GetBoolean();
        if (FirstProperty(inlineStyle, namedStyle, "bandedColumns") is { } bandedColumns)
            table.BandedColumns = bandedColumns.GetBoolean();
        if (FirstProperty(inlineStyle, namedStyle, "firstColumnEmphasis") is { } firstColumn)
            table.FirstColumn = firstColumn.GetBoolean();
        if (FirstProperty(inlineStyle, namedStyle, "lastColumnEmphasis") is { } lastColumn)
            table.LastColumn = lastColumn.GetBoolean();
        if (defaultCellFill is { } fallbackFill)
        {
            var type = fallbackFill.GetProperty("type").GetString();
            if (type == "none") table.NoDefaultCellFill = true;
            else if (type == "solid")
            {
                var fill = FillColor(fallbackFill, catalog)!.Value;
                table.DefaultCellFillRgb = fill.Rgb;
            }
        }
        if (defaultTextStyle is { } textStyle)
            table.DefaultTextStyle = BuildTextStyle(Property(textStyle, "defaultText"), catalog);
        ApplyAccessibility(table, element.Accessibility);
        return table;
    }

    private static PresentationTableCellFill BuildTableCellFill(
        JsonElement fill,
        Catalog catalog,
        string elementId)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return new PresentationTableCellFill { NoFill = true };
        if (type == "solid")
        {
            var resolved = FillColor(fill, catalog) ??
                throw new InvalidOperationException("Solid PPJ table-cell fill unexpectedly resolved to none.");
            var output = new PresentationTableCellFill { SolidRgb = resolved.Rgb };
            if (resolved.Opacity < 1) output.OpacityThousandthPercent = Opacity(resolved.Opacity);
            return output;
        }
        if (type == "gradient")
            return new PresentationTableCellFill
            {
                GradientFill = BuildGradientFill(fill, color => catalog.Color(color)),
            };
        throw Unsupported(elementId, $"table-cell {type} fills are not compiler-owned");
    }

    private static PresentationTableCellBorders BuildTableCellBorders(JsonElement source, Catalog catalog)
    {
        var output = new PresentationTableCellBorders();
        if (source.TryGetProperty("left", out var left)) output.Left = BuildTableCellBorder(left, catalog);
        if (source.TryGetProperty("top", out var top)) output.Top = BuildTableCellBorder(top, catalog);
        if (source.TryGetProperty("right", out var right)) output.Right = BuildTableCellBorder(right, catalog);
        if (source.TryGetProperty("bottom", out var bottom)) output.Bottom = BuildTableCellBorder(bottom, catalog);
        return output;
    }

    private static SpreadsheetChartLineStyleArtifact BuildTableCellBorder(JsonElement stroke, Catalog catalog)
    {
        var color = catalog.Color(stroke.GetProperty("color"));
        var output = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = color.Rgb },
            WidthPoints = stroke.GetProperty("width").GetDouble(),
            DashStyle = ChartDash(OptionalString(stroke, "dash")),
            Cap = OptionalString(stroke, "cap") ?? string.Empty,
            Join = OptionalString(stroke, "join") ?? string.Empty,
        };
        var opacity = OptionalDouble(stroke, "opacity") ?? color.Alpha;
        if (opacity < 1) output.OpacityThousandthPercent = Opacity(opacity);
        return output;
    }

    private static PresentationConnector BuildConnector(PpjConnectorElementModel element, JsonElement raw, Catalog catalog)
    {
        RejectUnsupportedFrameTransform(element.Id, element.Frame, "connector");
        var connector = new PresentationConnector
        {
            ConnectorType = element.ConnectorType,
            StartXEmu = Emu(EndpointX(element.From, element.Frame, true)),
            StartYEmu = Emu(EndpointY(element.From, element.Frame, true)),
            EndXEmu = Emu(EndpointX(element.To, element.Frame, false)),
            EndYEmu = Emu(EndpointY(element.To, element.Frame, false)),
        };
        ApplyLine(connector, raw.GetProperty("stroke"), catalog);
        connector.StartArrow = Arrow(OptionalString(raw, "startArrow"));
        connector.EndArrow = Arrow(OptionalString(raw, "endArrow"));
        ApplyAccessibility(connector, element.Accessibility);
        return connector;
    }

    private static PresentationGroup BuildGroup(PpjGroupElementModel element, JsonElement raw, Catalog catalog)
    {
        RejectUnsupportedFrameTransform(element.Id, element.Frame, "group");
        var group = new PresentationGroup
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            ChildLeftEmu = Emu(element.Frame.X),
            ChildTopEmu = Emu(element.Frame.Y),
            ChildWidthEmu = Emu(element.Frame.Width),
            ChildHeightEmu = Emu(element.Frame.Height),
        };
        var childJson = raw.GetProperty("elements").EnumerateArray().ToArray();
        for (var index = 0; index < element.Elements.Count; index++)
            group.Children.Add(BuildElement(element.Elements[index], childJson[index], catalog));
        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static PresentationShape BuildPlaceholder(PpjPlaceholderElementModel element, JsonElement raw, Catalog catalog)
    {
        var shape = BuildTextShape(element, raw, catalog, "textbox");
        shape.Placeholder = new PresentationPlaceholderIdentity
        {
            Type = PlaceholderType(element.PlaceholderType),
            Index = raw.TryGetProperty("index", out var index)
                ? checked((uint)index.GetInt64())
                : StableIndex(element.Id),
            InheritsGeometry = false,
        };
        shape.DirectFrame = new PresentationPlaceholderFrame
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
        };
        return shape;
    }

    private static PresentationShape ShapeFrame(PpjFrameModel frame, string geometry) => new()
    {
        Geometry = geometry,
        LeftEmu = Emu(frame.X),
        TopEmu = Emu(frame.Y),
        WidthEmu = Emu(frame.Width),
        HeightEmu = Emu(frame.Height),
    };

    private static void ApplyShapeStyle(
        PresentationShape target,
        JsonElement? named,
        JsonElement? inline,
        Catalog catalog,
        string elementId,
        bool hasText)
    {
        var fill = FirstProperty(inline, named, "fill");
        if (fill is { } fillValue)
        {
            if (fillValue.GetProperty("type").GetString() == "gradient")
            {
                target.GradientFill = BuildGradientFill(fillValue, color => catalog.Color(color));
            }
            else
            {
                var resolved = FillColor(fillValue, catalog);
                if (resolved is not null)
                {
                    target.FillRgb = resolved.Value.Rgb;
                    if (resolved.Value.Opacity < 1)
                        target.FillOpacityThousandthPercent = Opacity(resolved.Value.Opacity);
                }
            }
        }
        var stroke = FirstProperty(inline, named, "stroke");
        if (stroke is { } strokeValue) ApplyLine(target, strokeValue, catalog);
        else target.LineStyle = "none";
        var shadow = FirstProperty(inline, named, "shadow");
        if (shadow is { } shadowValue)
        {
            var shadowColor = catalog.Color(shadowValue.GetProperty("color"));
            target.Shadow = new PresentationShadow
            {
                ColorRgb = shadowColor.Rgb,
                BlurRadiusEmu = Emu(shadowValue.GetProperty("blur").GetDouble()),
                DistanceEmu = Emu(shadowValue.GetProperty("distance").GetDouble()),
                DirectionAngle60000 = Angle(shadowValue.GetProperty("angle").GetDouble()),
                OpacityThousandthPercent = Opacity(OptionalDouble(shadowValue, "opacity") ?? shadowColor.Alpha),
            };
        }
        var overallOpacity = FirstProperty(inline, named, "opacity");
        if (overallOpacity is { } opacity)
        {
            if (target.FillRgb.Length == 0)
                throw Unsupported(elementId, "shape opacity without a solid fill cannot be represented losslessly");
            if (opacity.GetDouble() < 1 && (stroke is not null || shadow is not null || hasText))
                throw Unsupported(elementId, "shape opacity with stroke, shadow, or text requires per-branch alpha and is not yet compiler-owned");
            var fillOpacity = target.HasFillOpacityThousandthPercent
                ? target.FillOpacityThousandthPercent / 100_000d
                : 1d;
            target.FillOpacityThousandthPercent = Opacity(fillOpacity * opacity.GetDouble());
        }
    }

    private static void ApplyTextBoxFill(
        PresentationShape target,
        JsonElement fill,
        Catalog catalog)
    {
        if (fill.GetProperty("type").GetString() == "gradient")
        {
            target.GradientFill = BuildGradientFill(fill, color => catalog.Color(color));
            return;
        }
        var resolved = FillColor(fill, catalog);
        if (resolved is null) return;
        target.FillRgb = resolved.Value.Rgb;
        if (resolved.Value.Opacity < 1)
            target.FillOpacityThousandthPercent = Opacity(resolved.Value.Opacity);
    }

    private static void ApplyLine(PresentationShape target, JsonElement stroke, Catalog catalog)
    {
        var color = catalog.Color(stroke.GetProperty("color"));
        target.LineRgb = color.Rgb;
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        var opacity = OptionalDouble(stroke, "opacity") ?? color.Alpha;
        if (opacity < 1) target.LineOpacityThousandthPercent = Opacity(opacity);
    }

    private static void ApplyLine(PresentationConnector target, JsonElement stroke, Catalog catalog)
    {
        var color = catalog.Color(stroke.GetProperty("color"));
        target.LineRgb = color.Rgb;
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        var opacity = OptionalDouble(stroke, "opacity") ?? color.Alpha;
        if (opacity < 1) target.LineOpacityThousandthPercent = Opacity(opacity);
    }

    private static PresentationTextBody BuildTextBody(
        JsonElement text,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var body = new PresentationTextBody();
        ApplyBodyProperties(body.BodyProperties = new PresentationTextBodyProperties(), namedStyle, inlineStyle);
        if (text.ValueKind == JsonValueKind.String)
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(inlineStyle, "paragraph"),
                null);
            paragraph.Runs.Add(BuildRun(text.GetString()!, namedStyle, inlineStyle, null, null, catalog));
            body.Paragraphs.Add(paragraph);
            return body;
        }
        foreach (var paragraphJson in text.GetProperty("paragraphs").EnumerateArray())
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(inlineStyle, "paragraph"),
                Property(paragraphJson, "style"));
            foreach (var run in paragraphJson.GetProperty("runs").EnumerateArray())
                paragraph.Runs.Add(BuildRun(
                    run.GetProperty("text").GetString()!,
                    namedStyle,
                    inlineStyle,
                    Property(run, "style"),
                    Property(run, "hyperlink"),
                    catalog));
            body.Paragraphs.Add(paragraph);
        }
        return body;
    }

    private static PresentationTextBody EmptyTextBody(JsonElement? namedStyle, JsonElement? inlineStyle)
    {
        var body = new PresentationTextBody();
        ApplyBodyProperties(body.BodyProperties = new PresentationTextBodyProperties(), namedStyle, inlineStyle);
        return body;
    }

    private static PresentationTextRun BuildRun(
        string text,
        JsonElement? namedBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        JsonElement? hyperlink,
        Catalog catalog)
    {
        var run = new PresentationTextRun { Text = text };
        var inlineDefault = FirstProperty(inlineBox, null, "defaultText");
        var namedDefault = FirstProperty(namedBox, null, "defaultText");
        var bold = FirstProperty(inlineRun, inlineDefault, namedDefault, "bold");
        var italic = FirstProperty(inlineRun, inlineDefault, namedDefault, "italic");
        var size = FirstProperty(inlineRun, inlineDefault, namedDefault, "size");
        var color = FirstProperty(inlineRun, inlineDefault, namedDefault, "color");
        var font = FirstProperty(inlineRun, inlineDefault, namedDefault, "font");
        var family = FirstProperty(inlineRun, inlineDefault, namedDefault, "fontFamily");
        var eastAsia = FirstProperty(inlineRun, inlineDefault, namedDefault, "fontFamilyEastAsia");
        var underline = FirstProperty(inlineRun, inlineDefault, namedDefault, "underline");
        var strike = FirstProperty(inlineRun, inlineDefault, namedDefault, "strike");
        var kerning = FirstProperty(inlineRun, inlineDefault, namedDefault, "kerning");
        var spacing = FirstProperty(inlineRun, inlineDefault, namedDefault, "letterSpacing");
        var baseline = FirstProperty(inlineRun, inlineDefault, namedDefault, "baseline");
        var capitalization = FirstProperty(inlineRun, inlineDefault, namedDefault, "capitalization");
        var language = FirstProperty(inlineRun, inlineDefault, namedDefault, "language");
        if (language is not null)
            throw Unsupported("text", "run language is valid PPJ but not yet compiler-owned");
        if (bold is { } boldValue) run.Bold = boldValue.GetBoolean();
        if (italic is { } italicValue) run.Italic = italicValue.GetBoolean();
        if (size is { } sizeValue) run.FontSizePoints = sizeValue.GetDouble();
        if (color is { } colorValue)
        {
            var resolved = catalog.Color(colorValue);
            if (resolved.Alpha != 1)
                throw Unsupported("text", "run color alpha is not yet compiler-owned");
            run.ColorRgb = resolved.Rgb;
        }
        if (family is { } familyValue) run.FontFamily = familyValue.GetString()!;
        else if (font is { } fontValue) run.FontFamily = catalog.Font(fontValue.GetString()!);
        if (eastAsia is { } eastAsiaValue) run.FontFamilyEastAsia = eastAsiaValue.GetString()!;
        else if (run.FontFamily.Length > 0) run.FontFamilyEastAsia = run.FontFamily;
        if (underline is { } underlineValue) run.Underline = NativeUnderline(underlineValue.GetString()!);
        if (strike is { } strikeValue) run.Strike = NativeStrike(strikeValue);
        if (kerning is { } kerningValue) run.FontKerningPoints = kerningValue.GetDouble();
        if (spacing is { } spacingValue) run.FontSpacingPoints = spacingValue.GetDouble();
        if (baseline is { } baselineValue) run.FontBaselinePercent = baselineValue.GetDouble();
        if (capitalization is { } capitalizationValue) run.FontCaps = capitalizationValue.GetString()!;
        if (hyperlink is { } link)
        {
            run.RunHyperlink = new PresentationRunHyperlink { Uri = link.GetProperty("uri").GetString()! };
            if (link.TryGetProperty("tooltip", out var tooltip)) run.RunHyperlink.Tooltip = tooltip.GetString()!;
        }
        return run;
    }

    private static PresentationTextStyle BuildTextStyle(JsonElement? style, Catalog catalog)
    {
        var output = new PresentationTextStyle();
        if (style is not { } value) return output;
        RejectProperties(value, "text", "language");
        if (value.TryGetProperty("bold", out var bold)) output.Bold = bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic)) output.Italic = italic.GetBoolean();
        if (value.TryGetProperty("size", out var size)) output.FontSizePoints = size.GetDouble();
        if (value.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            if (resolved.Alpha != 1)
                throw Unsupported("text", "text-style color alpha is not yet compiler-owned");
            output.ColorRgb = resolved.Rgb;
        }
        if (value.TryGetProperty("fontFamily", out var family)) output.FontFamily = family.GetString()!;
        else if (value.TryGetProperty("font", out var font)) output.FontFamily = catalog.Font(font.GetString()!);
        if (value.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) output.FontFamilyEastAsia = eastAsia.GetString()!;
        else if (output.HasFontFamily) output.FontFamilyEastAsia = output.FontFamily;
        if (value.TryGetProperty("underline", out var underline)) output.Underline = NativeUnderline(underline.GetString()!);
        if (value.TryGetProperty("strike", out var strike)) output.Strike = NativeStrike(strike);
        if (value.TryGetProperty("kerning", out var kerning)) output.FontKerningPoints = kerning.GetDouble();
        if (value.TryGetProperty("letterSpacing", out var spacing)) output.FontSpacingPoints = spacing.GetDouble();
        if (value.TryGetProperty("baseline", out var baseline)) output.FontBaselinePercent = baseline.GetDouble();
        if (value.TryGetProperty("capitalization", out var capitalization)) output.FontCaps = capitalization.GetString()!;
        return output;
    }

    private static void ApplyBodyProperties(
        PresentationTextBodyProperties target,
        JsonElement? named,
        JsonElement? inline)
    {
        var vertical = FirstProperty(inline, named, "verticalAlignment");
        if (vertical is { } verticalValue)
            target.VerticalAnchor = verticalValue.GetString() == "middle" ? "center" : verticalValue.GetString()!;
        var autoFit = FirstProperty(inline, named, "autoFit");
        if (autoFit is { } autoFitValue)
            target.AutoFitMode = autoFitValue.GetString() switch
            {
                "shrink-text" => "shrinkText",
                "resize-shape" => "resizeShape",
                _ => "none",
            };
        var wrap = FirstProperty(inline, named, "wrap");
        if (wrap is { } wrapValue) target.Wrap = wrapValue.GetString()!;
        var margins = FirstProperty(inline, named, "margins");
        if (margins is { } inset)
        {
            if (inset.TryGetProperty("left", out var left)) target.LeftInsetEmu = Emu(left.GetDouble());
            if (inset.TryGetProperty("top", out var top)) target.TopInsetEmu = Emu(top.GetDouble());
            if (inset.TryGetProperty("right", out var right)) target.RightInsetEmu = Emu(right.GetDouble());
            if (inset.TryGetProperty("bottom", out var bottom)) target.BottomInsetEmu = Emu(bottom.GetDouble());
        }
        var columns = FirstProperty(inline, named, "columns");
        if (columns is { } columnCount) target.Columns = checked((uint)columnCount.GetInt32());
        var gap = FirstProperty(inline, named, "columnGap");
        if (gap is { } columnGap) target.ColumnSpacingEmu = Emu(columnGap.GetDouble());
        var columnDirection = FirstProperty(inline, named, "columnDirection");
        if (columnDirection is { } direction) target.RightToLeftColumns = direction.GetString() == "right-to-left";
        var verticalText = FirstProperty(inline, named, "verticalText");
        if (verticalText is { } textMode) target.VerticalTextMode = textMode.GetString()!;
    }

    private static void ApplyParagraphStyle(
        PresentationTextParagraph target,
        JsonElement? named,
        JsonElement? inline,
        JsonElement? direct)
    {
        if (FirstProperty(direct, inline, named, "alignment") is { } alignment)
            target.Alignment = alignment.GetString()!;
        if (FirstProperty(direct, inline, named, "level") is { } level)
            target.Level = checked((uint)level.GetInt32());
        if (FirstProperty(direct, inline, named, "indent") is { } indent)
            target.MarginLeftEmu = Emu(indent.GetDouble());
        if (FirstProperty(direct, inline, named, "hanging") is { } hanging)
            target.IndentEmu = -Emu(hanging.GetDouble());
        if (FirstProperty(direct, inline, named, "spaceBefore") is { } before)
            target.SpaceBeforePoints = before.GetDouble();
        if (FirstProperty(direct, inline, named, "spaceBeforeMultiplier") is { } beforeMultiplier)
            target.SpaceBeforeMultiplier = beforeMultiplier.GetDouble();
        if (FirstProperty(direct, inline, named, "spaceAfter") is { } after)
            target.SpaceAfterPoints = after.GetDouble();
        if (FirstProperty(direct, inline, named, "spaceAfterMultiplier") is { } afterMultiplier)
            target.SpaceAfterMultiplier = afterMultiplier.GetDouble();
        if (FirstProperty(direct, inline, named, "lineSpacing") is { } spacing)
            target.LineSpacingPoints = spacing.GetDouble();
        if (FirstProperty(direct, inline, named, "lineSpacingMultiplier") is { } spacingMultiplier)
            target.LineSpacingMultiplier = spacingMultiplier.GetDouble();
        if (FirstProperty(direct, inline, named, "bullet") is { } bullet)
        {
            var kind = bullet.GetProperty("type").GetString();
            if (kind == "none") target.NoBullet = true;
            else if (kind == "character") target.BulletCharacter = bullet.GetProperty("character").GetString()!;
            else if (kind == "number")
            {
                target.AutoNumber = new PresentationAutoNumberBullet
                {
                    Scheme = bullet.TryGetProperty("scheme", out var scheme)
                        ? scheme.GetString()!
                        : NumberScheme(bullet.GetProperty("format").GetString()!),
                };
                if (bullet.TryGetProperty("startAt", out var startAt)) target.AutoNumber.StartAt = checked((uint)startAt.GetInt32());
            }
            if (bullet.TryGetProperty("fontFamily", out var bulletFont)) target.BulletFontFamily = bulletFont.GetString()!;
            if (bullet.TryGetProperty("color", out var bulletColor)) target.BulletColorRgb = NormalizeRgbToken(bulletColor);
            if (bullet.TryGetProperty("size", out var bulletSize)) target.BulletSizePoints = bulletSize.GetDouble();
            if (bullet.TryGetProperty("sizePercent", out var bulletSizePercent)) target.BulletSizePercent = bulletSizePercent.GetDouble();
        }
    }

    private static void ApplyTransform(PresentationShape target, PpjFrameModel frame)
    {
        if (frame.Rotation == 0 && !frame.FlipH && !frame.FlipV) return;
        target.Transform = new PresentationShapeTransform();
        if (frame.Rotation != 0) target.Transform.RotationAngle60000 = Angle(frame.Rotation);
        if (frame.FlipH) target.Transform.FlipHorizontal = true;
        if (frame.FlipV) target.Transform.FlipVertical = true;
    }

    private static void RejectUnsupportedFrameTransform(
        string elementId,
        PpjFrameModel frame,
        string elementKind)
    {
        if (frame.Rotation != 0 || frame.FlipH || frame.FlipV)
            throw Unsupported(elementId, $"{elementKind} frame rotation and flips are not yet compiler-owned");
    }

    private static PresentationBackground BuildBackground(JsonElement fill, Catalog catalog)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "solid")
        {
            var color = catalog.Color(fill.GetProperty("color"));
            var opacity = OptionalDouble(fill, "opacity") ?? color.Alpha;
            if (opacity != 1)
                throw Unsupported("background", "translucent solid slide backgrounds are not yet compiler-owned");
            return new PresentationBackground { Solid = true, ColorRgb = color.Rgb };
        }
        if (type == "gradient") return new PresentationBackground
        {
            GradientFill = BuildGradientFill(fill, color => catalog.Color(color)),
        };
        if (type == "image")
        {
            if (fill.GetProperty("fit").GetString() != "stretch")
                throw Unsupported("background", "native image backgrounds currently support stretch fit only");
            if (OptionalDouble(fill, "opacity") is { } opacity && opacity != 1)
                throw Unsupported("background", "translucent image backgrounds are not yet compiler-owned");
            return new PresentationBackground { ImageAssetId = catalog.NativeAssetId(fill.GetProperty("asset").GetString()!) };
        }
        if (type == "none") return new PresentationBackground();
        throw Unsupported("background", $"{type} slide backgrounds are not yet compiler-owned");
    }

    private static PresentationSpeakerNotes BuildNotes(JsonElement notes, Catalog catalog)
    {
        var body = BuildTextBody(notes, null, null, catalog);
        return new PresentationSpeakerNotes { TextBody = body, Text = Flatten(body) };
    }

    private static PresentationAnimation BuildAnimation(PpjAnimationModel source, IReadOnlyList<PpjElementModel> elements)
    {
        var target = Walk(elements).Single(element => element.Id == source.TargetId);
        var animation = new PresentationAnimation
        {
            Id = source.Id,
            TargetId = source.TargetId,
            TargetKind = target is PpjChartElementModel ? "chart" : target is PpjTextElementModel or PpjShapeElementModel or PpjPlaceholderElementModel ? "shape" : "element",
            Effect = source.Effect,
            Phase = source.Phase,
            Start = source.Start,
            Direction = source.Direction ?? string.Empty,
            TextBuild = source.TextBuild ?? string.Empty,
            ChartBuild = source.ChartBuild ?? string.Empty,
            DurationMs = checked((uint)source.DurationMs),
        };
        if (source.DelayMs > 0) animation.DelayMs = checked((uint)source.DelayMs);
        if (source.StaggerMs > 0) animation.StaggerMs = checked((uint)source.StaggerMs);
        if (source.AnimateChartBackground is { } animateChartBackground)
            animation.AnimateChartBackground = animateChartBackground;
        return animation;
    }

    private static void ApplyTransition(PresentationSlide slide, PpjPageModel page, PresentationSlide? previous)
    {
        if (page.Transition is null || page.Transition.Type == "none") return;
        if (page.Transition.Type == "morph")
        {
            if (previous is null || !page.Transition.FromPageId!.Equals(previous.Id, StringComparison.Ordinal))
                throw Unsupported(page.Id, "Morph must reference the immediately preceding page");
            var previousElements = WalkPresentation(previous.Elements).ToDictionary(element => element.Id, StringComparer.Ordinal);
            var currentElements = WalkPresentation(slide.Elements).ToDictionary(element => element.Id, StringComparer.Ordinal);
            foreach (var pair in page.Transition.MorphPairs)
            {
                if (!previousElements.TryGetValue(pair.FromElementId, out var from) ||
                    !currentElements.TryGetValue(pair.ToElementId, out var to))
                    throw Unsupported(page.Id, $"Morph pair {pair.Key} does not resolve across adjacent pages");
                from.Name = $"!!{pair.Key}";
                to.Name = $"!!{pair.Key}";
            }
            var morph = new PresentationMorph
            {
                FromSlideId = page.Transition.FromPageId!,
                DurationMs = checked((uint)page.Transition.DurationMs),
            };
            morph.Pairs.Add(page.Transition.MorphPairs.Select(pair => new PresentationMorphPair
            {
                Key = pair.Key,
                FromId = pair.FromElementId,
                ToId = pair.ToElementId,
            }));
            slide.Morph = morph;
            return;
        }
        slide.Transition = new PresentationTransition
        {
            Effect = page.Transition.Type,
            Speed = "medium",
            AdvanceOnClick = true,
        };
        if (page.Transition.DurationMs > 0) slide.Transition.DurationMs = checked((uint)page.Transition.DurationMs);
        if (page.Raw.GetProperty("transition").TryGetProperty("direction", out var direction))
            slide.Transition.Direction = direction.GetString()!;
    }

    private static void AddSections(PresentationArtifact target, PpjProgramModel program)
    {
        for (var index = 0; index < program.Sections.Count; index++)
        {
            var source = program.Sections[index];
            var section = new PresentationSectionArtifact
            {
                Id = source.Id,
                Name = source.Name,
                NativeId = DeterministicGuid(program.Meta.Id, source.Id),
            };
            section.SlideIds.Add(source.PageIds);
            target.Sections.Add(section);
        }
    }

    private static void AddCustomShows(PresentationArtifact target, PpjProgramModel program)
    {
        for (var index = 0; index < program.CustomShows.Count; index++)
        {
            var source = program.CustomShows[index];
            var show = new PresentationCustomShowArtifact
            {
                Id = source.Id,
                Name = source.Name,
                NativeId = checked((uint)(index + 1)),
            };
            show.SlideIds.Add(source.PageIds);
            target.CustomShows.Add(show);
        }
    }

    private static void AddComments(PresentationArtifact target, PpjProgramModel program)
    {
        if (!program.Root.TryGetProperty("comments", out var rawComments)) return;
        var pageById = target.Slides.ToDictionary(page => page.Id, StringComparer.Ordinal);
        foreach (var raw in rawComments.EnumerateArray())
        {
            var pageId = raw.GetProperty("page").GetString()!;
            var comment = new PresentationLegacyComment
            {
                Id = raw.GetProperty("id").GetString()!,
                Author = raw.GetProperty("author").GetString()!,
                Text = raw.GetProperty("text").GetString()!,
                CreatedAt = raw.TryGetProperty("createdAt", out var created) ? created.GetString()! : "1970-01-01T00:00:00Z",
            };
            if (raw.TryGetProperty("position", out var position))
            {
                comment.PositionXEmu = Emu(position.GetProperty("x").GetDouble());
                comment.PositionYEmu = Emu(position.GetProperty("y").GetDouble());
            }
            pageById[pageId].LegacyComments.Add(comment);
        }
    }

    private static IReadOnlyList<Asset> ValidateAssets(PpjProgramModel program, IEnumerable<Asset> supplied)
    {
        var requested = supplied.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        var declared = program.Assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        foreach (var id in requested.Keys.Where(id => !declared.ContainsKey(id)))
            throw new CodecException("ppj.asset.undeclared", $"PPJ compile received undeclared asset {id}.", "$.assets");
        var output = new List<Asset>(program.Assets.Count);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var declaration in program.Assets)
        {
            if (!requested.TryGetValue(declaration.Id, out var asset) || asset.Data.IsEmpty)
                throw new CodecException("ppj.asset.missing", $"PPJ asset {declaration.Id} has no supplied bytes.", "$.assets");
            var hash = Sha256(asset.Data.Span);
            if (!hash.Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) ||
                !hash.Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.hashMismatch", $"PPJ asset {declaration.Id} does not match its declared SHA-256.", "$.assets");
            if (!asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.mimeMismatch", $"PPJ asset {declaration.Id} does not match its declared MIME type.", "$.assets");
            var normalized = asset.Clone();
            normalized.FileName = declaration.Uri;
            normalized.Id = NativeAssetId(declaration.MimeType, hash);
            if (nativeIds.Add(normalized.Id)) output.Add(normalized);
        }
        return output;
    }

    private static void ApplyImageCrop(PresentationImage target, PpjImageElementModel element, JsonElement raw, Catalog catalog)
    {
        var crop = Property(raw, "crop");
        if (crop is { } explicitCrop)
        {
            target.Crop = new PresentationImageCrop
            {
                LeftThousandthPercent = CropValue(explicitCrop, "left"),
                TopThousandthPercent = CropValue(explicitCrop, "top"),
                RightThousandthPercent = CropValue(explicitCrop, "right"),
                BottomThousandthPercent = CropValue(explicitCrop, "bottom"),
            };
            return;
        }
        var fit = element.Fit ?? "stretch";
        if (fit is "stretch" or "none") return;
        if (fit == "tile")
            throw Unsupported(element.Id, "tiled image fills require a native tile transform compiler");
        var dimensions = catalog.AssetDimensions(element.AssetId);
        if (dimensions is null) throw Unsupported(element.Id, $"{fit} requires declared image dimensions");
        var sourceAspect = dimensions.Value.Width / dimensions.Value.Height;
        var frameAspect = element.Frame.Width / element.Frame.Height;
        if (fit == "cover")
        {
            var horizontal = sourceAspect > frameAspect ? (1 - frameAspect / sourceAspect) / 2 : 0;
            var vertical = sourceAspect < frameAspect ? (1 - sourceAspect / frameAspect) / 2 : 0;
            target.Crop = new PresentationImageCrop
            {
                LeftThousandthPercent = checked((int)Math.Round(horizontal * 100_000)),
                RightThousandthPercent = checked((int)Math.Round(horizontal * 100_000)),
                TopThousandthPercent = checked((int)Math.Round(vertical * 100_000)),
                BottomThousandthPercent = checked((int)Math.Round(vertical * 100_000)),
            };
        }
        else
        {
            var horizontal = sourceAspect < frameAspect ? (1 - sourceAspect / frameAspect) / 2 : 0;
            var vertical = sourceAspect > frameAspect ? (1 - frameAspect / sourceAspect) / 2 : 0;
            target.Crop = new PresentationImageCrop
            {
                LeftThousandthPercent = -checked((int)Math.Round(horizontal * 100_000)),
                RightThousandthPercent = -checked((int)Math.Round(horizontal * 100_000)),
                TopThousandthPercent = -checked((int)Math.Round(vertical * 100_000)),
                BottomThousandthPercent = -checked((int)Math.Round(vertical * 100_000)),
            };
        }
    }

    private static void ApplyChartStyle(
        PresentationChart chart,
        JsonElement? named,
        JsonElement? inline,
        Catalog catalog,
        string elementId)
    {
        var legend = FirstProperty(inline, named, "legend")?.GetString();
        chart.HasLegend = legend is not null and not "none";
        if (chart.HasLegend) chart.LegendPosition = legend!;
        if (FirstProperty(inline, named, "stacking") is { } stacking)
        {
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area or SpreadsheetChartType.Combo))
                throw Unsupported(elementId, "stacking applies only to bar, column, line, area, and combo charts");
            chart.Grouping = stacking.GetString()!;
        }
        if (FirstProperty(inline, named, "gapWidth") is { } gapWidth)
        {
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Combo))
                throw Unsupported(elementId, "gapWidth applies only to bar, column, and combo charts");
            chart.GapWidth = checked((uint)gapWidth.GetInt32());
        }
        var axisBearing = chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut);
        if (FirstProperty(inline, named, "showCategoryAxis") is { } showCategoryAxis)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have category axes");
            chart.ShowCategoryAxis = showCategoryAxis.GetBoolean();
        }
        if (FirstProperty(inline, named, "showValueAxis") is { } showValueAxis)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have value axes");
            chart.ShowValueAxis = showValueAxis.GetBoolean();
        }
        if (FirstProperty(inline, named, "showGridlines") is { } showGridlines)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have value-axis gridlines");
            // DrawingML represents "no major gridlines" by omitting the node.
            // Keep false canonical rather than inventing a presence-only wire
            // distinction that the package cannot round-trip.
            if (showGridlines.GetBoolean()) chart.ShowGridlines = true;
        }
        if (FirstProperty(inline, named, "chartAreaFill") is { } chartAreaFill)
            chart.ChartAreaFill = BuildChartSurfaceFill(chartAreaFill, catalog, $"{elementId} chart area");
        if (FirstProperty(inline, named, "plotAreaFill") is { } plotAreaFill)
            chart.PlotAreaFill = BuildChartSurfaceFill(plotAreaFill, catalog, $"{elementId} plot area");
        if (FirstProperty(inline, named, "titleTextStyle") is { } titleTextStyle)
        {
            if (chart.Title.Length == 0)
                throw Unsupported(elementId, "titleTextStyle requires a non-empty chart title");
            chart.TitleTextStyle = new SpreadsheetChartTextStyleArtifact();
            if (titleTextStyle.TryGetProperty("fontSize", out var fontSize))
                chart.TitleTextStyle.FontSizePoints = fontSize.GetDouble();
        }
        var smooth = FirstProperty(inline, named, "smooth");
        var varyColors = FirstProperty(inline, named, "varyColors");
        if (smooth is not null || varyColors is not null)
        {
            if (chart.Type != SpreadsheetChartType.Line)
                throw Unsupported(elementId, "smooth and varyColors apply only to line charts");
            if (smooth is not null || varyColors?.GetBoolean() == true)
            {
                chart.LineOptions = new SpreadsheetChartLineOptionsArtifact();
                if (smooth is { } explicitSmooth) chart.LineOptions.Smooth = explicitSmooth.GetBoolean();
                if (varyColors?.GetBoolean() == true) chart.LineOptions.VaryColors = true;
            }
        }
        var structuredLabels = FirstProperty(inline, named, "dataLabels");
        var labels = FirstProperty(inline, named, "showDataLabels");
        var legacyPosition = FirstProperty(inline, named, "dataLabelPosition");
        if (structuredLabels is not null && (labels is not null || legacyPosition is not null))
            throw Unsupported(elementId, "structured dataLabels cannot be combined with legacy data-label fields");
        if (structuredLabels is { } dataLabels)
        {
            chart.DataLabels = new SpreadsheetChartDataLabelsArtifact
            {
                ShowValue = dataLabels.TryGetProperty("showValue", out var showValue) && showValue.GetBoolean(),
                ShowCategoryName = dataLabels.TryGetProperty("showCategory", out var showCategory) && showCategory.GetBoolean(),
            };
            if (dataLabels.TryGetProperty("showSeries", out var showSeries)) chart.DataLabels.ShowSeriesName = showSeries.GetBoolean();
            if (dataLabels.TryGetProperty("showPercent", out var showPercent)) chart.DataLabels.ShowPercent = showPercent.GetBoolean();
            if (dataLabels.TryGetProperty("position", out var position)) chart.DataLabels.Position = LabelPosition(position.GetString()!);
        }
        else if (labels is { } showLabels && showLabels.GetBoolean())
        {
            chart.DataLabels = new SpreadsheetChartDataLabelsArtifact
            {
                ShowValue = true,
                ShowSeriesName = false,
            };
            var position = legacyPosition?.GetString();
            if (position is not null) chart.DataLabels.Position = LabelPosition(position);
        }
        else if (legacyPosition is not null)
            throw Unsupported("chart", "dataLabelPosition requires showDataLabels: true");
    }

    private static SpreadsheetChartSurfaceFill BuildChartSurfaceFill(JsonElement fill, Catalog catalog, string subject)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return new SpreadsheetChartSurfaceFill { NoFill = true };
        if (type != "solid") throw Unsupported(subject, "chart and plot areas currently support only none or solid fills");
        var resolved = FillColor(fill, catalog) ?? throw new InvalidOperationException("Solid PPJ fill unexpectedly resolved to none.");
        var output = new SpreadsheetChartSurfaceFill { SolidRgb = resolved.Rgb };
        if (resolved.Opacity < 1) output.OpacityThousandthPercent = Opacity(resolved.Opacity);
        return output;
    }

    private static void ApplyCustomGeometry(PresentationShape target, JsonElement geometry, string elementId)
    {
        if (!geometry.TryGetProperty("viewBox", out var viewBox) ||
            !geometry.TryGetProperty("paths", out var paths))
            throw Unsupported(elementId, "custom geometry has no compiler-owned path graph");
        var originX = viewBox.GetProperty("x").GetDouble();
        var originY = viewBox.GetProperty("y").GetDouble();
        var width = CustomPathCoordinate(viewBox.GetProperty("width").GetDouble());
        var height = CustomPathCoordinate(viewBox.GetProperty("height").GetDouble());
        foreach (var sourcePath in paths.EnumerateArray())
        {
            var path = new PresentationCustomGeometryPath { Width = width, Height = height };
            if (sourcePath.TryGetProperty("fill", out var fill))
                path.FillMode = fill.GetBoolean()
                    ? PresentationCustomGeometryPath.Types.FillMode.Normal
                    : PresentationCustomGeometryPath.Types.FillMode.None;
            if (sourcePath.TryGetProperty("stroke", out var stroke)) path.Stroke = stroke.GetBoolean();
            foreach (var sourceCommand in sourcePath.GetProperty("commands").EnumerateArray())
            {
                var command = new PresentationCustomGeometryCommand();
                switch (sourceCommand.GetProperty("op").GetString())
                {
                    case "moveTo":
                        command.MoveTo = CustomPoint(sourceCommand, originX, originY, "x", "y");
                        break;
                    case "lineTo":
                        command.LineTo = CustomPoint(sourceCommand, originX, originY, "x", "y");
                        break;
                    case "quadraticTo":
                        command.QuadraticBezierTo = new PresentationCustomGeometryQuadraticBezier
                        {
                            Control = CustomPoint(sourceCommand, originX, originY, "x1", "y1"),
                            End = CustomPoint(sourceCommand, originX, originY, "x", "y"),
                        };
                        break;
                    case "cubicTo":
                        command.CubicBezierTo = new PresentationCustomGeometryCubicBezier
                        {
                            Control1 = CustomPoint(sourceCommand, originX, originY, "x1", "y1"),
                            Control2 = CustomPoint(sourceCommand, originX, originY, "x2", "y2"),
                            End = CustomPoint(sourceCommand, originX, originY, "x", "y"),
                        };
                        break;
                    case "close":
                        command.Close = true;
                        break;
                    default:
                        throw Unsupported(elementId, "custom geometry contains a path operation outside the PPJ vocabulary");
                }
                path.Commands.Add(command);
            }
            target.CustomPaths.Add(path);
        }
    }

    private static PresentationCustomGeometryPoint CustomPoint(
        JsonElement command,
        double originX,
        double originY,
        string xName,
        string yName) => new()
    {
        X = CustomPathCoordinate(command.GetProperty(xName).GetDouble() - originX),
        Y = CustomPathCoordinate(command.GetProperty(yName).GetDouble() - originY),
    };

    private static long CustomPathCoordinate(double value) =>
        checked((long)Math.Round(value * CustomPathUnitsPerPoint, MidpointRounding.AwayFromZero));

    private static void ApplyAccessibility(PresentationShape target, PpjAccessibilityModel? source) =>
        target.Accessibility = Accessibility(source);

    private static void ApplyAccessibility(PresentationImage target, PpjAccessibilityModel? source)
    {
        if (source is null) return;
        if (source.Decorative) target.AccessibilityDecorative = true;
        else
        {
            if (source.Title is not null) target.AccessibilityTitle = source.Title;
            if (source.Description is not null) target.AltText = source.Description;
        }
    }

    private static void ApplyAccessibility(PresentationChart target, PpjAccessibilityModel? source) =>
        target.Accessibility = Accessibility(source);

    private static void ApplyAccessibility(PresentationTable target, PpjAccessibilityModel? source) =>
        target.Accessibility = Accessibility(source);

    private static void ApplyAccessibility(PresentationConnector target, PpjAccessibilityModel? source) =>
        target.Accessibility = Accessibility(source);

    private static void ApplyAccessibility(PresentationGroup target, PpjAccessibilityModel? source) =>
        target.Accessibility = Accessibility(source);

    private static PresentationNonVisualAccessibility? Accessibility(PpjAccessibilityModel? source)
    {
        if (source is null) return null;
        var output = new PresentationNonVisualAccessibility();
        if (source.Decorative) output.Decorative = true;
        else
        {
            if (source.Title is not null) output.Title = source.Title;
            if (source.Description is not null) output.Description = source.Description;
        }
        return output;
    }

    private static IEnumerable<PpjElementModel> Walk(IEnumerable<PpjElementModel> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element is PpjGroupElementModel group)
                foreach (var child in Walk(group.Elements)) yield return child;
        }
    }

    private static IEnumerable<PresentationElement> WalkPresentation(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in WalkPresentation(element.Group.Children)) yield return child;
        }
    }

    private static double EndpointX(PpjConnectorEndpointModel endpoint, PpjFrameModel frame, bool start) =>
        endpoint.X ?? (start ? frame.X : frame.X + frame.Width);

    private static double EndpointY(PpjConnectorEndpointModel endpoint, PpjFrameModel frame, bool start) =>
        endpoint.Y ?? frame.Y + frame.Height / 2;

    private static string PlaceholderType(string value) => value switch
    {
        "centerTitle" => "ctrTitle",
        "subtitle" => "subTitle",
        "content" => "body",
        _ => value,
    };

    private static uint StableIndex(string id)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(id));
        return Math.Max(1, BitConverter.ToUInt32(bytes, 0));
    }

    private static string DeterministicGuid(string programId, string id)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes($"{programId}\0{id}"));
        bytes[6] = (byte)((bytes[6] & 0x0f) | 0x50);
        bytes[8] = (byte)((bytes[8] & 0x3f) | 0x80);
        return new Guid(bytes.AsSpan(0, 16)).ToString("B", CultureInfo.InvariantCulture).ToUpperInvariant();
    }

    private static string CategoryText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString()!,
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        _ => string.Empty,
    };

    private static SpreadsheetChartType ChartType(string value) => value switch
    {
        "bar" or "column" => SpreadsheetChartType.Bar,
        "line" => SpreadsheetChartType.Line,
        "area" => SpreadsheetChartType.Area,
        "pie" => SpreadsheetChartType.Pie,
        "doughnut" => SpreadsheetChartType.Doughnut,
        "scatter" => SpreadsheetChartType.Scatter,
        "bubble" => SpreadsheetChartType.Bubble,
        "combo" => SpreadsheetChartType.Combo,
        _ => throw Unsupported("chart", $"chart type {value} is not compiler-owned"),
    };

    private static SpreadsheetChartLineDashStyle ChartDash(string? value) => value switch
    {
        null or "solid" => SpreadsheetChartLineDashStyle.Solid,
        "dash" or "long-dash" => SpreadsheetChartLineDashStyle.Dashed,
        "dot" => SpreadsheetChartLineDashStyle.Dotted,
        "dash-dot" => SpreadsheetChartLineDashStyle.DashDot,
        _ => SpreadsheetChartLineDashStyle.Unspecified,
    };

    private static SpreadsheetChartMarkerSymbol Marker(string value) => value switch
    {
        "none" => SpreadsheetChartMarkerSymbol.None,
        "dot" => SpreadsheetChartMarkerSymbol.Dot,
        "circle" => SpreadsheetChartMarkerSymbol.Circle,
        "square" => SpreadsheetChartMarkerSymbol.Square,
        "diamond" => SpreadsheetChartMarkerSymbol.Diamond,
        "triangle" => SpreadsheetChartMarkerSymbol.Triangle,
        "x" => SpreadsheetChartMarkerSymbol.X,
        "star" => SpreadsheetChartMarkerSymbol.Star,
        "plus" => SpreadsheetChartMarkerSymbol.Plus,
        "dash" => SpreadsheetChartMarkerSymbol.Dash,
        _ => SpreadsheetChartMarkerSymbol.Unspecified,
    };

    private static SpreadsheetChartTrendlineType TrendlineType(string value) => value switch
    {
        "exponential" => SpreadsheetChartTrendlineType.Exponential,
        "linear" => SpreadsheetChartTrendlineType.Linear,
        "logarithmic" => SpreadsheetChartTrendlineType.Logarithmic,
        "moving-average" => SpreadsheetChartTrendlineType.MovingAverage,
        "polynomial" => SpreadsheetChartTrendlineType.Polynomial,
        "power" => SpreadsheetChartTrendlineType.Power,
        _ => SpreadsheetChartTrendlineType.Unspecified,
    };

    private static SpreadsheetChartErrorBarDirection ErrorBarDirection(string value) => value switch
    {
        "x" => SpreadsheetChartErrorBarDirection.X,
        "y" => SpreadsheetChartErrorBarDirection.Y,
        _ => SpreadsheetChartErrorBarDirection.Unspecified,
    };

    private static SpreadsheetChartErrorBarType ErrorBarType(string value) => value switch
    {
        "both" => SpreadsheetChartErrorBarType.Both,
        "minus" => SpreadsheetChartErrorBarType.Minus,
        "plus" => SpreadsheetChartErrorBarType.Plus,
        _ => SpreadsheetChartErrorBarType.Unspecified,
    };

    private static SpreadsheetChartErrorBarValueType ErrorBarValueType(string value) => value switch
    {
        "fixed-value" => SpreadsheetChartErrorBarValueType.FixedValue,
        "percentage" => SpreadsheetChartErrorBarValueType.Percentage,
        "standard-deviation" => SpreadsheetChartErrorBarValueType.StandardDeviation,
        "standard-error" => SpreadsheetChartErrorBarValueType.StandardError,
        _ => SpreadsheetChartErrorBarValueType.Unspecified,
    };

    private static SpreadsheetChartDataLabelPosition LabelPosition(string value) => value switch
    {
        "best-fit" => SpreadsheetChartDataLabelPosition.BestFit,
        "bottom" => SpreadsheetChartDataLabelPosition.Bottom,
        "center" => SpreadsheetChartDataLabelPosition.Center,
        "inside-base" => SpreadsheetChartDataLabelPosition.InsideBase,
        "inside-end" => SpreadsheetChartDataLabelPosition.InsideEnd,
        "left" => SpreadsheetChartDataLabelPosition.Left,
        "outside-end" => SpreadsheetChartDataLabelPosition.OutsideEnd,
        "right" => SpreadsheetChartDataLabelPosition.Right,
        "top" => SpreadsheetChartDataLabelPosition.Top,
        "above" => SpreadsheetChartDataLabelPosition.Top,
        "below" => SpreadsheetChartDataLabelPosition.Bottom,
        _ => SpreadsheetChartDataLabelPosition.Unspecified,
    };

    private static string LineStyle(string? value) => value switch
    {
        null or "solid" => "solid",
        "dash" or "long-dash" => "dashed",
        "dot" => "dotted",
        "dash-dot" => "dash-dot",
        _ => "solid",
    };

    private static string Arrow(string? value) => value is null or "none" ? string.Empty : value;

    private static int CropValue(JsonElement crop, string property) =>
        crop.TryGetProperty(property, out var value) ? checked((int)Math.Round(value.GetDouble() * 100_000)) : 0;

    private static void RejectProperties(JsonElement value, string elementId, params string[] names)
    {
        if (value.ValueKind != JsonValueKind.Object) return;
        foreach (var name in names)
            if (value.TryGetProperty(name, out _))
                throw Unsupported(elementId, $"{name} is valid PPJ but not yet compiler-owned for this element");
    }

    private static string NativeUnderline(string value) => value switch
    {
        "single" => "sng",
        "double" => "dbl",
        _ => value,
    };

    private static string NativeStrike(JsonElement value) => value.ValueKind == JsonValueKind.True
        ? "sngStrike"
        : value.ValueKind == JsonValueKind.False
            ? "noStrike"
            : value.GetString()!;

    private static string NumberScheme(string value) => value switch
    {
        "decimal" => "arabicPeriod",
        "lower-alpha" => "alphaLcPeriod",
        "upper-alpha" => "alphaUcPeriod",
        "lower-roman" => "romanLcPeriod",
        "upper-roman" => "romanUcPeriod",
        _ => throw Unsupported("text", $"unsupported numbered bullet format {value}"),
    };

    private static string NormalizeRgbToken(JsonElement color)
    {
        if (color.ValueKind != JsonValueKind.String)
            throw Unsupported("text", "theme bullet colors require the theme-aware text compiler");
        var resolved = ParseHexColor(color.GetString()!);
        if (resolved.Alpha != 1)
            throw Unsupported("text", "bullet color alpha is not yet compiler-owned");
        return resolved.Rgb;
    }

    private static string Flatten(PpjTextContentModel text) =>
        text.PlainText ?? string.Join('\n', text.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    private static string Flatten(PresentationTextBody body) =>
        string.Join('\n', body.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    private static long Emu(double points) => checked((long)Math.Round(points * EmuPerPoint));

    private static IReadOnlyList<long> ScaledExtents(IReadOnlyList<double> values, long totalEmu)
    {
        var total = values.Sum();
        if (total <= 0) throw new CodecException("ppj.table.extent", "PPJ table row and column extents must be positive.");
        var output = new long[values.Count];
        long used = 0;
        for (var index = 0; index < values.Count; index++)
        {
            output[index] = index == values.Count - 1
                ? totalEmu - used
                : checked((long)Math.Round(totalEmu * values[index] / total));
            used += output[index];
        }
        return output;
    }
    private static int Angle(double degrees) => checked((int)Math.Round(degrees * 60_000));
    private static uint Opacity(double value) => checked((uint)Math.Round(value * 100_000));
    private static string Sha256(ReadOnlySpan<byte> value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();

    private static string DisplayName(string? name, string? role, string id)
    {
        var value = string.IsNullOrWhiteSpace(name) ? string.IsNullOrWhiteSpace(role) ? id : role : name;
        return value!.Length <= 255 ? value : value[..255];
    }

    private static string? OptionalString(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) ? property.GetString() : null;

    private static double? OptionalDouble(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) ? property.GetDouble() : null;

    private static JsonElement? Property(JsonElement value, string name) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) ? property : null;

    private static JsonElement? FirstProperty(JsonElement? first, JsonElement? second, string name) =>
        Property(first, name) ?? Property(second, name);

    private static JsonElement? FirstProperty(JsonElement? first, JsonElement? second, JsonElement? third, string name) =>
        Property(first, name) ?? Property(second, name) ?? Property(third, name);

    private static JsonElement? Property(JsonElement? value, string name) =>
        value is { ValueKind: JsonValueKind.Object } element && element.TryGetProperty(name, out var property) ? property : null;

    private static CodecException Unsupported(string owner, string message) =>
        new("unsupported_ppj_compile_feature", $"PPJ {owner}: {message}.");

    private sealed class Catalog
    {
        private readonly Dictionary<string, (string Rgb, double Alpha)> _colors;
        private readonly Dictionary<string, string> _fonts;
        private readonly Dictionary<string, JsonElement> _textStyles;
        private readonly Dictionary<string, JsonElement> _shapeStyles;
        private readonly Dictionary<string, JsonElement> _chartStyles;
        private readonly Dictionary<string, JsonElement> _tableStyles;
        private readonly Dictionary<string, (double Width, double Height)> _assetDimensions;
        private readonly Dictionary<string, string> _nativeAssetIds;

        internal PresentationThemeArtifact Theme { get; }

        internal Catalog(JsonElement root)
        {
            var design = root.GetProperty("design");
            _colors = design.GetProperty("theme").GetProperty("colors").EnumerateArray()
                .ToDictionary(
                    color => color.GetProperty("id").GetString()!,
                    color => ParseHexColor(color.GetProperty("value").GetString()!),
                    StringComparer.Ordinal);
            _fonts = design.GetProperty("fonts").EnumerateArray()
                .ToDictionary(
                    font => font.GetProperty("id").GetString()!,
                    font => font.GetProperty("family").GetString()!,
                    StringComparer.Ordinal);
            var styles = design.GetProperty("styles");
            _textStyles = Styles(styles, "text");
            _shapeStyles = Styles(styles, "shape");
            _chartStyles = Styles(styles, "chart");
            _tableStyles = Styles(styles, "table");
            _assetDimensions = root.TryGetProperty("assets", out var assets)
                ? assets.EnumerateArray()
                    .Where(asset => asset.TryGetProperty("widthPx", out _) && asset.TryGetProperty("heightPx", out _))
                    .ToDictionary(
                        asset => asset.GetProperty("id").GetString()!,
                        asset => (asset.GetProperty("widthPx").GetDouble(), asset.GetProperty("heightPx").GetDouble()),
                        StringComparer.Ordinal)
                : new Dictionary<string, (double Width, double Height)>(StringComparer.Ordinal);
            _nativeAssetIds = root.TryGetProperty("assets", out assets)
                ? assets.EnumerateArray().ToDictionary(
                    asset => asset.GetProperty("id").GetString()!,
                    asset => PpjAuthoredPresentationCompiler.NativeAssetId(
                        asset.GetProperty("mimeType").GetString()!,
                        asset.GetProperty("sha256").GetString()!),
                    StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
            Theme = BuildTheme(design);
        }

        private static PresentationThemeArtifact BuildTheme(JsonElement design)
        {
            var theme = design.GetProperty("theme");
            var output = new PresentationThemeArtifact();
            if (theme.TryGetProperty("name", out var name)) output.Name = name.GetString()!;
            output.AccentRgb.Add(theme.GetProperty("colors").EnumerateArray()
                .Take(6)
                .Select(color => ParseHexColor(color.GetProperty("value").GetString()!).Rgb));
            var fonts = design.GetProperty("fonts").EnumerateArray().ToArray();
            if (fonts.Length > 0)
            {
                output.MajorFontFamily = fonts[0].GetProperty("family").GetString()!;
                output.MinorFontFamily = fonts.Length > 1
                    ? fonts[1].GetProperty("family").GetString()!
                    : output.MajorFontFamily;
            }
            return output;
        }

        internal (string Rgb, double Alpha) Color(JsonElement color)
        {
            if (color.ValueKind == JsonValueKind.String) return ParseHexColor(color.GetString()!);
            var token = color.GetProperty("token").GetString()!;
            if (!_colors.TryGetValue(token, out var value))
                throw new CodecException("ppj.color.unknown", $"PPJ color token {token} is not declared.");
            return (value.Rgb, color.TryGetProperty("alpha", out var alpha) ? alpha.GetDouble() : value.Alpha);
        }

        internal string Font(string id) => _fonts.TryGetValue(id, out var value)
            ? value
            : throw new CodecException("ppj.font.unknown", $"PPJ font token {id} is not declared.");

        internal JsonElement? TextStyle(string? id) => Style(_textStyles, id);
        internal JsonElement? ShapeStyle(string? id) => Style(_shapeStyles, id);
        internal JsonElement? ChartStyle(string? id) => Style(_chartStyles, id);
        internal JsonElement? TableStyle(string? id) => Style(_tableStyles, id);
        internal (double Width, double Height)? AssetDimensions(string id) =>
            _assetDimensions.TryGetValue(id, out var dimensions) ? dimensions : null;
        internal string NativeAssetId(string id) => _nativeAssetIds.TryGetValue(id, out var nativeId)
            ? nativeId
            : throw new CodecException("ppj.asset.unknown", $"PPJ asset {id} is not declared.");

        private static Dictionary<string, JsonElement> Styles(JsonElement root, string kind) =>
            root.TryGetProperty(kind, out var styles)
                ? styles.EnumerateArray().ToDictionary(
                    style => style.GetProperty("id").GetString()!,
                    style => style.GetProperty("style").Clone(),
                    StringComparer.Ordinal)
                : new Dictionary<string, JsonElement>(StringComparer.Ordinal);

        private static JsonElement? Style(IReadOnlyDictionary<string, JsonElement> styles, string? id) =>
            id is not null && styles.TryGetValue(id, out var style) ? style : null;
    }

    private static (string Rgb, double Opacity)? FillColor(JsonElement fill, Catalog catalog)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return null;
        if (type != "solid") throw Unsupported("fill", $"{type} fills require the extended native fill compiler");
        var color = catalog.Color(fill.GetProperty("color"));
        var opacity = OptionalDouble(fill, "opacity") ?? color.Alpha;
        return (color.Rgb, opacity);
    }

    private static PresentationGradientFill BuildGradientFill(
        JsonElement fill,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor)
    {
        var kind = OptionalString(fill, "kind") ?? "linear";
        var output = new PresentationGradientFill
        {
            Kind = kind switch
            {
                "linear" => PresentationGradientFill.Types.Kind.Linear,
                "radial" => PresentationGradientFill.Types.Kind.Radial,
                _ => throw Unsupported("fill", $"unsupported gradient kind {kind}"),
            },
        };
        if (output.Kind == PresentationGradientFill.Types.Kind.Linear)
        {
            var degrees = OptionalDouble(fill, "angle") ?? 0;
            var normalized = ((degrees % 360) + 360) % 360;
            output.Angle60000 = Angle(normalized);
        }
        else if (fill.TryGetProperty("angle", out _))
        {
            throw Unsupported("fill", "radial gradients cannot define a linear angle");
        }
        foreach (var item in fill.GetProperty("stops").EnumerateArray())
        {
            var color = resolveColor(item.GetProperty("color"));
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = Opacity(item.GetProperty("offset").GetDouble()),
                ColorRgb = color.Rgb,
            };
            var alpha = OptionalDouble(item, "opacity") ?? color.Alpha;
            if (alpha < 1) stop.OpacityThousandthPercent = Opacity(alpha);
            output.Stops.Add(stop);
        }
        PptxGradientFillCodec.Validate(output, "PPJ gradient");
        return output;
    }

    private static (string Rgb, double Alpha) ParseHexColor(string value)
    {
        var normalized = value.TrimStart('#').ToUpperInvariant();
        return normalized.Length == 8
            ? (normalized[..6], Convert.ToByte(normalized[6..], 16) / 255d)
            : (normalized, 1d);
    }

    private static string NativeAssetId(string mimeType, string sha256) => mimeType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
        ? $"asset/presentation/picture-bullet/{sha256.ToLowerInvariant()}"
        : throw new CodecException("ppj.asset.unsupportedPurpose", $"PPJ authored presentation asset MIME {mimeType} does not have a native compiler purpose.");
}
