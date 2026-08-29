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
        var fileSha256 = Sha256(exported.File);
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
        return new(exported.File, receipt, exported.Diagnostics);
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
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id);
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
        if (element.Data.Series.Any(series => series.Values.Any(value => value is null)))
            throw Unsupported(element.Id, "null chart values require a missing-value-aware native chart cache");
        if (element.ChartType is "radar" or "waterfall")
            throw Unsupported(element.Id, $"{element.ChartType} chart authoring is not yet compiler-owned");

        var chart = new PresentationChart
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Type = ChartType(element.ChartType),
            Title = element.Title is null ? string.Empty : Flatten(element.Title),
        };
        chart.Categories.Add(element.Data.Categories.Select(CategoryText));
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyChartStyle(chart, namedStyle, inlineStyle);

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        for (var index = 0; index < element.Data.Series.Count; index++)
        {
            var source = element.Data.Series[index];
            var series = BuildSeries(source, seriesJson[index], catalog);
            if (chart.Type == SpreadsheetChartType.Combo)
            {
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

    private static SpreadsheetChartSeriesArtifact BuildSeries(PpjChartSeriesModel source, JsonElement raw, Catalog catalog)
    {
        var series = new SpreadsheetChartSeriesArtifact { Name = source.Name };
        series.Values.Add(source.Values.Select(value => value!.Value));
        if (raw.TryGetProperty("fill", out var fill) && FillColor(fill, catalog) is { } fillColor)
            series.Fill = new SpreadsheetColor { Rgb = fillColor.Rgb };
        else if (raw.TryGetProperty("color", out var color))
            series.Fill = new SpreadsheetColor { Rgb = catalog.Color(color).Rgb };
        if (raw.TryGetProperty("stroke", out var stroke))
        {
            series.Line = new SpreadsheetChartLineStyleArtifact
            {
                Color = new SpreadsheetColor { Rgb = catalog.Color(stroke.GetProperty("color")).Rgb },
                DashStyle = ChartDash(OptionalString(stroke, "dash")),
            };
            if (stroke.TryGetProperty("width", out var width)) series.Line.WidthPoints = width.GetDouble();
        }
        if (raw.TryGetProperty("marker", out var marker))
            series.Marker = new SpreadsheetChartMarkerArtifact { Symbol = Marker(marker.GetString()!) };
        return series;
    }

    private static PresentationTable BuildTable(PpjTableElementModel element, JsonElement raw, Catalog catalog)
    {
        var table = new PresentationTable
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
        };
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
            foreach (var cell in source.Cells)
            {
                while (cursor < element.Columns.Count && occupied[rowIndex, cursor]) cursor++;
                if (cursor >= element.Columns.Count)
                    throw Unsupported(element.Id, "table span expansion exceeded the declared physical grid");
                row.Cells[cursor].Text = Flatten(cell.Text);
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
        var style = Property(raw, "style") ?? catalog.TableStyle(element.StyleRef);
        if (style is { } tableStyle)
        {
            if (tableStyle.TryGetProperty("headerRows", out var headerRows))
            {
                if (headerRows.GetInt32() > 1)
                    throw Unsupported(element.Id, "native table authoring currently supports at most one header row");
                table.FirstRow = headerRows.GetInt32() == 1;
            }
            if (tableStyle.TryGetProperty("bandedRows", out var bandedRows)) table.BandedRows = bandedRows.GetBoolean();
            if (tableStyle.TryGetProperty("bandedColumns", out var bandedColumns)) table.BandedColumns = bandedColumns.GetBoolean();
            if (tableStyle.TryGetProperty("firstColumnEmphasis", out var firstColumn)) table.FirstColumn = firstColumn.GetBoolean();
            if (tableStyle.TryGetProperty("lastColumnEmphasis", out var lastColumn)) table.LastColumn = lastColumn.GetBoolean();
            if (tableStyle.TryGetProperty("defaultCellFill", out var cellFill))
            {
                var fill = FillColor(cellFill, catalog);
                if (fill is null) table.NoDefaultCellFill = true;
                else
                {
                    if (fill.Value.Opacity != 1)
                        throw Unsupported(element.Id, "table-cell fill opacity is not yet compiler-owned");
                    table.DefaultCellFillRgb = fill.Value.Rgb;
                }
            }
            if (tableStyle.TryGetProperty("defaultTextStyle", out var textStyle))
                table.DefaultTextStyle = BuildTextStyle(Property(textStyle, "defaultText"), catalog);
        }
        ApplyAccessibility(table, element.Accessibility);
        return table;
    }

    private static PresentationConnector BuildConnector(PpjConnectorElementModel element, JsonElement raw, Catalog catalog)
    {
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
            Index = StableIndex(element.Id),
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
        string elementId)
    {
        var fill = FirstProperty(inline, named, "fill");
        if (fill is { } fillValue)
        {
            var resolved = FillColor(fillValue, catalog);
            if (resolved is not null)
            {
                target.FillRgb = resolved.Value.Rgb;
                if (resolved.Value.Opacity < 1)
                    target.FillOpacityThousandthPercent = Opacity(resolved.Value.Opacity);
            }
        }
        var stroke = FirstProperty(inline, named, "stroke");
        if (stroke is { } strokeValue) ApplyLine(target, strokeValue, catalog);
        else target.LineStyle = "none";
        var shadow = FirstProperty(inline, named, "shadow");
        if (shadow is { } shadowValue)
        {
            target.Shadow = new PresentationShadow
            {
                ColorRgb = catalog.Color(shadowValue.GetProperty("color")).Rgb,
                BlurRadiusEmu = Emu(shadowValue.GetProperty("blur").GetDouble()),
                DistanceEmu = Emu(shadowValue.GetProperty("distance").GetDouble()),
                DirectionAngle60000 = Angle(shadowValue.GetProperty("angle").GetDouble()),
                OpacityThousandthPercent = Opacity(OptionalDouble(shadowValue, "opacity") ?? 1),
            };
        }
        var overallOpacity = FirstProperty(inline, named, "opacity");
        if (overallOpacity is { } opacity)
        {
            if (target.FillRgb.Length == 0)
                throw Unsupported(elementId, "shape opacity without a solid fill cannot be represented losslessly");
            target.FillOpacityThousandthPercent = Opacity(opacity.GetDouble());
        }
    }

    private static void ApplyLine(PresentationShape target, JsonElement stroke, Catalog catalog)
    {
        target.LineRgb = catalog.Color(stroke.GetProperty("color")).Rgb;
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        if (OptionalDouble(stroke, "opacity") is not null)
            throw Unsupported("stroke", "stroke opacity is not yet compiler-owned");
    }

    private static void ApplyLine(PresentationConnector target, JsonElement stroke, Catalog catalog)
    {
        target.LineRgb = catalog.Color(stroke.GetProperty("color")).Rgb;
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        if (OptionalDouble(stroke, "opacity") is not null)
            throw Unsupported("connector", "connector stroke opacity is not yet compiler-owned");
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
            paragraph.Runs.Add(BuildRun(text.GetString()!, namedStyle, inlineStyle, null, catalog));
            body.Paragraphs.Add(paragraph);
            return body;
        }
        foreach (var paragraphJson in text.GetProperty("paragraphs").EnumerateArray())
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(paragraph, Property(paragraphJson, "style") ?? FirstProperty(inlineStyle, namedStyle, "paragraph"));
            foreach (var run in paragraphJson.GetProperty("runs").EnumerateArray())
                paragraph.Runs.Add(BuildRun(run.GetProperty("text").GetString()!, namedStyle, inlineStyle, Property(run, "style"), catalog));
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
        if (bold is { } boldValue) run.Bold = boldValue.GetBoolean();
        if (italic is { } italicValue) run.Italic = italicValue.GetBoolean();
        if (size is { } sizeValue) run.FontSizePoints = sizeValue.GetDouble();
        if (color is { } colorValue) run.ColorRgb = catalog.Color(colorValue).Rgb;
        if (family is { } familyValue) run.FontFamily = familyValue.GetString()!;
        else if (font is { } fontValue) run.FontFamily = catalog.Font(fontValue.GetString()!);
        if (run.FontFamily.Length > 0) run.FontFamilyEastAsia = run.FontFamily;
        RejectTextEffects(inlineRun);
        return run;
    }

    private static PresentationTextStyle BuildTextStyle(JsonElement? style, Catalog catalog)
    {
        var output = new PresentationTextStyle();
        if (style is not { } value) return output;
        if (value.TryGetProperty("bold", out var bold)) output.Bold = bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic)) output.Italic = italic.GetBoolean();
        if (value.TryGetProperty("size", out var size)) output.FontSizePoints = size.GetDouble();
        if (value.TryGetProperty("color", out var color)) output.ColorRgb = catalog.Color(color).Rgb;
        if (value.TryGetProperty("fontFamily", out var family)) output.FontFamily = family.GetString()!;
        else if (value.TryGetProperty("font", out var font)) output.FontFamily = catalog.Font(font.GetString()!);
        if (output.HasFontFamily) output.FontFamilyEastAsia = output.FontFamily;
        RejectTextEffects(style);
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
    }

    private static void ApplyParagraphStyle(PresentationTextParagraph target, JsonElement? style)
    {
        if (style is not { } value) return;
        if (value.TryGetProperty("alignment", out var alignment)) target.Alignment = alignment.GetString()!;
        if (value.TryGetProperty("level", out var level)) target.Level = checked((uint)level.GetInt32());
        if (value.TryGetProperty("indent", out var indent)) target.MarginLeftEmu = Emu(indent.GetDouble());
        if (value.TryGetProperty("hanging", out var hanging)) target.IndentEmu = -Emu(hanging.GetDouble());
        if (value.TryGetProperty("spaceBefore", out var before)) target.SpaceBeforePoints = before.GetDouble();
        if (value.TryGetProperty("spaceAfter", out var after)) target.SpaceAfterPoints = after.GetDouble();
        if (value.TryGetProperty("lineSpacing", out var spacing)) target.LineSpacingPoints = spacing.GetDouble();
        if (value.TryGetProperty("bullet", out var bullet))
        {
            var kind = bullet.GetProperty("kind").GetString();
            if (kind == "none") target.NoBullet = true;
            else if (kind == "character") target.BulletCharacter = bullet.GetProperty("character").GetString()!;
            else throw Unsupported("text", "numbered and image PPJ bullets require their native bullet compiler");
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

    private static PresentationBackground BuildBackground(JsonElement fill, Catalog catalog)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "solid") return new PresentationBackground { Solid = true, ColorRgb = catalog.Color(fill.GetProperty("color")).Rgb };
        if (type == "image") return new PresentationBackground { ImageAssetId = catalog.NativeAssetId(fill.GetProperty("asset").GetString()!) };
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

    private static void ApplyChartStyle(PresentationChart chart, JsonElement? named, JsonElement? inline)
    {
        var legend = FirstProperty(inline, named, "legend")?.GetString();
        chart.HasLegend = legend is not null and not "none";
        var labels = FirstProperty(inline, named, "showDataLabels");
        if (labels is { } showLabels && showLabels.GetBoolean())
        {
            chart.DataLabels = new SpreadsheetChartDataLabelsArtifact { ShowValue = true };
            var position = FirstProperty(inline, named, "dataLabelPosition")?.GetString();
            if (position is not null) chart.DataLabels.Position = LabelPosition(position);
        }
    }

    private static void ApplyCustomGeometry(PresentationShape target, JsonElement geometry, string elementId)
    {
        if (!geometry.TryGetProperty("paths", out _))
            throw Unsupported(elementId, "custom geometry has no compiler-owned path graph");
        throw Unsupported(elementId, "custom PPJ path lowering is not yet implemented");
    }

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
        "circle" => SpreadsheetChartMarkerSymbol.Circle,
        "square" => SpreadsheetChartMarkerSymbol.Square,
        "diamond" => SpreadsheetChartMarkerSymbol.Diamond,
        "triangle" => SpreadsheetChartMarkerSymbol.Triangle,
        _ => SpreadsheetChartMarkerSymbol.Unspecified,
    };

    private static SpreadsheetChartDataLabelPosition LabelPosition(string value) => value switch
    {
        "center" => SpreadsheetChartDataLabelPosition.Center,
        "inside-end" => SpreadsheetChartDataLabelPosition.InsideEnd,
        "outside-end" => SpreadsheetChartDataLabelPosition.OutsideEnd,
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
        foreach (var name in names)
            if (value.TryGetProperty(name, out _))
                throw Unsupported(elementId, $"{name} is valid PPJ but not yet compiler-owned for this element");
    }

    private static void RejectTextEffects(JsonElement? value)
    {
        if (value is not { } style) return;
        if (style.TryGetProperty("underline", out var underline) && underline.GetString() != "none" ||
            style.TryGetProperty("strike", out var strike) && strike.GetBoolean() ||
            style.TryGetProperty("letterSpacing", out _) || style.TryGetProperty("baseline", out _))
            throw Unsupported("text", "underline, strike, letter spacing, and baseline require the extended native text compiler");
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
                    color => (NormalizeRgb(color.GetProperty("value").GetString()!), 1d),
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
                .Select(color => NormalizeRgb(color.GetProperty("value").GetString()!)));
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
            if (color.ValueKind == JsonValueKind.String) return (NormalizeRgb(color.GetString()!), 1);
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

    private static string NormalizeRgb(string value) => value.TrimStart('#').ToUpperInvariant();

    private static string NativeAssetId(string mimeType, string sha256) => mimeType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
        ? $"asset/presentation/picture-bullet/{sha256.ToLowerInvariant()}"
        : throw new CodecException("ppj.asset.unsupportedPurpose", $"PPJ authored presentation asset MIME {mimeType} does not have a native compiler purpose.");
}
