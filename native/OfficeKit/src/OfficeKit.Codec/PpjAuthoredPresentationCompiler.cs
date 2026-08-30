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

        AddMasterLayoutState(presentation, program, catalog);

        var expandedByPage = expansion.Pages.ToDictionary(page => page.Id, StringComparer.Ordinal);
        for (var pageIndex = 0; pageIndex < program.Pages.Count; pageIndex++)
        {
            var page = program.Pages[pageIndex];
            var expanded = expandedByPage[page.Id];
            var slide = new PresentationSlide
            {
                Id = page.Id,
                Name = DisplayName(page.Name, page.Role, page.Id),
                LayoutId = page.LayoutId ?? string.Empty,
            };
            if (page.Raw.TryGetProperty("hidden", out var hidden)) slide.Hidden = hidden.GetBoolean();
            if (page.Raw.TryGetProperty("background", out var background))
                slide.Background = BuildBackground(background, catalog, program.Design.Width, program.Design.Height);
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

    private static void AddMasterLayoutState(
        PresentationArtifact presentation,
        PpjProgramModel program,
        Catalog catalog)
    {
        foreach (var source in program.Design.Masters)
        {
            var master = new PresentationMaster
            {
                Id = source.Id,
                Name = source.Name,
                TextStyles = new PresentationMasterTextStyles(),
            };
            if (source.Background is { } background)
                master.Background = BuildBackground(background, catalog, program.Design.Width, program.Design.Height);
            master.TextStyles.TitleLevels.Add(source.TitleTextLevels.Select(level => BuildMasterTextLevel(level, catalog)));
            master.TextStyles.BodyLevels.Add(source.BodyTextLevels.Select(level => BuildMasterTextLevel(level, catalog)));
            master.TextStyles.OtherLevels.Add(source.OtherTextLevels.Select(level => BuildMasterTextLevel(level, catalog)));
            master.Placeholders.Add(source.Placeholders.Select(placeholder =>
                BuildLayoutPlaceholder(placeholder, catalog)));
            presentation.Masters.Add(master);
        }

        foreach (var source in program.Design.Layouts)
        {
            var layout = new PresentationLayout
            {
                Id = source.Id,
                Name = source.Name,
                MasterId = source.MasterId,
                Type = source.LayoutType,
            };
            if (source.Background is { } background)
                layout.Background = BuildBackground(background, catalog, program.Design.Width, program.Design.Height);
            layout.Placeholders.Add(source.Placeholders.Select(placeholder =>
                BuildLayoutPlaceholder(placeholder, catalog)));
            presentation.Layouts.Add(layout);
        }
    }

    private static PresentationTextParagraph BuildMasterTextLevel(JsonElement source, Catalog catalog)
    {
        var level = new PresentationTextParagraph();
        ApplyParagraphStyle(level, null, null, source, catalog);
        return level;
    }

    private static PresentationPlaceholder BuildLayoutPlaceholder(
        PpjLayoutPlaceholderModel source,
        Catalog catalog)
    {
        var placeholder = new PresentationPlaceholder
        {
            Id = source.Id,
            Name = source.Name,
            Type = PlaceholderType(source.PlaceholderType),
            Index = source.Index,
            DirectFrame = new PresentationPlaceholderFrame
            {
                LeftEmu = Emu(source.Frame.X),
                TopEmu = Emu(source.Frame.Y),
                WidthEmu = Emu(source.Frame.Width),
                HeightEmu = Emu(source.Frame.Height),
            },
        };
        if (source.Frame.Rotation != 0)
            placeholder.DirectFrame.RotationAngle60000 = Angle(source.Frame.Rotation);
        if (source.Frame.FlipH) placeholder.DirectFrame.FlipHorizontal = true;
        if (source.Frame.FlipV) placeholder.DirectFrame.FlipVertical = true;
        placeholder.TextBody = source.Text is null
            ? EmptyTextBody(null, source.Style)
            : BuildTextBody(source.Raw.GetProperty("text"), null, source.Style, catalog);
        return placeholder;
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
            case PpjChartElementModel { ChartType: "heatmap" } heatmap:
                output.Group = BuildHeatmap(heatmap, raw, catalog);
                break;
            case PpjChartElementModel { ChartType: "candlestick" } candlestick:
                output.Group = BuildCandlestick(candlestick, raw, catalog);
                break;
            case PpjChartElementModel { ChartType: "treemap" } treemap:
                output.Group = BuildTreemap(treemap, raw, catalog);
                break;
            case PpjChartElementModel { ChartType: "sunburst" } sunburst:
                output.Group = BuildSunburst(sunburst, raw, catalog);
                break;
            case PpjChartElementModel { ChartType: "sankey" } sankey:
                output.Group = BuildSankey(sankey, raw, catalog);
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
            case PpjSmartArtElementModel { Mode: "authored" } diagram:
                output.Group = BuildAuthoredDiagram(diagram, raw, catalog);
                break;
            case PpjSmartArtElementModel:
                throw Unsupported(element.Id, "source-bound SmartArt requires a bound source package");
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
        else shape.PresetAdjustments.Add(element.GeometryAdjustments);
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
            if (mask.GetProperty("kind").GetString() == "preset")
            {
                image.MaskPreset = mask.GetProperty("preset").GetString()!;
                image.MaskPresetAdjustments.Add(element.MaskAdjustments);
            }
            else
            {
                var customMask = new PresentationShape { Geometry = "custom" };
                ApplyCustomGeometry(customMask, mask, element.Id + " image mask");
                image.CustomMaskPaths.Add(customMask.CustomPaths);
            }
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
            image.Shadow = BuildShadow(shadow, catalog);
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
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "rich chart-title formatting is not yet compiler-owned");

        var isWaterfall = element.ChartType == "waterfall";
        var chart = new PresentationChart
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Type = isWaterfall ? SpreadsheetChartType.Bar : ChartType(element.ChartType),
            Title = element.Title is null ? string.Empty : Flatten(element.Title),
            BarDirection = element.ChartType == "bar" ? "bar" : element.ChartType is "column" or "waterfall" ? "column" : string.Empty,
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) chart.FrameTransform = frameTransform;
        chart.Categories.Add(element.Data.Categories.Select(CategoryText));
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        if (isWaterfall) ValidateWaterfallCompileProfile(element, raw, namedStyle, inlineStyle);
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
            chart.XAxis = rawXAxis is { } xAxis ? BuildChartAxis(xAxis, catalog) : new SpreadsheetChartAxisArtifact();
            chart.YAxis = rawYAxis is { } yAxis ? BuildChartAxis(yAxis, catalog) : new SpreadsheetChartAxisArtifact();
        }
        var rawSecondaryXAxis = Property(raw, "secondaryXAxis");
        var rawSecondaryYAxis = Property(raw, "secondaryYAxis");
        if (rawSecondaryXAxis is not null || rawSecondaryYAxis is not null)
        {
            if (chart.Type != SpreadsheetChartType.Combo)
                throw Unsupported(element.Id, "secondary axes require a combo chart");
            chart.SecondaryXAxis = rawSecondaryXAxis is { } xAxis
                ? BuildChartAxis(xAxis, catalog)
                : new SpreadsheetChartAxisArtifact();
            chart.SecondaryYAxis = rawSecondaryYAxis is { } yAxis
                ? BuildChartAxis(yAxis, catalog)
                : new SpreadsheetChartAxisArtifact();
        }

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        if (isWaterfall)
        {
            BuildWaterfallSeries(chart, element, namedStyle, inlineStyle, catalog);
            ApplyAccessibility(chart, element.Accessibility);
            return chart;
        }
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

    private static PresentationGroup BuildHeatmap(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "vector heatmap titles must use the bounded string form");

        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateHeatmapCompileProfile(element, raw, namedStyle, inlineStyle);
        var style = FirstProperty(inlineStyle, namedStyle, "heatmap")!.Value;
        var scale = OptionalString(style, "scale") ?? "linear";
        var colors = style.GetProperty("colors").EnumerateArray()
            .Select(color => HeatmapColor(catalog.Color(color)))
            .ToArray();
        var values = element.Data.Series.SelectMany(series => series.Values)
            .Where(value => value is not null)
            .Select(value => value!.Value)
            .ToArray();
        var inferredMinimum = values.Min();
        var inferredMaximum = values.Max();
        double minimum;
        double maximum;
        if (style.TryGetProperty("domain", out var domain))
        {
            minimum = domain[0].GetDouble();
            maximum = domain[1].GetDouble();
        }
        else if (scale == "diverging")
        {
            var extent = Math.Max(Math.Abs(inferredMinimum), Math.Abs(inferredMaximum));
            if (extent == 0) extent = 1;
            minimum = -extent;
            maximum = extent;
        }
        else if (inferredMinimum == inferredMaximum)
        {
            var extent = Math.Max(1, Math.Abs(inferredMinimum) * 0.05);
            minimum = inferredMinimum - extent;
            maximum = inferredMaximum + extent;
        }
        else
        {
            minimum = inferredMinimum;
            maximum = inferredMaximum;
        }
        var midpoint = OptionalDouble(style, "midpoint") ?? 0;

        var showValues = OptionalBoolean(style, "showValues") ?? false;
        var showColorBar = OptionalBoolean(style, "showColorBar") ?? true;
        var cellGap = OptionalDouble(style, "cellGap") ?? 1.5;
        var axisStyle = Property(style, "axisTextStyle");
        var valueStyle = Property(style, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.11, 20, 36);
        var longestRowLabel = element.Data.Series.Max(series => series.Name.Length);
        var maximumLabelWidth = Math.Max(52, Math.Min(150, width * 0.24));
        var leftLabelWidth = Math.Clamp(24 + longestRowLabel * 4.8, 52, maximumLabelWidth);
        var bottomLabelHeight = Math.Clamp(height * 0.09, 20, 36);
        var colorBarWidth = showColorBar ? Math.Clamp(width * 0.09, 46, 72) : 0;
        var gridX = x + leftLabelWidth;
        var gridY = y + titleHeight;
        var gridWidth = width - leftLabelWidth - colorBarWidth;
        var gridHeight = height - titleHeight - bottomLabelHeight;
        var columnCount = element.Data.Categories.Count;
        var rowCount = element.Data.Series.Count;
        var cellWidth = gridWidth / columnCount;
        var cellHeight = gridHeight / rowCount;
        if (cellWidth < 8 || cellHeight < 8 || cellGap >= cellWidth || cellGap >= cellHeight)
            throw Unsupported(element.Id, "vector heatmap frame is too small for its matrix and cell gap");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(element.Id, "title"),
                "heatmap title",
                x,
                y,
                width,
                titleHeight,
                titleText,
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(element.Id, $"row/{rowIndex}"),
                $"heatmap row {rowIndex + 1}",
                x,
                gridY + rowIndex * cellHeight,
                leftLabelWidth - 6,
                cellHeight,
                element.Data.Series[rowIndex].Name,
                axisStyle,
                catalog,
                Math.Clamp(cellHeight * 0.34, 6, 10),
                "right"));

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var value = element.Data.Series[rowIndex].Values[columnIndex];
                var cellX = gridX + columnIndex * cellWidth + cellGap / 2;
                var cellY = gridY + rowIndex * cellHeight + cellGap / 2;
                var cellShape = ShapeFrame(
                    new PpjFrameModel(cellX, cellY, cellWidth - cellGap, cellHeight - cellGap, 0, false, false),
                    "rect");
                cellShape.LineStyle = "none";
                (byte R, byte G, byte B, double Alpha)? cellColor = null;
                if (value is not null)
                {
                    cellColor = HeatmapInterpolate(value.Value, minimum, maximum, midpoint, scale, colors);
                    cellShape.FillRgb = HeatmapHex(cellColor.Value);
                    if (cellColor.Value.Alpha < 1) cellShape.FillOpacityThousandthPercent = Opacity(cellColor.Value.Alpha);
                }
                else if (style.TryGetProperty("missingFill", out var missingFill))
                {
                    cellColor = HeatmapColor(catalog.Color(missingFill));
                    cellShape.FillRgb = HeatmapHex(cellColor.Value);
                    if (cellColor.Value.Alpha < 1) cellShape.FillOpacityThousandthPercent = Opacity(cellColor.Value.Alpha);
                }
                if (style.TryGetProperty("cellStroke", out var cellStroke)) ApplyLine(cellShape, cellStroke, catalog);
                group.Children.Add(new PresentationElement
                {
                    Id = HeatmapNativeId(element.Id, $"cell/{rowIndex}/{columnIndex}"),
                    Name = $"heatmap cell {rowIndex + 1},{columnIndex + 1}",
                    Shape = cellShape,
                });

                if (showValues && value is not null)
                {
                    var contrast = HeatmapLuminance(cellColor!.Value) > 0.52 ? "111827" : "FFFFFF";
                    group.Children.Add(VectorChartTextElement(
                        HeatmapNativeId(element.Id, $"value/{rowIndex}/{columnIndex}"),
                        $"heatmap value {rowIndex + 1},{columnIndex + 1}",
                        cellX + 1,
                        cellY + 1,
                        cellWidth - cellGap - 2,
                        cellHeight - cellGap - 2,
                        value.Value.ToString("0.##", CultureInfo.InvariantCulture),
                        valueStyle,
                        catalog,
                        Math.Clamp(Math.Min(cellWidth, cellHeight) * 0.28, 6, 10),
                        "center",
                        contrast));
                }
            }
        }

        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(element.Id, $"column/{columnIndex}"),
                $"heatmap column {columnIndex + 1}",
                gridX + columnIndex * cellWidth,
                gridY + gridHeight + 3,
                cellWidth,
                bottomLabelHeight - 3,
                element.Data.Categories[columnIndex].GetString()!,
                axisStyle,
                catalog,
                Math.Clamp(bottomLabelHeight * 0.3, 6, 9),
                "center"));

        if (showColorBar)
            AddHeatmapColorBar(group, element.Id, gridX + gridWidth + 15, gridY, colorBarWidth - 15, gridHeight, minimum, maximum, midpoint, scale, colors, axisStyle, catalog);

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateHeatmapCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Categories.Count is < 1 or > 32 || element.Data.Series.Count is < 1 or > 32)
            throw Unsupported(element.Id, "vector heatmaps support a 1..32 by 1..32 matrix");
        if (element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != element.Data.Categories.Count)
            throw Unsupported(element.Id, "vector heatmap categories must be unique non-empty strings");
        if (element.Data.Series.Any(series => string.IsNullOrWhiteSpace(series.Name)) ||
            element.Data.Series.Select(series => series.Name).Distinct(StringComparer.Ordinal).Count() != element.Data.Series.Count)
            throw Unsupported(element.Id, "vector heatmap series names must be unique and non-empty");
        if (element.Data.Series.Any(series => series.Values.Count != element.Data.Categories.Count) ||
            element.Data.Series.SelectMany(series => series.Values).All(value => value is null))
            throw Unsupported(element.Id, "vector heatmap rows must match the category count and contain at least one numeric value");
        foreach (var series in element.Data.Series)
            foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
                if (series.Raw.TryGetProperty(property, out _))
                    throw Unsupported(element.Id, $"vector heatmap series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _)) throw Unsupported(element.Id, "vector heatmaps generate their own matrix labels");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"vector heatmaps do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "heatmap") is not { } heatmap)
            throw Unsupported(element.Id, "vector heatmaps require style.heatmap");
        var scale = OptionalString(heatmap, "scale") ?? "linear";
        var colorCount = heatmap.GetProperty("colors").GetArrayLength();
        if ((scale == "linear" && colorCount != 2) || (scale == "diverging" && colorCount != 3))
            throw Unsupported(element.Id, scale == "diverging" ? "diverging heatmaps require three colors" : "linear heatmaps require two colors");
        if (heatmap.TryGetProperty("domain", out var domain))
        {
            if (domain[0].GetDouble() >= domain[1].GetDouble())
                throw Unsupported(element.Id, "vector heatmap domain minimum must be smaller than its maximum");
            var effectiveMidpoint = OptionalDouble(heatmap, "midpoint") ?? 0;
            if (scale == "diverging" && (effectiveMidpoint <= domain[0].GetDouble() || effectiveMidpoint >= domain[1].GetDouble()))
                throw Unsupported(element.Id, "diverging heatmap midpoint must lie inside its explicit domain");
        }
        if (heatmap.TryGetProperty("midpoint", out var midpoint))
        {
            if (scale != "diverging") throw Unsupported(element.Id, "heatmap midpoint requires a diverging scale");
        }
    }

    private static PresentationElement VectorChartTextElement(
        string id,
        string name,
        double x,
        double y,
        double width,
        double height,
        string text,
        JsonElement? style,
        Catalog catalog,
        double defaultFontSize,
        string alignment,
        string? fallbackColor = null)
    {
        var shape = ShapeFrame(new PpjFrameModel(x, y, width, height, 0, false, false), "textbox");
        shape.LineStyle = "none";
        shape.Text = text;
        shape.TextBody = new PresentationTextBody
        {
            BodyProperties = new PresentationTextBodyProperties
            {
                NoLeftInset = true,
                NoTopInset = true,
                NoRightInset = true,
                NoBottomInset = true,
                VerticalAnchor = "center",
                Wrap = "none",
                AutoFitMode = "shrinkText",
            },
        };
        var paragraph = new PresentationTextParagraph { Alignment = alignment };
        paragraph.Runs.Add(BuildVectorChartRun(text, style, catalog, defaultFontSize, fallbackColor));
        shape.TextBody.Paragraphs.Add(paragraph);
        return new PresentationElement { Id = id, Name = name, Shape = shape };
    }

    private static PresentationTextRun BuildVectorChartRun(
        string text,
        JsonElement? style,
        Catalog catalog,
        double defaultFontSize,
        string? fallbackColor)
    {
        var run = new PresentationTextRun { Text = text, FontSizePoints = defaultFontSize };
        if (style is not { } value)
        {
            if (fallbackColor is not null) run.ColorRgb = fallbackColor;
            return run;
        }
        if (value.TryGetProperty("fontSize", out var fontSize)) run.FontSizePoints = fontSize.GetDouble();
        if (value.TryGetProperty("fontFamily", out var fontFamily)) run.FontFamily = fontFamily.GetString()!;
        if (value.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) run.FontFamilyEastAsia = eastAsia.GetString()!;
        else if (run.HasFontFamily) run.FontFamilyEastAsia = run.FontFamily;
        if (value.TryGetProperty("bold", out var bold)) run.Bold = bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic)) run.Italic = italic.GetBoolean();
        if (value.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            run.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) run.ColorOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else if (fallbackColor is not null) run.ColorRgb = fallbackColor;
        return run;
    }

    private static void AddHeatmapColorBar(
        PresentationGroup group,
        string elementId,
        double x,
        double y,
        double width,
        double height,
        double minimum,
        double maximum,
        double midpoint,
        string scale,
        IReadOnlyList<(byte R, byte G, byte B, double Alpha)> colors,
        JsonElement? axisStyle,
        Catalog catalog)
    {
        var barWidth = Math.Min(12, Math.Max(6, width * 0.32));
        var bar = ShapeFrame(new PpjFrameModel(x, y, barWidth, height, 0, false, false), "rect");
        bar.LineStyle = "none";
        bar.GradientFill = new PresentationGradientFill
        {
            Kind = PresentationGradientFill.Types.Kind.Linear,
            Angle60000 = Angle(90),
        };
        for (var index = 0; index < colors.Count; index++)
        {
            var color = colors[index];
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = colors.Count == 1 ? 0U : checked((uint)Math.Round(index * 100_000d / (colors.Count - 1))),
                ColorRgb = HeatmapHex(color),
            };
            if (color.Alpha < 1) stop.OpacityThousandthPercent = Opacity(color.Alpha);
            bar.GradientFill.Stops.Add(stop);
        }
        group.Children.Add(new PresentationElement { Id = HeatmapNativeId(elementId, "colorbar"), Name = "heatmap color scale", Shape = bar });
        var labelX = x + barWidth + 4;
        var labelWidth = Math.Max(20, width - barWidth - 4);
        group.Children.Add(VectorChartTextElement(
            HeatmapNativeId(elementId, "colorbar/max"),
            "heatmap scale maximum",
            labelX,
            y,
            labelWidth,
            14,
            maximum.ToString("0.##", CultureInfo.InvariantCulture),
            axisStyle,
            catalog,
            7,
            "left"));
        group.Children.Add(VectorChartTextElement(
            HeatmapNativeId(elementId, "colorbar/min"),
            "heatmap scale minimum",
            labelX,
            y + height - 14,
            labelWidth,
            14,
            minimum.ToString("0.##", CultureInfo.InvariantCulture),
            axisStyle,
            catalog,
            7,
            "left"));
        if (scale == "diverging")
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(elementId, "colorbar/mid"),
                "heatmap scale midpoint",
                labelX,
                y + height / 2 - 7,
                labelWidth,
                14,
                midpoint.ToString("0.##", CultureInfo.InvariantCulture),
                axisStyle,
                catalog,
                7,
                "left"));
    }

    private static (byte R, byte G, byte B, double Alpha) HeatmapColor((string Rgb, double Alpha) color) => (
        byte.Parse(color.Rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        byte.Parse(color.Rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        byte.Parse(color.Rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        color.Alpha);

    private static string HeatmapNativeId(string elementId, string suffix) =>
        $"{elementId}/heatmap/{suffix}";

    private static (byte R, byte G, byte B, double Alpha) HeatmapInterpolate(
        double value,
        double minimum,
        double maximum,
        double midpoint,
        string scale,
        IReadOnlyList<(byte R, byte G, byte B, double Alpha)> colors)
    {
        if (scale == "diverging")
        {
            if (value <= midpoint)
                return HeatmapMix(colors[0], colors[1], (Math.Clamp(value, minimum, midpoint) - minimum) / (midpoint - minimum));
            return HeatmapMix(colors[1], colors[2], (Math.Clamp(value, midpoint, maximum) - midpoint) / (maximum - midpoint));
        }
        return HeatmapMix(colors[0], colors[1], (Math.Clamp(value, minimum, maximum) - minimum) / (maximum - minimum));
    }

    private static (byte R, byte G, byte B, double Alpha) HeatmapMix(
        (byte R, byte G, byte B, double Alpha) from,
        (byte R, byte G, byte B, double Alpha) to,
        double amount) => (
            checked((byte)Math.Round(from.R + (to.R - from.R) * amount)),
            checked((byte)Math.Round(from.G + (to.G - from.G) * amount)),
            checked((byte)Math.Round(from.B + (to.B - from.B) * amount)),
            from.Alpha + (to.Alpha - from.Alpha) * amount);

    private static string HeatmapHex((byte R, byte G, byte B, double Alpha) color) =>
        $"{color.R:X2}{color.G:X2}{color.B:X2}";

    private static double HeatmapLuminance((byte R, byte G, byte B, double Alpha) color)
    {
        static double Channel(byte value)
        {
            var normalized = value / 255d;
            return normalized <= 0.03928 ? normalized / 12.92 : Math.Pow((normalized + 0.055) / 1.055, 2.4);
        }
        return 0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);
    }

    private static PresentationGroup BuildCandlestick(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "vector candlestick titles must use the bounded string form");

        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateCandlestickCompileProfile(element, raw, namedStyle, inlineStyle);
        var candlestick = FirstProperty(inlineStyle, namedStyle, "candlestick")!.Value;
        var series = element.Data.Series[0];
        var categories = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var closeValues = series.Values.Select(value => value!.Value).ToArray();
        var openValues = series.OpenValues;
        var highValues = series.HighValues;
        var lowValues = series.LowValues;
        var isOhlc = openValues.Count != 0;

        var xAxis = Property(raw, "xAxis");
        var yAxis = Property(raw, "yAxis");
        var showXAxis = xAxis is null || OptionalBoolean(xAxis.Value, "visible") != false;
        var showYAxis = yAxis is null || OptionalBoolean(yAxis.Value, "visible") != false;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");
        var axisStyle = Property(candlestick, "axisTextStyle");
        var valueStyle = Property(candlestick, "valueTextStyle");
        var xAxisTextStyle = Property(xAxis, "textStyle") ?? axisStyle;
        var yAxisTextStyle = Property(yAxis, "textStyle") ?? axisStyle;
        var xAxisTitleStyle = Property(xAxis, "titleTextStyle") ?? axisStyle;
        var yAxisTitleStyle = Property(yAxis, "titleTextStyle") ?? axisStyle;
        var showCloseValues = OptionalBoolean(candlestick, "showCloseValues") ?? false;
        if (showCloseValues && categories.Length > 16)
            throw Unsupported(element.Id, "showCloseValues is limited to 16 candlesticks to keep labels readable");

        var dataMinimum = lowValues.Min();
        var dataMaximum = highValues.Max();
        var dataRange = dataMaximum - dataMinimum;
        if (dataRange <= 0) dataRange = Math.Max(1, Math.Abs(dataMaximum) * 0.1);
        var majorUnit = yAxis is { } y && OptionalDouble(y, "majorUnit") is { } explicitMajor
            ? explicitMajor
            : NiceVectorChartStep(dataRange / 4);
        var pad = dataRange * 0.05;
        var minimum = yAxis is { } minimumAxis && OptionalDouble(minimumAxis, "min") is { } explicitMinimum
            ? explicitMinimum
            : Math.Floor((dataMinimum - pad) / majorUnit) * majorUnit;
        var maximum = yAxis is { } maximumAxis && OptionalDouble(maximumAxis, "max") is { } explicitMaximum
            ? explicitMaximum
            : Math.Ceiling((dataMaximum + pad) / majorUnit) * majorUnit;
        if (maximum <= minimum) throw Unsupported(element.Id, "candlestick yAxis maximum must be greater than its minimum");

        var tickCount = checked((int)Math.Floor((maximum - minimum) / majorUnit + 1e-9)) + 1;
        if (tickCount is < 2 or > 12)
            throw Unsupported(element.Id, "candlestick yAxis must expand to between 2 and 12 major ticks");

        var x = element.Frame.X;
        var yPosition = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var leftInset = showYAxis ? Math.Clamp(width * 0.1, 42, 68) : 8;
        var rightInset = showCloseValues ? Math.Clamp(width * 0.08, 32, 48) : 10;
        var bottomInset = showXAxis
            ? (xAxis is { } xAxisValue && xAxisValue.TryGetProperty("title", out _) ? 38 : 24)
            : 8;
        var topInset = titleHeight + 8;
        var plotX = x + leftInset;
        var plotY = yPosition + topInset;
        var plotWidth = width - leftInset - rightInset;
        var plotHeight = height - topInset - bottomInset;
        if (plotWidth < 120 || plotHeight < 80)
            throw Unsupported(element.Id, "candlestick frame is too small for native axes and editable marks");
        var slotWidth = plotWidth / categories.Length;
        if (slotWidth < 4.5)
            throw Unsupported(element.Id, "candlestick frame is too narrow for the requested observation count");

        var bodyWidthRatio = OptionalDouble(candlestick, "bodyWidthRatio") ?? 0.55;
        var bodyWidth = Math.Max(1.25, slotWidth * bodyWidthRatio);
        var labelInterval = xAxis is { } xAxisLabel && OptionalDouble(xAxisLabel, "tickLabelInterval") is { } interval
            ? checked((int)interval)
            : Math.Max(1, checked((int)Math.Ceiling(categories.Length / 12d)));
        var numberFormat = yAxis is { } formattedAxis ? OptionalString(formattedAxis, "numberFormat") : null;
        if (numberFormat is not null && !CandlestickNumberFormats.Contains(numberFormat, StringComparer.Ordinal))
            throw Unsupported(element.Id, $"candlestick yAxis numberFormat {numberFormat} is outside the bounded profile");

        double PlotY(double value) => plotY + (maximum - value) * plotHeight / (maximum - minimum);
        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(yPosition),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(yPosition),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTextElement(
                CandlestickNativeId(element.Id, "title"),
                "candlestick title",
                x,
                yPosition,
                width,
                titleHeight,
                titleText,
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        if (showYAxis)
        {
            var gridlineStroke = Property(candlestick, "gridlineStroke");
            for (var tickIndex = 0; tickIndex < tickCount; tickIndex++)
            {
                var tickValue = minimum + tickIndex * majorUnit;
                var tickY = PlotY(tickValue);
                if (gridlineStroke is { } gridline)
                    group.Children.Add(VectorChartLineElement(
                        CandlestickNativeId(element.Id, $"grid/{tickIndex}"),
                        $"candlestick gridline {tickIndex + 1}",
                        plotX,
                        tickY,
                        plotX + plotWidth,
                        tickY,
                        gridline,
                        catalog));
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"y-label/{tickIndex}"),
                    $"candlestick value-axis label {tickIndex + 1}",
                    x,
                    tickY - 7,
                    leftInset - 7,
                    14,
                    FormatCandlestickValue(tickValue, numberFormat),
                    yAxisTextStyle,
                    catalog,
                    8,
                    "right"));
            }
            if (yAxis is { } yAxisWithTitle && yAxisWithTitle.TryGetProperty("title", out var yTitle))
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, "y-title"),
                    "candlestick value-axis title",
                    x,
                    plotY - 18,
                    leftInset + Math.Min(120, plotWidth * 0.3),
                    16,
                    yTitle.GetString()!,
                    yAxisTitleStyle,
                    catalog,
                    8,
                    "left"));
        }

        var wickStyle = candlestick.GetProperty("wick");
        for (var index = 0; index < categories.Length; index++)
        {
            var centerX = plotX + (index + 0.5) * slotWidth;
            group.Children.Add(VectorChartLineElement(
                CandlestickNativeId(element.Id, $"wick/{index}"),
                $"candlestick wick {index + 1}",
                centerX,
                PlotY(highValues[index]),
                centerX,
                PlotY(lowValues[index]),
                wickStyle,
                catalog));

            if (isOhlc)
            {
                var open = openValues[index];
                var close = closeValues[index];
                var top = PlotY(Math.Max(open, close));
                var bottom = PlotY(Math.Min(open, close));
                var bodyHeight = Math.Max(1.5, bottom - top);
                if (bodyHeight > bottom - top) top = Math.Clamp((top + bottom - bodyHeight) / 2, plotY, plotY + plotHeight - bodyHeight);
                var role = close >= open ? "up" : "down";
                var bodyStyle = candlestick.GetProperty(role);
                var body = ShapeFrame(
                    new PpjFrameModel(centerX - bodyWidth / 2, top, bodyWidth, bodyHeight, 0, false, false),
                    "rect");
                ApplyTextBoxFill(body, bodyStyle.GetProperty("fill"), catalog);
                if (bodyStyle.TryGetProperty("stroke", out var bodyStroke)) ApplyLine(body, bodyStroke, catalog);
                else body.LineStyle = "none";
                group.Children.Add(new PresentationElement
                {
                    Id = CandlestickNativeId(element.Id, $"body/{index}"),
                    Name = $"candlestick {role} body {index + 1}",
                    Shape = body,
                });
            }
            else
            {
                var closeY = PlotY(closeValues[index]);
                group.Children.Add(VectorChartLineElement(
                    CandlestickNativeId(element.Id, $"close/{index}"),
                    $"candlestick close tick {index + 1}",
                    centerX,
                    closeY,
                    centerX + bodyWidth / 2,
                    closeY,
                    wickStyle,
                    catalog));
            }

            if (showXAxis && index % labelInterval == 0)
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"x-label/{index}"),
                    $"candlestick category label {index + 1}",
                    plotX + index * slotWidth,
                    plotY + plotHeight + 3,
                    Math.Max(slotWidth, 30),
                    15,
                    categories[index],
                    xAxisTextStyle,
                    catalog,
                    Math.Clamp(slotWidth * 0.24, 6, 9),
                    "center"));

            if (showCloseValues)
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"close-value/{index}"),
                    $"candlestick close value {index + 1}",
                    centerX - slotWidth / 2,
                    Math.Max(plotY, PlotY(highValues[index]) - 16),
                    slotWidth,
                    14,
                    FormatCandlestickValue(closeValues[index], numberFormat),
                    valueStyle,
                    catalog,
                    Math.Clamp(slotWidth * 0.22, 6, 8),
                    "center"));
        }

        if (showXAxis && xAxis is { } xAxisWithTitle && xAxisWithTitle.TryGetProperty("title", out var xTitle))
            group.Children.Add(VectorChartTextElement(
                CandlestickNativeId(element.Id, "x-title"),
                "candlestick category-axis title",
                plotX,
                yPosition + height - 17,
                plotWidth,
                16,
                xTitle.GetString()!,
                xAxisTitleStyle,
                catalog,
                8,
                "center"));

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateCandlestickCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Categories.Count is < 1 or > 64 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != element.Data.Categories.Count)
            throw Unsupported(element.Id, "candlestick categories must be 1..64 unique non-empty strings");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "candlestick charts require exactly one OHLC or HLC series");
        var series = element.Data.Series[0];
        var count = element.Data.Categories.Count;
        if (series.Values.Count != count || series.Values.Any(value => value is null) ||
            series.HighValues.Count != count || series.LowValues.Count != count ||
            (series.OpenValues.Count != 0 && series.OpenValues.Count != count))
            throw Unsupported(element.Id, "candlestick open/high/low/close channels must align with the category count");
        for (var index = 0; index < count; index++)
        {
            var low = series.LowValues[index];
            var high = series.HighValues[index];
            var close = series.Values[index]!.Value;
            if (low > high || close < low || close > high ||
                (series.OpenValues.Count != 0 && (series.OpenValues[index] < low || series.OpenValues[index] > high)))
                throw Unsupported(element.Id, "every open and close must lie inside its low/high interval");
        }
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"candlestick series do not support {property}");
        foreach (var property in new[] { "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "candlestick charts do not support secondary axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"candlestick charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "candlestick") is not { } candlestick)
            throw Unsupported(element.Id, "candlestick charts require style.candlestick");
        foreach (var role in new[] { "up", "down" })
        {
            var fillType = candlestick.GetProperty(role).GetProperty("fill").GetProperty("type").GetString();
            if (fillType is "none" or "image")
                throw Unsupported(element.Id, $"candlestick {role} fill must be solid or a bounded gradient");
        }
        if (raw.TryGetProperty("yAxis", out var yAxis))
        {
            var lowest = series.LowValues.Min();
            var highest = series.HighValues.Max();
            if (OptionalDouble(yAxis, "min") is { } minimum && minimum > lowest)
                throw Unsupported(element.Id, "candlestick yAxis.min must not clip the lowest observation");
            if (OptionalDouble(yAxis, "max") is { } maximum && maximum < highest)
                throw Unsupported(element.Id, "candlestick yAxis.max must not clip the highest observation");
        }
    }

    private static PresentationElement VectorChartLineElement(
        string id,
        string name,
        double startX,
        double startY,
        double endX,
        double endY,
        JsonElement stroke,
        Catalog catalog)
    {
        var connector = new PresentationConnector
        {
            ConnectorType = "straight",
            StartXEmu = Emu(startX),
            StartYEmu = Emu(startY),
            EndXEmu = Emu(endX),
            EndYEmu = Emu(endY),
        };
        ApplyLine(connector, stroke, catalog);
        return new PresentationElement { Id = id, Name = name, Connector = connector };
    }

    private static string CandlestickNativeId(string elementId, string suffix) =>
        $"{elementId}/candlestick/{suffix}";

    private static double NiceVectorChartStep(double raw)
    {
        if (!double.IsFinite(raw) || raw <= 0) return 1;
        var magnitude = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        var normalized = raw / magnitude;
        var nice = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
        return nice * magnitude;
    }

    private static readonly string[] CandlestickNumberFormats = ["0", "0.0", "0.00", "#,##0", "#,##0.0", "#,##0.00"];

    private static string FormatCandlestickValue(double value, string? numberFormat) =>
        value.ToString(numberFormat ?? "0.##", CultureInfo.InvariantCulture);

    private static PresentationGroup BuildTreemap(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "vector treemap titles must use the bounded string form");

        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateTreemapCompileProfile(element, raw, namedStyle, inlineStyle);
        var treemap = FirstProperty(inlineStyle, namedStyle, "treemap")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var values = series.Values.Select(value => value!.Value).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new TreemapNode(
            index,
            name,
            values[index],
            series.Parents[index] is { } parent ? indexes[parent] : null)).ToArray();
        foreach (var node in nodes)
            if (node.ParentIndex is { } parentIndex) nodes[parentIndex].Children.Add(node.Index);
        var roots = nodes.Where(node => node.ParentIndex is null).Select(node => node.Index).ToArray();

        var rootColors = treemap.GetProperty("rootColors").EnumerateArray()
            .Select(catalog.Color)
            .ToArray();
        var border = Property(treemap, "border");
        var gap = OptionalDouble(treemap, "gap") ?? 2;
        var headerHeight = OptionalDouble(treemap, "headerHeight") ?? 17;
        var depthLighten = OptionalDouble(treemap, "depthLighten") ?? 0.08;
        var showValues = OptionalBoolean(treemap, "showValues") ?? false;
        var labelStyle = Property(treemap, "labelTextStyle");
        var valueStyle = Property(treemap, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var plot = new TreemapRectangle(x, y + titleHeight + 4, width, height - titleHeight - 4);
        if (plot.Width < 140 || plot.Height < 90)
            throw Unsupported(element.Id, "treemap frame is too small for a readable native hierarchy");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTextElement(
                TreemapNativeId(element.Id, "title"),
                "treemap title",
                x,
                y,
                width,
                titleHeight,
                titleText,
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var rootPlacements = SquarifyTreemap(roots, nodes, plot);
        for (var rootOrder = 0; rootOrder < rootPlacements.Count; rootOrder++)
        {
            var placement = rootPlacements[rootOrder];
            RenderTreemapNode(
                group,
                element.Id,
                nodes,
                placement.NodeIndex,
                placement.Rectangle,
                rootColors[rootOrder % rootColors.Length],
                depth: 0,
                gap,
                headerHeight,
                depthLighten,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateTreemapCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var count = element.Data.Categories.Count;
        if (count is < 1 or > 128 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != count)
            throw Unsupported(element.Id, "treemap categories must be 1..128 unique non-empty strings");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "treemap charts require exactly one hierarchy series");
        var series = element.Data.Series[0];
        if (series.Values.Count != count || series.Values.Any(value => value is null || value <= 0) || series.Parents.Count != count)
            throw Unsupported(element.Id, "treemap values and parents must be complete, aligned, and strictly positive");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"treemap series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "treemap charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"treemap charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "treemap") is null)
            throw Unsupported(element.Id, "treemap charts require style.treemap");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var rootCount = 0;
        for (var index = 0; index < count; index++)
        {
            var parent = series.Parents[index];
            if (parent is null)
            {
                rootCount++;
                continue;
            }
            if (!indexes.ContainsKey(parent))
                throw Unsupported(element.Id, $"treemap parent {parent} does not name a declared category");
            if (string.Equals(parent, names[index], StringComparison.Ordinal))
                throw Unsupported(element.Id, "a treemap node cannot parent itself");
        }
        if (rootCount is < 1 or > 16)
            throw Unsupported(element.Id, "treemap charts require between 1 and 16 roots");

        for (var index = 0; index < count; index++)
        {
            var seen = new HashSet<int>();
            var current = index;
            var depth = 0;
            while (true)
            {
                if (!seen.Add(current)) throw Unsupported(element.Id, $"treemap parent chain for {names[index]} contains a cycle");
                var parent = series.Parents[current];
                if (parent is null) break;
                current = indexes[parent];
                depth++;
                if (depth > 8) throw Unsupported(element.Id, $"treemap node {names[index]} exceeds the maximum depth of eight");
            }
        }

        var childSums = new Dictionary<int, double>();
        for (var index = 0; index < count; index++)
            if (series.Parents[index] is { } parent)
            {
                var parentIndex = indexes[parent];
                childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
            }
        foreach (var pair in childSums)
        {
            var declared = series.Values[pair.Key]!.Value;
            var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
            if (Math.Abs(declared - pair.Value) > tolerance)
                throw Unsupported(element.Id, $"treemap parent {names[pair.Key]} must equal its direct-child sum");
        }
    }

    private static void RenderTreemapNode(
        PresentationGroup group,
        string elementId,
        IReadOnlyList<TreemapNode> nodes,
        int nodeIndex,
        TreemapRectangle allocated,
        (string Rgb, double Alpha) rootColor,
        int depth,
        double gap,
        double headerHeight,
        double depthLighten,
        bool showValues,
        JsonElement? border,
        JsonElement? labelStyle,
        JsonElement? valueStyle,
        Catalog catalog)
    {
        var node = nodes[nodeIndex];
        var localGap = Math.Min(gap, Math.Max(0, Math.Min(allocated.Width, allocated.Height) - 0.5) / 2);
        var rectangle = allocated.Inset(localGap / 2);
        if (rectangle.Width < 0.5 || rectangle.Height < 0.5)
            throw Unsupported(elementId, $"treemap node {node.Name} is too small for a stable native rectangle");

        var color = LightenTreemapColor(rootColor, Math.Min(0.72, depth * depthLighten));
        var shape = ShapeFrame(
            new PpjFrameModel(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, 0, false, false),
            "rect");
        shape.FillRgb = color.Rgb;
        if (color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(color.Alpha);
        if (border is { } stroke) ApplyLine(shape, stroke, catalog);
        else shape.LineStyle = "none";
        group.Children.Add(new PresentationElement
        {
            Id = TreemapNativeId(elementId, $"node/{node.Index}"),
            Name = $"treemap node {node.Name}",
            Shape = shape,
        });

        var contrast = TreemapLuminance(color.Rgb) > 0.5 ? "111827" : "FFFFFF";
        if (node.Children.Count != 0)
        {
            var effectiveHeader = Math.Min(headerHeight, Math.Max(0, rectangle.Height * 0.28));
            if (effectiveHeader >= 9 && rectangle.Width >= 28)
                group.Children.Add(VectorChartTextElement(
                    TreemapNativeId(elementId, $"label/{node.Index}"),
                    $"treemap label {node.Name}",
                    rectangle.X + 3,
                    rectangle.Y + 1,
                    rectangle.Width - 6,
                    effectiveHeader - 2,
                    node.Name,
                    labelStyle,
                    catalog,
                    Math.Clamp(Math.Min(effectiveHeader * 0.55, rectangle.Width / Math.Max(4, node.Name.Length) * 1.4), 6, 11),
                    "left",
                    contrast));
            var childArea = new TreemapRectangle(
                rectangle.X,
                rectangle.Y + effectiveHeader,
                rectangle.Width,
                rectangle.Height - effectiveHeader);
            if (childArea.Width < 1 || childArea.Height < 1)
                throw Unsupported(elementId, $"treemap node {node.Name} leaves no room for its children");
            foreach (var placement in SquarifyTreemap(node.Children, nodes, childArea))
                RenderTreemapNode(
                    group,
                    elementId,
                    nodes,
                    placement.NodeIndex,
                    placement.Rectangle,
                    rootColor,
                    depth + 1,
                    gap,
                    headerHeight,
                    depthLighten,
                    showValues,
                    border,
                    labelStyle,
                    valueStyle,
                    catalog);
            return;
        }

        if (rectangle.Width < 28 || rectangle.Height < 14) return;
        var labelHeight = showValues && rectangle.Height >= 28 ? rectangle.Height * 0.58 : rectangle.Height;
        group.Children.Add(VectorChartTextElement(
            TreemapNativeId(elementId, $"label/{node.Index}"),
            $"treemap label {node.Name}",
            rectangle.X + 3,
            rectangle.Y + 1,
            rectangle.Width - 6,
            Math.Max(10, labelHeight - 2),
            node.Name,
            labelStyle,
            catalog,
            Math.Clamp(Math.Min(labelHeight * 0.34, rectangle.Width / Math.Max(4, node.Name.Length) * 1.5), 6, 12),
            "left",
            contrast));
        if (showValues && rectangle.Height >= 28)
            group.Children.Add(VectorChartTextElement(
                TreemapNativeId(elementId, $"value/{node.Index}"),
                $"treemap value {node.Name}",
                rectangle.X + 3,
                rectangle.Y + labelHeight,
                rectangle.Width - 6,
                rectangle.Height - labelHeight - 2,
                node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                valueStyle,
                catalog,
                Math.Clamp((rectangle.Height - labelHeight) * 0.45, 6, 10),
                "left",
                contrast));
    }

    private static IReadOnlyList<TreemapPlacement> SquarifyTreemap(
        IReadOnlyList<int> nodeIndexes,
        IReadOnlyList<TreemapNode> nodes,
        TreemapRectangle rectangle)
    {
        var total = nodeIndexes.Sum(index => nodes[index].Value);
        var scale = rectangle.Width * rectangle.Height / total;
        var pending = nodeIndexes
            .Select((nodeIndex, order) => new TreemapArea(nodeIndex, nodes[nodeIndex].Value * scale, order))
            .OrderByDescending(item => item.Area)
            .ThenBy(item => item.Order)
            .ToList();
        var placements = new List<TreemapPlacement>(pending.Count);
        var row = new List<TreemapArea>();
        var remaining = rectangle;
        while (pending.Count != 0)
        {
            var candidate = pending[0];
            var side = Math.Min(remaining.Width, remaining.Height);
            if (row.Count == 0 || TreemapWorst(row.Append(candidate), side) <= TreemapWorst(row, side))
            {
                row.Add(candidate);
                pending.RemoveAt(0);
                continue;
            }
            remaining = LayoutTreemapRow(row, remaining, placements);
            row.Clear();
        }
        if (row.Count != 0) LayoutTreemapRow(row, remaining, placements);
        return placements;
    }

    private static double TreemapWorst(IEnumerable<TreemapArea> row, double side)
    {
        var values = row.Select(item => item.Area).ToArray();
        var sum = values.Sum();
        var squaredSide = side * side;
        return Math.Max(squaredSide * values.Max() / (sum * sum), (sum * sum) / (squaredSide * values.Min()));
    }

    private static TreemapRectangle LayoutTreemapRow(
        IReadOnlyList<TreemapArea> row,
        TreemapRectangle rectangle,
        ICollection<TreemapPlacement> placements)
    {
        var area = row.Sum(item => item.Area);
        if (rectangle.Width >= rectangle.Height)
        {
            var rowWidth = area / rectangle.Height;
            var cursorY = rectangle.Y;
            for (var index = 0; index < row.Count; index++)
            {
                var height = index == row.Count - 1
                    ? rectangle.Y + rectangle.Height - cursorY
                    : row[index].Area / rowWidth;
                placements.Add(new TreemapPlacement(row[index].NodeIndex, new TreemapRectangle(rectangle.X, cursorY, rowWidth, height)));
                cursorY += height;
            }
            return new TreemapRectangle(rectangle.X + rowWidth, rectangle.Y, Math.Max(0, rectangle.Width - rowWidth), rectangle.Height);
        }

        var rowHeight = area / rectangle.Width;
        var cursorX = rectangle.X;
        for (var index = 0; index < row.Count; index++)
        {
            var width = index == row.Count - 1
                ? rectangle.X + rectangle.Width - cursorX
                : row[index].Area / rowHeight;
            placements.Add(new TreemapPlacement(row[index].NodeIndex, new TreemapRectangle(cursorX, rectangle.Y, width, rowHeight)));
            cursorX += width;
        }
        return new TreemapRectangle(rectangle.X, rectangle.Y + rowHeight, rectangle.Width, Math.Max(0, rectangle.Height - rowHeight));
    }

    private static (string Rgb, double Alpha) LightenTreemapColor((string Rgb, double Alpha) color, double amount)
    {
        byte Lighten(byte channel) => checked((byte)Math.Round(channel + (255 - channel) * amount));
        var red = byte.Parse(color.Rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = byte.Parse(color.Rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = byte.Parse(color.Rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return ($"{Lighten(red):X2}{Lighten(green):X2}{Lighten(blue):X2}", color.Alpha);
    }

    private static double TreemapLuminance(string rgb)
    {
        static double Channel(byte value)
        {
            var normalized = value / 255d;
            return normalized <= 0.03928 ? normalized / 12.92 : Math.Pow((normalized + 0.055) / 1.055, 2.4);
        }
        var red = byte.Parse(rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = byte.Parse(rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = byte.Parse(rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return 0.2126 * Channel(red) + 0.7152 * Channel(green) + 0.0722 * Channel(blue);
    }

    private static string TreemapNativeId(string elementId, string suffix) =>
        $"{elementId}/treemap/{suffix}";

    private sealed class TreemapNode(int index, string name, double value, int? parentIndex)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal double Value { get; } = value;
        internal int? ParentIndex { get; } = parentIndex;
        internal List<int> Children { get; } = [];
    }

    private readonly record struct TreemapRectangle(double X, double Y, double Width, double Height)
    {
        internal TreemapRectangle Inset(double amount) => new(
            X + amount,
            Y + amount,
            Math.Max(0, Width - amount * 2),
            Math.Max(0, Height - amount * 2));
    }

    private readonly record struct TreemapPlacement(int NodeIndex, TreemapRectangle Rectangle);
    private readonly record struct TreemapArea(int NodeIndex, double Area, int Order);

    private static PresentationGroup BuildSunburst(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "vector sunburst titles must use the bounded string form");

        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateSunburstCompileProfile(element, raw, namedStyle, inlineStyle);
        var sunburst = FirstProperty(inlineStyle, namedStyle, "sunburst")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var values = series.Values.Select(value => value!.Value).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new SunburstNode(
            index,
            name,
            values[index],
            series.Parents[index] is { } parent ? indexes[parent] : null)).ToArray();
        foreach (var node in nodes)
            if (node.ParentIndex is { } parentIndex) nodes[parentIndex].Children.Add(node.Index);
        var roots = nodes.Where(node => node.ParentIndex is null).Select(node => node.Index).ToArray();
        foreach (var root in roots) AssignSunburstDepth(nodes, root, 0);
        var levelCount = nodes.Max(node => node.Depth) + 1;

        var rootColors = sunburst.GetProperty("rootColors").EnumerateArray().Select(catalog.Color).ToArray();
        var border = Property(sunburst, "border");
        var innerRadiusRatio = OptionalDouble(sunburst, "innerRadiusRatio") ?? 0;
        var ringGap = OptionalDouble(sunburst, "ringGap") ?? 1.5;
        var segmentGapRadians = (OptionalDouble(sunburst, "segmentGapDegrees") ?? 0.8) * Math.PI / 180;
        var startAngle = (OptionalDouble(sunburst, "startAngle") ?? -90) * Math.PI / 180;
        var clockwise = OptionalBoolean(sunburst, "clockwise") ?? true;
        var direction = clockwise ? 1d : -1d;
        var depthLighten = OptionalDouble(sunburst, "depthLighten") ?? 0.08;
        var showValues = OptionalBoolean(sunburst, "showValues") ?? false;
        var labelStyle = Property(sunburst, "labelTextStyle");
        var valueStyle = Property(sunburst, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var availableHeight = height - titleHeight - 4;
        var diameter = Math.Min(width, availableHeight);
        if (diameter < 160)
            throw Unsupported(element.Id, "sunburst frame is too small for a readable native hierarchy");
        var plotX = x + (width - diameter) / 2;
        var plotY = y + titleHeight + 4 + (availableHeight - diameter) / 2;
        var outerRadius = diameter / 2;
        var innerRadius = outerRadius * innerRadiusRatio;
        var ringWidth = (outerRadius - innerRadius) / levelCount;
        if (ringWidth - ringGap < 8)
            throw Unsupported(element.Id, "sunburst rings are too narrow after applying the configured gap");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTextElement(
                SunburstNativeId(element.Id, "title"),
                "sunburst title",
                x,
                y,
                width,
                titleHeight,
                titleText,
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var total = roots.Sum(root => nodes[root].Value);
        var cursor = startAngle;
        for (var rootOrder = 0; rootOrder < roots.Length; rootOrder++)
        {
            var root = roots[rootOrder];
            var span = Math.PI * 2 * nodes[root].Value / total * direction;
            RenderSunburstNode(
                group,
                element.Id,
                nodes,
                root,
                cursor,
                cursor + span,
                rootColors[rootOrder % rootColors.Length],
                plotX,
                plotY,
                diameter,
                innerRadius,
                ringWidth,
                ringGap,
                segmentGapRadians,
                depthLighten,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
            cursor += span;
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateSunburstCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var count = element.Data.Categories.Count;
        if (count is < 1 or > 96 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != count)
            throw Unsupported(element.Id, "sunburst categories must be 1..96 unique non-empty strings");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "sunburst charts require exactly one hierarchy series");
        var series = element.Data.Series[0];
        if (series.Values.Count != count || series.Values.Any(value => value is null || value <= 0) || series.Parents.Count != count)
            throw Unsupported(element.Id, "sunburst values and parents must be complete, aligned, and strictly positive");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"sunburst series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "sunburst charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"sunburst charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "sunburst") is null)
            throw Unsupported(element.Id, "sunburst charts require style.sunburst");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var rootCount = 0;
        for (var index = 0; index < count; index++)
        {
            var parent = series.Parents[index];
            if (parent is null)
            {
                rootCount++;
                continue;
            }
            if (!indexes.ContainsKey(parent))
                throw Unsupported(element.Id, $"sunburst parent {parent} does not name a declared category");
            if (string.Equals(parent, names[index], StringComparison.Ordinal))
                throw Unsupported(element.Id, "a sunburst node cannot parent itself");
        }
        if (rootCount is < 1 or > 16)
            throw Unsupported(element.Id, "sunburst charts require between 1 and 16 roots");

        for (var index = 0; index < count; index++)
        {
            var seen = new HashSet<int>();
            var current = index;
            var depth = 0;
            while (true)
            {
                if (!seen.Add(current)) throw Unsupported(element.Id, $"sunburst parent chain for {names[index]} contains a cycle");
                var parent = series.Parents[current];
                if (parent is null) break;
                current = indexes[parent];
                depth++;
                if (depth > 6) throw Unsupported(element.Id, $"sunburst node {names[index]} exceeds the maximum depth of six");
            }
        }

        var childSums = new Dictionary<int, double>();
        for (var index = 0; index < count; index++)
            if (series.Parents[index] is { } parent)
            {
                var parentIndex = indexes[parent];
                childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
            }
        foreach (var pair in childSums)
        {
            var declared = series.Values[pair.Key]!.Value;
            var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
            if (Math.Abs(declared - pair.Value) > tolerance)
                throw Unsupported(element.Id, $"sunburst parent {names[pair.Key]} must equal its direct-child sum");
        }
    }

    private static void AssignSunburstDepth(IReadOnlyList<SunburstNode> nodes, int nodeIndex, int depth)
    {
        nodes[nodeIndex].Depth = depth;
        foreach (var child in nodes[nodeIndex].Children) AssignSunburstDepth(nodes, child, depth + 1);
    }

    private static void RenderSunburstNode(
        PresentationGroup group,
        string elementId,
        IReadOnlyList<SunburstNode> nodes,
        int nodeIndex,
        double startAngle,
        double endAngle,
        (string Rgb, double Alpha) rootColor,
        double plotX,
        double plotY,
        double diameter,
        double baseInnerRadius,
        double ringWidth,
        double ringGap,
        double segmentGapRadians,
        double depthLighten,
        bool showValues,
        JsonElement? border,
        JsonElement? labelStyle,
        JsonElement? valueStyle,
        Catalog catalog)
    {
        var node = nodes[nodeIndex];
        var sign = Math.Sign(endAngle - startAngle);
        var span = Math.Abs(endAngle - startAngle);
        var effectiveSpan = span - segmentGapRadians;
        if (effectiveSpan <= Math.PI / 720)
            throw Unsupported(elementId, $"sunburst node {node.Name} is too narrow after applying the configured segment gap");
        var visibleStart = startAngle + sign * segmentGapRadians / 2;
        var visibleEnd = endAngle - sign * segmentGapRadians / 2;
        var innerRadius = baseInnerRadius + node.Depth * ringWidth + ringGap / 2;
        var outerRadius = baseInnerRadius + (node.Depth + 1) * ringWidth - ringGap / 2;
        if (outerRadius <= innerRadius)
            throw Unsupported(elementId, $"sunburst node {node.Name} has no visible ring thickness");

        var color = LightenTreemapColor(rootColor, Math.Min(0.72, node.Depth * depthLighten));
        var shape = ShapeFrame(new PpjFrameModel(plotX, plotY, diameter, diameter, 0, false, false), "custom");
        shape.FillRgb = color.Rgb;
        if (color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(color.Alpha);
        if (border is { } stroke) ApplyLine(shape, stroke, catalog);
        else shape.LineStyle = "none";
        shape.CustomPaths.Add(BuildSunburstSectorPath(
            diameter,
            innerRadius,
            outerRadius,
            visibleStart,
            visibleEnd,
            border is not null));
        group.Children.Add(new PresentationElement
        {
            Id = SunburstNativeId(elementId, $"sector/{node.Index}"),
            Name = $"sunburst sector {node.Name}",
            Shape = shape,
        });

        var midRadius = (innerRadius + outerRadius) / 2;
        var midAngle = (visibleStart + visibleEnd) / 2;
        var arcLength = midRadius * effectiveSpan;
        var thickness = outerRadius - innerRadius;
        var estimatedLabelWidth = Math.Max(28, node.Name.Length * 5.5);
        if (thickness >= (showValues ? 25 : 16) && arcLength >= estimatedLabelWidth * 1.35 && effectiveSpan >= Math.PI / 12)
        {
            var centerX = plotX + diameter / 2 + Math.Cos(midAngle) * midRadius;
            var centerY = plotY + diameter / 2 + Math.Sin(midAngle) * midRadius;
            var textWidth = Math.Min(Math.Max(estimatedLabelWidth, 36), Math.Min(110, arcLength * 0.72));
            var labelHeight = showValues ? Math.Min(14, thickness * 0.42) : Math.Min(16, thickness * 0.64);
            var contrast = TreemapLuminance(color.Rgb) > 0.5 ? "111827" : "FFFFFF";
            group.Children.Add(VectorChartTextElement(
                SunburstNativeId(elementId, $"label/{node.Index}"),
                $"sunburst label {node.Name}",
                centerX - textWidth / 2,
                centerY - (showValues ? labelHeight : labelHeight / 2),
                textWidth,
                labelHeight,
                node.Name,
                labelStyle,
                catalog,
                Math.Clamp(Math.Min(labelHeight * 0.62, textWidth / Math.Max(4, node.Name.Length) * 1.5), 6, 11),
                "center",
                contrast));
            if (showValues)
                group.Children.Add(VectorChartTextElement(
                    SunburstNativeId(elementId, $"value/{node.Index}"),
                    $"sunburst value {node.Name}",
                    centerX - textWidth / 2,
                    centerY,
                    textWidth,
                    Math.Min(12, thickness * 0.34),
                    node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                    valueStyle,
                    catalog,
                    Math.Clamp(thickness * 0.18, 6, 9),
                    "center",
                    contrast));
        }

        if (node.Children.Count == 0) return;
        var cursor = startAngle;
        var parentSpan = endAngle - startAngle;
        foreach (var childIndex in node.Children)
        {
            var childSpan = parentSpan * nodes[childIndex].Value / node.Value;
            RenderSunburstNode(
                group,
                elementId,
                nodes,
                childIndex,
                cursor,
                cursor + childSpan,
                rootColor,
                plotX,
                plotY,
                diameter,
                baseInnerRadius,
                ringWidth,
                ringGap,
                segmentGapRadians,
                depthLighten,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
            cursor += childSpan;
        }
    }

    private static PresentationCustomGeometryPath BuildSunburstSectorPath(
        double diameter,
        double innerRadius,
        double outerRadius,
        double startAngle,
        double endAngle,
        bool stroke)
    {
        const long viewport = 100_000;
        const double center = viewport / 2d;
        var scale = viewport / diameter;
        var inner = innerRadius * scale;
        var outer = outerRadius * scale;
        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = stroke,
        };
        path.Commands.Add(MoveTo(SunburstPoint(center, outer, startAngle)));
        AddSunburstArc(path, center, outer, startAngle, endAngle);
        if (inner <= 0.5)
        {
            path.Commands.Add(LineTo(new PresentationCustomGeometryPoint { X = viewport / 2, Y = viewport / 2 }));
        }
        else
        {
            path.Commands.Add(LineTo(SunburstPoint(center, inner, endAngle)));
            AddSunburstArc(path, center, inner, endAngle, startAngle);
        }
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static void AddSunburstArc(
        PresentationCustomGeometryPath path,
        double center,
        double radius,
        double startAngle,
        double endAngle)
    {
        var span = endAngle - startAngle;
        var segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(span) / (Math.PI / 2)));
        var step = span / segments;
        for (var segment = 0; segment < segments; segment++)
        {
            var start = startAngle + step * segment;
            var end = start + step;
            var kappa = 4d / 3d * Math.Tan(step / 4d);
            var startPoint = SunburstPoint(center, radius, start);
            var endPoint = SunburstPoint(center, radius, end);
            var control1 = new PresentationCustomGeometryPoint
            {
                X = checked((long)Math.Round(startPoint.X - kappa * radius * Math.Sin(start), MidpointRounding.AwayFromZero)),
                Y = checked((long)Math.Round(startPoint.Y + kappa * radius * Math.Cos(start), MidpointRounding.AwayFromZero)),
            };
            var control2 = new PresentationCustomGeometryPoint
            {
                X = checked((long)Math.Round(endPoint.X + kappa * radius * Math.Sin(end), MidpointRounding.AwayFromZero)),
                Y = checked((long)Math.Round(endPoint.Y - kappa * radius * Math.Cos(end), MidpointRounding.AwayFromZero)),
            };
            path.Commands.Add(new PresentationCustomGeometryCommand
            {
                CubicBezierTo = new PresentationCustomGeometryCubicBezier
                {
                    Control1 = control1,
                    Control2 = control2,
                    End = endPoint,
                },
            });
        }
    }

    private static PresentationCustomGeometryPoint SunburstPoint(double center, double radius, double angle) => new()
    {
        X = checked((long)Math.Round(center + radius * Math.Cos(angle), MidpointRounding.AwayFromZero)),
        Y = checked((long)Math.Round(center + radius * Math.Sin(angle), MidpointRounding.AwayFromZero)),
    };

    private static PresentationCustomGeometryCommand MoveTo(PresentationCustomGeometryPoint point) => new() { MoveTo = point };
    private static PresentationCustomGeometryCommand LineTo(PresentationCustomGeometryPoint point) => new() { LineTo = point };

    private static string SunburstNativeId(string elementId, string suffix) =>
        $"{elementId}/sunburst/{suffix}";

    private sealed class SunburstNode(int index, string name, double value, int? parentIndex)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal double Value { get; } = value;
        internal int? ParentIndex { get; } = parentIndex;
        internal int Depth { get; set; }
        internal List<int> Children { get; } = [];
    }

    private static PresentationGroup BuildSankey(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
            throw Unsupported(element.Id, "vector sankey titles must use the bounded string form");

        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateSankeyCompileProfile(element, raw, namedStyle, inlineStyle);
        var sankey = FirstProperty(inlineStyle, namedStyle, "sankey")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new SankeyNode(index, name)).ToArray();
        var edges = Enumerable.Range(0, series.Values.Count).Select(index => new SankeyEdge(
            index,
            indexes[series.Sources[index]],
            indexes[series.Targets[index]],
            series.Values[index]!.Value)).ToArray();
        foreach (var edge in edges)
        {
            nodes[edge.SourceIndex].Outgoing.Add(edge.Index);
            nodes[edge.TargetIndex].Incoming.Add(edge.Index);
            nodes[edge.SourceIndex].OutgoingValue += edge.Value;
            nodes[edge.TargetIndex].IncomingValue += edge.Value;
        }

        var indegree = nodes.Select(node => node.Incoming.Count).ToArray();
        var queue = new Queue<int>(nodes.Where(node => indegree[node.Index] == 0).Select(node => node.Index));
        while (queue.Count != 0)
        {
            var nodeIndex = queue.Dequeue();
            foreach (var edgeIndex in nodes[nodeIndex].Outgoing)
            {
                var edge = edges[edgeIndex];
                nodes[edge.TargetIndex].Column = Math.Max(nodes[edge.TargetIndex].Column, nodes[nodeIndex].Column + 1);
                if (--indegree[edge.TargetIndex] == 0) queue.Enqueue(edge.TargetIndex);
            }
        }
        var maximumColumn = nodes.Max(node => node.Column);

        var nodeColors = sankey.GetProperty("nodeColors").EnumerateArray().Select(catalog.Color).ToArray();
        var nodeStroke = Property(sankey, "nodeStroke");
        var nodeWidth = OptionalDouble(sankey, "nodeWidth") ?? 12;
        var nodeGap = OptionalDouble(sankey, "nodeGap") ?? 9;
        var justify = (OptionalString(sankey, "nodeAlign") ?? "justify") == "justify";
        if (justify)
            foreach (var node in nodes.Where(node => node.Outgoing.Count == 0)) node.Column = maximumColumn;
        var flowOpacity = OptionalDouble(sankey, "flowOpacity") ?? 0.45;
        var flowCurvature = OptionalDouble(sankey, "flowCurvature") ?? 0.7;
        var flowColorMode = OptionalString(sankey, "flowColorMode") ?? "source";
        var showValues = OptionalBoolean(sankey, "showValues") ?? false;
        var labelStyle = Property(sankey, "labelTextStyle");
        var valueStyle = Property(sankey, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");
        foreach (var node in nodes) node.Color = nodeColors[node.Index % nodeColors.Length];

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var plotX = x;
        var plotY = y + titleHeight + 4;
        var plotWidth = width;
        var plotHeight = height - titleHeight - 4;
        if (plotWidth < 260 || plotHeight < 130 || maximumColumn < 1)
            throw Unsupported(element.Id, "sankey frame or graph depth is too small for a readable native flow layout");

        var columns = nodes.GroupBy(node => node.Column).OrderBy(group => group.Key).ToArray();
        var scale = columns.Min(column =>
        {
            var available = plotHeight - nodeGap * (column.Count() - 1);
            var magnitude = column.Sum(node => node.Value);
            return available / magnitude;
        });
        if (!double.IsFinite(scale) || scale <= 0)
            throw Unsupported(element.Id, "sankey node gaps leave no room for positive flow thickness");
        var columnStep = (plotWidth - nodeWidth) / maximumColumn;
        if (columnStep - nodeWidth < 48)
            throw Unsupported(element.Id, "sankey columns leave insufficient horizontal room for native ribbons and labels");
        foreach (var column in columns)
        {
            var ordered = column.OrderBy(node => node.Index).ToArray();
            var occupied = ordered.Sum(node => node.Value * scale) + nodeGap * (ordered.Length - 1);
            var cursorY = plotY + (plotHeight - occupied) / 2;
            foreach (var node in ordered)
            {
                node.X = plotX + node.Column * columnStep;
                node.Y = cursorY;
                node.Height = node.Value * scale;
                cursorY += node.Height + nodeGap;
            }
        }

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTextElement(
                SankeyNativeId(element.Id, "title"),
                "sankey title",
                x,
                y,
                width,
                titleHeight,
                titleText,
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var outgoingOffsets = new double[nodes.Length];
        var incomingOffsets = new double[nodes.Length];
        foreach (var edge in edges)
        {
            var source = nodes[edge.SourceIndex];
            var target = nodes[edge.TargetIndex];
            var thickness = edge.Value * scale;
            var sourceTop = source.Y + outgoingOffsets[source.Index];
            var targetTop = target.Y + incomingOffsets[target.Index];
            var color = flowColorMode == "target" ? target.Color : source.Color;
            var ribbon = ShapeFrame(new PpjFrameModel(plotX, plotY, plotWidth, plotHeight, 0, false, false), "custom");
            ribbon.FillRgb = color.Rgb;
            ribbon.FillOpacityThousandthPercent = Opacity(flowOpacity * color.Alpha);
            ribbon.LineStyle = "none";
            ribbon.CustomPaths.Add(BuildSankeyRibbonPath(
                plotX,
                plotY,
                plotWidth,
                plotHeight,
                source.X + nodeWidth,
                sourceTop,
                sourceTop + thickness,
                target.X,
                targetTop,
                targetTop + thickness,
                flowCurvature));
            group.Children.Add(new PresentationElement
            {
                Id = SankeyNativeId(element.Id, $"flow/{edge.Index}"),
                Name = $"sankey flow {source.Name} to {target.Name}",
                Shape = ribbon,
            });
            outgoingOffsets[source.Index] += thickness;
            incomingOffsets[target.Index] += thickness;
        }

        var labelWidth = Math.Clamp(columnStep - nodeWidth - 8, 42, 120);
        foreach (var node in nodes)
        {
            var shape = ShapeFrame(new PpjFrameModel(node.X, node.Y, nodeWidth, node.Height, 0, false, false), "rect");
            shape.FillRgb = node.Color.Rgb;
            if (node.Color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(node.Color.Alpha);
            if (nodeStroke is { } stroke) ApplyLine(shape, stroke, catalog);
            else shape.LineStyle = "none";
            group.Children.Add(new PresentationElement
            {
                Id = SankeyNativeId(element.Id, $"node/{node.Index}"),
                Name = $"sankey node {node.Name}",
                Shape = shape,
            });

            if (node.Height < 10) continue;
            var placeLeft = node.Column == maximumColumn;
            var textX = placeLeft ? node.X - labelWidth - 4 : node.X + nodeWidth + 4;
            var labelHeight = showValues && node.Height >= 24 ? Math.Min(14, node.Height * 0.48) : Math.Min(16, node.Height);
            group.Children.Add(VectorChartTextElement(
                SankeyNativeId(element.Id, $"label/{node.Index}"),
                $"sankey label {node.Name}",
                textX,
                node.Y,
                labelWidth,
                labelHeight,
                node.Name,
                labelStyle,
                catalog,
                Math.Clamp(Math.Min(labelHeight * 0.62, labelWidth / Math.Max(4, node.Name.Length) * 1.5), 6, 10),
                placeLeft ? "right" : "left",
                "16324F"));
            if (showValues && node.Height >= 24)
                group.Children.Add(VectorChartTextElement(
                    SankeyNativeId(element.Id, $"value/{node.Index}"),
                    $"sankey value {node.Name}",
                    textX,
                    node.Y + labelHeight,
                    labelWidth,
                    Math.Min(12, node.Height - labelHeight),
                    node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                    valueStyle,
                    catalog,
                    Math.Clamp(node.Height * 0.18, 6, 9),
                    placeLeft ? "right" : "left",
                    "52606D"));
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateSankeyCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var nodeCount = element.Data.Categories.Count;
        if (nodeCount is < 2 or > 64 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != nodeCount)
            throw Unsupported(element.Id, "sankey categories must be 2..64 unique non-empty node names");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "sankey charts require exactly one directed-flow series");
        var series = element.Data.Series[0];
        var edgeCount = series.Values.Count;
        if (edgeCount is < 1 or > 256 || series.Values.Any(value => value is null || value <= 0) ||
            series.Sources.Count != edgeCount || series.Targets.Count != edgeCount)
            throw Unsupported(element.Id, "sankey sources, targets, and positive flow values must align across 1..256 edges");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"sankey series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "sankey charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap", "sunburst" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"sankey charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "sankey") is null)
            throw Unsupported(element.Id, "sankey charts require style.sankey");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var indegree = new int[nodeCount];
        var outgoing = Enumerable.Range(0, nodeCount).Select(_ => new List<int>()).ToArray();
        var incomingFlow = new double[nodeCount];
        var outgoingFlow = new double[nodeCount];
        var used = new bool[nodeCount];
        var edgeKeys = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < edgeCount; index++)
        {
            if (!indexes.TryGetValue(series.Sources[index], out var source) || !indexes.TryGetValue(series.Targets[index], out var target))
                throw Unsupported(element.Id, "every sankey source and target must name a declared category");
            if (source == target) throw Unsupported(element.Id, "a sankey edge cannot target its source node");
            if (!edgeKeys.Add(series.Sources[index] + "\u0000" + series.Targets[index]))
                throw Unsupported(element.Id, "duplicate sankey endpoints must be combined into one explicit flow value");
            outgoing[source].Add(target);
            indegree[target]++;
            used[source] = true;
            used[target] = true;
            outgoingFlow[source] += series.Values[index]!.Value;
            incomingFlow[target] += series.Values[index]!.Value;
        }
        if (used.Any(value => !value)) throw Unsupported(element.Id, "every declared sankey node must be incident to an edge");

        var queue = new Queue<int>(Enumerable.Range(0, nodeCount).Where(index => indegree[index] == 0));
        var visited = 0;
        while (queue.Count != 0)
        {
            var source = queue.Dequeue();
            visited++;
            foreach (var target in outgoing[source]) if (--indegree[target] == 0) queue.Enqueue(target);
        }
        if (visited != nodeCount) throw Unsupported(element.Id, "sankey edges must form a directed acyclic graph");
        for (var index = 0; index < nodeCount; index++)
            if (incomingFlow[index] > 0 && outgoingFlow[index] > 0)
            {
                var tolerance = Math.Max(1e-9, Math.Max(incomingFlow[index], outgoingFlow[index]) * 1e-9);
                if (Math.Abs(incomingFlow[index] - outgoingFlow[index]) > tolerance)
                    throw Unsupported(element.Id, $"sankey internal node {names[index]} must conserve flow");
            }
    }

    private static PresentationCustomGeometryPath BuildSankeyRibbonPath(
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        double sourceX,
        double sourceTop,
        double sourceBottom,
        double targetX,
        double targetTop,
        double targetBottom,
        double curvature)
    {
        const long viewport = 100_000;
        PresentationCustomGeometryPoint Point(double x, double y) => new()
        {
            X = checked((long)Math.Round((x - plotX) / plotWidth * viewport, MidpointRounding.AwayFromZero)),
            Y = checked((long)Math.Round((y - plotY) / plotHeight * viewport, MidpointRounding.AwayFromZero)),
        };
        var controlFraction = curvature * 0.5;
        var control1X = sourceX + (targetX - sourceX) * controlFraction;
        var control2X = targetX - (targetX - sourceX) * controlFraction;
        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = false,
        };
        path.Commands.Add(MoveTo(Point(sourceX, sourceTop)));
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = Point(control1X, sourceTop),
                Control2 = Point(control2X, targetTop),
                End = Point(targetX, targetTop),
            },
        });
        path.Commands.Add(LineTo(Point(targetX, targetBottom)));
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = Point(control2X, targetBottom),
                Control2 = Point(control1X, sourceBottom),
                End = Point(sourceX, sourceBottom),
            },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static string SankeyNativeId(string elementId, string suffix) =>
        $"{elementId}/sankey/{suffix}";

    private sealed class SankeyNode(int index, string name)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal List<int> Incoming { get; } = [];
        internal List<int> Outgoing { get; } = [];
        internal double IncomingValue { get; set; }
        internal double OutgoingValue { get; set; }
        internal double Value => Math.Max(IncomingValue, OutgoingValue);
        internal int Column { get; set; }
        internal double X { get; set; }
        internal double Y { get; set; }
        internal double Height { get; set; }
        internal (string Rgb, double Alpha) Color { get; set; }
    }

    private sealed record SankeyEdge(int Index, int SourceIndex, int TargetIndex, double Value);

    private static void ValidateWaterfallCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "waterfall charts require exactly one semantic series");
        var series = raw.GetProperty("data").GetProperty("series")[0];
        RejectProperties(
            series,
            element.Id,
            "chartType",
            "axis",
            "color",
            "fill",
            "stroke",
            "marker",
            "trendlines",
            "errorBars",
            "xValues",
            "bubbleSizes");
        foreach (var name in new[] { "stacking", "showDataLabels", "dataLabelPosition", "dataLabels", "smooth", "varyColors" })
            if (FirstProperty(inlineStyle, namedStyle, name) is not null)
                throw Unsupported(element.Id, $"{name} is outside the bounded waterfall style profile");
        if (FirstProperty(inlineStyle, namedStyle, "legend") is { } legend && legend.GetString() != "none")
            throw Unsupported(element.Id, "waterfall charts do not expose their internal lowering series through a legend");
        if (FirstProperty(inlineStyle, namedStyle, "waterfall") is not { } waterfall)
            throw Unsupported(element.Id, "waterfall charts require style.waterfall increase, decrease, and total role styles");
        foreach (var role in new[] { "increase", "decrease", "total" })
        {
            var fillType = waterfall.GetProperty(role).GetProperty("fill").GetProperty("type").GetString();
            if (fillType is "none" or "image")
                throw Unsupported(element.Id, $"waterfall {role} fill must be solid or a bounded gradient");
        }
    }

    private static void BuildWaterfallSeries(
        PresentationChart chart,
        PpjChartElementModel element,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var semantic = element.Data.Series[0];
        if (semantic.PointRoles.Count != semantic.Values.Count || semantic.Values.Any(value => value is null))
            throw Unsupported(element.Id, "waterfall values and pointRoles must be complete and aligned");

        var count = semantic.Values.Count;
        var offset = new double[count];
        var increase = new double[count];
        var decrease = new double[count];
        var total = new double[count];
        var increaseMissing = new List<uint>(count);
        var decreaseMissing = new List<uint>(count);
        var totalMissing = new List<uint>(count);
        var running = 0d;
        for (var index = 0; index < count; index++)
        {
            var value = semantic.Values[index]!.Value;
            if (semantic.PointRoles[index] == "total")
            {
                offset[index] = 0;
                total[index] = value;
                increaseMissing.Add(checked((uint)index));
                decreaseMissing.Add(checked((uint)index));
                running = value;
                continue;
            }

            var next = running + value;
            offset[index] = Math.Min(running, next);
            if (value >= 0)
            {
                increase[index] = value;
                decreaseMissing.Add(checked((uint)index));
            }
            else
            {
                decrease[index] = -value;
                increaseMissing.Add(checked((uint)index));
            }
            totalMissing.Add(checked((uint)index));
            running = next;
        }

        var waterfall = FirstProperty(inlineStyle, namedStyle, "waterfall")!.Value;
        chart.Grouping = "stacked";
        if (!chart.HasGapWidth) chart.GapWidth = 60;
        chart.HasLegend = false;
        chart.LegendPosition = string.Empty;
        chart.DataLabels = null;
        chart.YAxis ??= new SpreadsheetChartAxisArtifact();
        if (!chart.YAxis.HasMinimum) chart.YAxis.Minimum = 0;

        var offsetSeries = new SpreadsheetChartSeriesArtifact
        {
            Name = "__offset__",
            SeriesFill = new SpreadsheetChartSurfaceFill { NoFill = true },
            Line = new SpreadsheetChartLineStyleArtifact
            {
                Color = new SpreadsheetColor { Rgb = "000000" },
                WidthPoints = 0,
                OpacityThousandthPercent = 0,
            },
        };
        offsetSeries.Values.Add(offset);
        chart.Series.Add(offsetSeries);
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("increase"), increase, increaseMissing, catalog, element.Id, "increase"));
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("decrease"), decrease, decreaseMissing, catalog, element.Id, "decrease"));
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("total"), total, totalMissing, catalog, element.Id, "total"));
    }

    private static SpreadsheetChartSeriesArtifact WaterfallRoleSeries(
        JsonElement role,
        IReadOnlyList<double> values,
        IEnumerable<uint> missingIndexes,
        Catalog catalog,
        string elementId,
        string roleName)
    {
        var output = new SpreadsheetChartSeriesArtifact
        {
            Name = role.GetProperty("label").GetString()!,
            SeriesFill = BuildChartFill(role.GetProperty("fill"), catalog, $"{elementId} waterfall {roleName} fill"),
        };
        output.Values.Add(values);
        output.MissingValueIndexes.Add(missingIndexes);
        if (role.TryGetProperty("stroke", out var stroke)) output.Line = BuildChartLine(stroke, catalog);
        return output;
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
        for (var index = 0; index < source.Values.Count; index++)
        {
            var value = source.Values[index];
            if (value is null)
            {
                series.Values.Add(0d);
                series.MissingValueIndexes.Add(checked((uint)index));
            }
            else series.Values.Add(value.Value);
        }
        series.XValues.Add(source.XValues);
        series.BubbleSizes.Add(source.BubbleSizes);
        if (raw.TryGetProperty("fill", out var fill))
        {
            var chartFill = BuildChartFill(fill, catalog, $"{source.Id} series fill");
            if (chartFill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb && !chartFill.HasOpacityThousandthPercent)
                series.Fill = new SpreadsheetColor { Rgb = chartFill.SolidRgb };
            else
                series.SeriesFill = chartFill;
        }
        else if (raw.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            if (resolved.Alpha == 1)
                series.Fill = new SpreadsheetColor { Rgb = resolved.Rgb };
            else
                series.SeriesFill = new SpreadsheetChartSurfaceFill
                {
                    SolidRgb = resolved.Rgb,
                    OpacityThousandthPercent = Opacity(resolved.Alpha),
                };
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

    private static SpreadsheetChartAxisArtifact BuildChartAxis(JsonElement source, Catalog catalog)
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
            axis.TextStyle = BuildChartTextStyle(textStyle, catalog);
        if (source.TryGetProperty("titleTextStyle", out var titleTextStyle))
        {
            if (axis.Title.Length == 0) throw Unsupported("chart axis", "titleTextStyle requires a non-empty axis title");
            axis.TitleTextStyle = BuildChartTextStyle(titleTextStyle, catalog);
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
            marker.Fill = new SpreadsheetColor { Rgb = color.Rgb };
            if (color.Alpha < 1) marker.FillOpacityThousandthPercent = Opacity(color.Alpha);
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
        var table = new PresentationTable
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) table.FrameTransform = frameTransform;
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
                {
                    var cellWidthEmu = table.ColumnWidthsEmu
                        .Skip(cursor)
                        .Take(cell.ColumnSpan)
                        .Sum();
                    var cellHeightEmu = rowHeights
                        .Skip(rowIndex)
                        .Take(cell.RowSpan)
                        .Sum();
                    targetCell.Fill = BuildTableCellFill(
                        cellFill,
                        cellWidthEmu,
                        cellHeightEmu,
                        catalog,
                        element.Id,
                        $"table {element.Id} row {rowIndex} cell {sourceCellIndex} fill");
                }
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
        double cellWidth,
        double cellHeight,
        Catalog catalog,
        string elementId,
        string path)
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
        if (type == "image")
            return new PresentationTableCellFill
            {
                ImagePaint = PpjImagePaintLowering.Build(
                    fill,
                    cellWidth,
                    cellHeight,
                    catalog.NativeAssetId,
                    catalog.AssetDimensions,
                    path),
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

    private static PresentationGroup BuildAuthoredDiagram(
        PpjSmartArtElementModel element,
        JsonElement raw,
        Catalog catalog)
    {
        var nodeJson = raw.GetProperty("nodes").EnumerateArray().ToDictionary(
            node => node.GetProperty("id").GetString()!,
            node => node,
            StringComparer.Ordinal);
        var layout = LayoutDiagramNodes(element);
        if (layout.Any(item => item.Frame.Width < 8 || item.Frame.Height < 8) ||
            element.Layout == "picture" && layout.Any(item =>
                DiagramPictureImageFrame(item.Frame).Height < 8 || DiagramPictureLabelFrame(item.Frame).Height < 8))
            throw Unsupported(element.Id, "authored diagram frame is too small for its layout and node count");
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
        if (BuildFrameTransform(element.Frame) is { } transform) group.FrameTransform = transform;

        var nodeIds = layout.ToDictionary(
            item => item.Node.Id,
            item => DiagramChildId(element.Id, "node", item.Node.Id),
            StringComparer.Ordinal);
        if (raw.TryGetProperty("connector", out var connectorStyle))
        {
            foreach (var edge in DiagramEdges(element, layout))
            {
                var from = layout[edge.FromIndex];
                var to = layout[edge.ToIndex];
                var connector = BuildDiagramConnector(
                    from.Frame,
                    to.Frame,
                    nodeIds[from.Node.Id],
                    nodeIds[to.Node.Id],
                    connectorStyle,
                    catalog,
                    element.Layout!);
                group.Children.Add(new PresentationElement
                {
                    Id = DiagramChildId(element.Id, "edge", $"{from.Node.Id}.{to.Node.Id}"),
                    Name = $"{from.Node.Id} to {to.Node.Id}",
                    Connector = connector,
                });
            }
        }

        for (var index = 0; index < layout.Count; index++)
        {
            var item = layout[index];
            var rawNode = nodeJson[item.Node.Id];
            var nodeId = nodeIds[item.Node.Id];
            if (element.Layout == "picture")
            {
                group.Children.Add(new PresentationElement
                {
                    Id = nodeId,
                    Name = DisplayName(null, null, item.Node.Id),
                    Shape = BuildDiagramNodeShape(element, item, raw, rawNode, catalog, includeText: false),
                });
                group.Children.Add(new PresentationElement
                {
                    Id = DiagramChildId(element.Id, "image", item.Node.Id),
                    Name = $"{item.Node.Id} image",
                    Image = BuildDiagramNodeImage(item, catalog),
                });
                group.Children.Add(new PresentationElement
                {
                    Id = DiagramChildId(element.Id, "label", item.Node.Id),
                    Name = $"{item.Node.Id} label",
                    Shape = BuildDiagramNodeLabel(element, item, rawNode, catalog),
                });
            }
            else
            {
                group.Children.Add(new PresentationElement
                {
                    Id = nodeId,
                    Name = DisplayName(null, null, item.Node.Id),
                    Shape = BuildDiagramNodeShape(element, item, raw, rawNode, catalog, includeText: true),
                });
            }
        }
        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static PresentationShape BuildDiagramNodeShape(
        PpjSmartArtElementModel element,
        DiagramLayoutNode item,
        JsonElement rawElement,
        JsonElement rawNode,
        Catalog catalog,
        bool includeText)
    {
        var geometry = Property(rawNode, "geometry") ?? Property(rawElement, "nodeGeometry");
        var geometryKind = geometry is null ? "preset" : geometry.Value.GetProperty("kind").GetString();
        var geometryName = geometryKind == "custom"
            ? "custom"
            : geometry is null ? "rect" : geometry.Value.GetProperty("preset").GetString()!;
        var shape = ShapeFrame(item.Frame, geometryName);
        var shapeStyle = catalog.ShapeStyle(item.Node.ShapeStyleRef ?? element.ShapeStyleRef);
        ApplyShapeStyle(shape, shapeStyle, null, catalog, item.Node.Id, includeText);
        if (geometry is { } value)
        {
            if (geometryKind == "custom") ApplyCustomGeometry(shape, value, item.Node.Id);
            else if (value.TryGetProperty("adjustments", out var adjustments))
                shape.PresetAdjustments.Add(adjustments.EnumerateArray().Select(item => item.GetInt32()));
        }
        if (includeText)
        {
            var textStyle = catalog.TextStyle(item.Node.StyleRef ?? element.TextStyleRef);
            shape.TextBody = BuildTextBody(rawNode.GetProperty("text"), textStyle, null, catalog);
            shape.Text = Flatten(shape.TextBody);
        }
        return shape;
    }

    private static PresentationShape BuildDiagramNodeLabel(
        PpjSmartArtElementModel element,
        DiagramLayoutNode item,
        JsonElement rawNode,
        Catalog catalog)
    {
        var frame = DiagramPictureLabelFrame(item.Frame);
        var shape = ShapeFrame(frame, "textbox");
        shape.LineStyle = "none";
        var textStyle = catalog.TextStyle(item.Node.StyleRef ?? element.TextStyleRef);
        shape.TextBody = BuildTextBody(rawNode.GetProperty("text"), textStyle, null, catalog);
        shape.Text = Flatten(shape.TextBody);
        return shape;
    }

    private static PresentationImage BuildDiagramNodeImage(DiagramLayoutNode item, Catalog catalog)
    {
        var frame = DiagramPictureImageFrame(item.Frame);
        var image = new PresentationImage
        {
            AssetId = catalog.NativeAssetId(item.Node.AssetId!),
            AltText = FlattenText(item.Node.Text),
            LeftEmu = Emu(frame.X),
            TopEmu = Emu(frame.Y),
            WidthEmu = Emu(frame.Width),
            HeightEmu = Emu(frame.Height),
        };
        if (catalog.AssetDimensions(item.Node.AssetId!) is { } dimensions)
            image.Crop = DiagramCoverCrop(frame.Width, frame.Height, dimensions.Width, dimensions.Height);
        return image;
    }

    private static PresentationImageCrop DiagramCoverCrop(
        double frameWidth,
        double frameHeight,
        double imageWidth,
        double imageHeight)
    {
        var crop = new PresentationImageCrop();
        var frameRatio = frameWidth / frameHeight;
        var imageRatio = imageWidth / imageHeight;
        if (imageRatio > frameRatio)
        {
            var edge = checked((int)Math.Round((1 - frameRatio / imageRatio) * 50_000));
            crop.LeftThousandthPercent = edge;
            crop.RightThousandthPercent = edge;
        }
        else if (imageRatio < frameRatio)
        {
            var edge = checked((int)Math.Round((1 - imageRatio / frameRatio) * 50_000));
            crop.TopThousandthPercent = edge;
            crop.BottomThousandthPercent = edge;
        }
        return crop;
    }

    private static PresentationConnector BuildDiagramConnector(
        PpjFrameModel from,
        PpjFrameModel to,
        string fromId,
        string toId,
        JsonElement style,
        Catalog catalog,
        string layout)
    {
        var horizontal = layout == "process" || Math.Abs((to.X + to.Width / 2) - (from.X + from.Width / 2)) >=
            Math.Abs((to.Y + to.Height / 2) - (from.Y + from.Height / 2));
        var forward = horizontal
            ? to.X + to.Width / 2 >= from.X + from.Width / 2
            : to.Y + to.Height / 2 >= from.Y + from.Height / 2;
        var startX = horizontal ? (forward ? from.X + from.Width : from.X) : from.X + from.Width / 2;
        var startY = horizontal ? from.Y + from.Height / 2 : (forward ? from.Y + from.Height : from.Y);
        var endX = horizontal ? (forward ? to.X : to.X + to.Width) : to.X + to.Width / 2;
        var endY = horizontal ? to.Y + to.Height / 2 : (forward ? to.Y : to.Y + to.Height);
        var connector = new PresentationConnector
        {
            ConnectorType = layout == "hierarchy" ? "elbow" : "straight",
            StartXEmu = Emu(startX),
            StartYEmu = Emu(startY),
            EndXEmu = Emu(endX),
            EndYEmu = Emu(endY),
            StartTargetId = fromId,
            EndTargetId = toId,
            StartConnectionSiteIndex = 0,
            EndConnectionSiteIndex = 0,
            StartArrow = Arrow(OptionalString(style, "startArrow")),
            EndArrow = Arrow(OptionalString(style, "endArrow")),
        };
        ApplyLine(connector, style.GetProperty("stroke"), catalog);
        return connector;
    }

    private static IReadOnlyList<DiagramEdge> DiagramEdges(
        PpjSmartArtElementModel element,
        IReadOnlyList<DiagramLayoutNode> layout)
    {
        if (layout.Count < 2) return [];
        var indexes = layout.Select((item, index) => (item.Node.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var explicitEdges = layout
            .Select((item, index) => item.Node.ParentId is null
                ? (DiagramEdge?)null
                : new DiagramEdge(indexes[item.Node.ParentId], index))
            .Where(item => item is not null)
            .Select(item => item!.Value)
            .ToArray();
        if (explicitEdges.Length > 0) return explicitEdges;
        if (element.Layout == "process")
            return Enumerable.Range(0, layout.Count - 1).Select(index => new DiagramEdge(index, index + 1)).ToArray();
        if (element.Layout == "cycle")
            return Enumerable.Range(0, layout.Count).Select(index => new DiagramEdge(index, (index + 1) % layout.Count)).ToArray();
        if (element.Layout == "relationship")
            return Enumerable.Range(1, layout.Count - 1).Select(index => new DiagramEdge(0, index)).ToArray();
        return [];
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramNodes(PpjSmartArtElementModel element) =>
        element.Layout switch
        {
            "list" => LayoutDiagramGrid(element, 1),
            "process" => LayoutDiagramGrid(element, element.Nodes.Count),
            "cycle" => LayoutDiagramRadial(element, centerFirst: false),
            "hierarchy" => LayoutDiagramHierarchy(element),
            "relationship" => LayoutDiagramRadial(element, centerFirst: true),
            "matrix" => LayoutDiagramGrid(element, (int)Math.Ceiling(Math.Sqrt(element.Nodes.Count))),
            "pyramid" => LayoutDiagramPyramid(element),
            "picture" => LayoutDiagramGrid(element, (int)Math.Ceiling(Math.Sqrt(element.Nodes.Count))),
            _ => throw Unsupported(element.Id, $"authored diagram layout {element.Layout} is not compiler-owned"),
        };

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramGrid(PpjSmartArtElementModel element, int columns)
    {
        columns = Math.Max(1, Math.Min(columns, element.Nodes.Count));
        var rows = (int)Math.Ceiling(element.Nodes.Count / (double)columns);
        var gap = DiagramGap(element.Frame, Math.Max(rows, columns));
        var width = (element.Frame.Width - gap * (columns - 1)) / columns;
        var height = (element.Frame.Height - gap * (rows - 1)) / rows;
        return element.Nodes.Select((node, index) => new DiagramLayoutNode(
            node,
            new PpjFrameModel(
                element.Frame.X + (index % columns) * (width + gap),
                element.Frame.Y + (index / columns) * (height + gap),
                width,
                height,
                0,
                false,
                false))).ToArray();
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramRadial(PpjSmartArtElementModel element, bool centerFirst)
    {
        if (element.Nodes.Count == 1)
            return [new DiagramLayoutNode(element.Nodes[0], new PpjFrameModel(
                element.Frame.X + element.Frame.Width * 0.3,
                element.Frame.Y + element.Frame.Height * 0.3,
                element.Frame.Width * 0.4,
                element.Frame.Height * 0.4,
                0,
                false,
                false))];
        var radialCount = centerFirst ? element.Nodes.Count - 1 : element.Nodes.Count;
        var nodeWidth = Math.Min(element.Frame.Width * 0.24, element.Frame.Width / Math.Max(2, Math.Ceiling(Math.Sqrt(radialCount))));
        var nodeHeight = Math.Min(element.Frame.Height * 0.24, element.Frame.Height / Math.Max(2, Math.Ceiling(Math.Sqrt(radialCount))));
        var centerX = element.Frame.X + element.Frame.Width / 2;
        var centerY = element.Frame.Y + element.Frame.Height / 2;
        var radiusX = Math.Max(0, (element.Frame.Width - nodeWidth) / 2);
        var radiusY = Math.Max(0, (element.Frame.Height - nodeHeight) / 2);
        var result = new List<DiagramLayoutNode>(element.Nodes.Count);
        if (centerFirst)
            result.Add(new DiagramLayoutNode(element.Nodes[0], new PpjFrameModel(
                centerX - nodeWidth / 2,
                centerY - nodeHeight / 2,
                nodeWidth,
                nodeHeight,
                0,
                false,
                false)));
        for (var index = 0; index < radialCount; index++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * index / radialCount;
            result.Add(new DiagramLayoutNode(
                element.Nodes[centerFirst ? index + 1 : index],
                new PpjFrameModel(
                    centerX + radiusX * Math.Cos(angle) - nodeWidth / 2,
                    centerY + radiusY * Math.Sin(angle) - nodeHeight / 2,
                    nodeWidth,
                    nodeHeight,
                    0,
                    false,
                    false)));
        }
        return result;
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramHierarchy(PpjSmartArtElementModel element)
    {
        var byId = element.Nodes.ToDictionary(node => node.Id, StringComparer.Ordinal);
        var depths = new Dictionary<string, int>(StringComparer.Ordinal);
        int depth(PpjSmartArtNodeModel node)
        {
            if (depths.TryGetValue(node.Id, out var found)) return found;
            var value = node.ParentId is null ? 0 : depth(byId[node.ParentId]) + 1;
            depths[node.Id] = value;
            return value;
        }
        foreach (var node in element.Nodes) depth(node);
        var levels = element.Nodes.GroupBy(node => depths[node.Id]).OrderBy(group => group.Key).ToArray();
        var verticalGap = DiagramGap(element.Frame, levels.Length);
        var rowHeight = (element.Frame.Height - verticalGap * (levels.Length - 1)) / levels.Length;
        var result = new List<DiagramLayoutNode>(element.Nodes.Count);
        foreach (var level in levels)
        {
            var nodes = level.ToArray();
            var horizontalGap = DiagramGap(element.Frame, nodes.Length);
            var width = (element.Frame.Width - horizontalGap * (nodes.Length - 1)) / nodes.Length;
            for (var index = 0; index < nodes.Length; index++)
                result.Add(new DiagramLayoutNode(nodes[index], new PpjFrameModel(
                    element.Frame.X + index * (width + horizontalGap),
                    element.Frame.Y + level.Key * (rowHeight + verticalGap),
                    width,
                    rowHeight,
                    0,
                    false,
                    false)));
        }
        return element.Nodes.Select(node => result.Single(item => item.Node.Id == node.Id)).ToArray();
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramPyramid(PpjSmartArtElementModel element)
    {
        var gap = DiagramGap(element.Frame, element.Nodes.Count);
        var height = (element.Frame.Height - gap * (element.Nodes.Count - 1)) / element.Nodes.Count;
        return element.Nodes.Select((node, index) =>
        {
            var scale = 0.45 + 0.55 * (index + 1) / element.Nodes.Count;
            var width = element.Frame.Width * scale;
            return new DiagramLayoutNode(node, new PpjFrameModel(
                element.Frame.X + (element.Frame.Width - width) / 2,
                element.Frame.Y + index * (height + gap),
                width,
                height,
                0,
                false,
                false));
        }).ToArray();
    }

    private static PpjFrameModel DiagramPictureImageFrame(PpjFrameModel frame) => new(
        frame.X,
        frame.Y,
        frame.Width,
        frame.Height * 0.7,
        0,
        false,
        false);

    private static PpjFrameModel DiagramPictureLabelFrame(PpjFrameModel frame) => new(
        frame.X,
        frame.Y + frame.Height * 0.7,
        frame.Width,
        frame.Height * 0.3,
        0,
        false,
        false);

    private static double DiagramGap(PpjFrameModel frame, int count) =>
        count <= 1 ? 0 : Math.Min(18, Math.Max(4, Math.Min(frame.Width, frame.Height) * 0.025));

    private static string DiagramChildId(string owner, string kind, string suffix)
    {
        var candidate = $"{owner}.{kind}.{suffix}";
        if (candidate.Length <= 255) return candidate;
        var ownerPrefix = owner.Length <= 180 ? owner : owner[..180];
        return $"{ownerPrefix}.{kind}.{Sha256(Encoding.UTF8.GetBytes(candidate))[..24]}";
    }

    private static string FlattenText(PpjTextContentModel text) => text.PlainText ??
        string.Join("\n", text.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    private sealed record DiagramLayoutNode(PpjSmartArtNodeModel Node, PpjFrameModel Frame);
    private readonly record struct DiagramEdge(int FromIndex, int ToIndex);

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
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
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
            var fillType = fillValue.GetProperty("type").GetString();
            if (fillType == "gradient")
            {
                target.GradientFill = BuildGradientFill(fillValue, color => catalog.Color(color));
            }
            else if (fillType == "image")
            {
                target.ImageFill = PpjImagePaintLowering.Build(
                    fillValue,
                    target.WidthEmu / EmuPerPoint,
                    target.HeightEmu / EmuPerPoint,
                    catalog.NativeAssetId,
                    catalog.AssetDimensions,
                    $"element {elementId} fill");
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
            target.Shadow = BuildShadow(shadowValue, catalog);
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
        var fillType = fill.GetProperty("type").GetString();
        if (fillType == "gradient")
        {
            target.GradientFill = BuildGradientFill(fill, color => catalog.Color(color));
            return;
        }
        if (fillType == "image")
        {
            target.ImageFill = PpjImagePaintLowering.Build(
                fill,
                target.WidthEmu / EmuPerPoint,
                target.HeightEmu / EmuPerPoint,
                catalog.NativeAssetId,
                catalog.AssetDimensions,
                "text fill");
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
                null,
                catalog);
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
                Property(paragraphJson, "style"),
                catalog);
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
        var paint = FirstTextPaint(inlineRun, inlineDefault, namedDefault);
        var shadow = FirstProperty(inlineRun, inlineDefault, namedDefault, "shadow");
        var highlight = FirstProperty(inlineRun, inlineDefault, namedDefault, "highlight");
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
        if (bold is { } boldValue) run.Bold = boldValue.GetBoolean();
        if (italic is { } italicValue) run.Italic = italicValue.GetBoolean();
        if (size is { } sizeValue) run.FontSizePoints = sizeValue.GetDouble();
        if (paint is { Kind: "color" } colorPaint)
        {
            var resolved = catalog.Color(colorPaint.Value);
            run.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) run.ColorOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else if (paint is { Kind: "gradient" } gradientPaint)
        {
            run.GradientFill = BuildGradientFill(gradientPaint.Value, color => catalog.Color(color));
        }
        if (shadow is { } shadowValue) run.Shadow = BuildShadow(shadowValue, catalog);
        if (highlight is { } highlightValue)
        {
            var resolved = catalog.Color(highlightValue);
            if (resolved.Alpha < 1)
                throw Unsupported("text", "highlight alpha is not part of the bounded DrawingML highlight profile");
            run.HighlightRgb = resolved.Rgb;
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
        if (language is { } languageValue) run.Language = languageValue.GetString()!;
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
        if (value.TryGetProperty("bold", out var bold)) output.Bold = bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic)) output.Italic = italic.GetBoolean();
        if (value.TryGetProperty("size", out var size)) output.FontSizePoints = size.GetDouble();
        if (value.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            output.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) output.ColorOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        if (value.TryGetProperty("gradient", out var gradient))
        {
            if (value.TryGetProperty("color", out _))
                throw new CodecException("ppj.text.paintConflict", "PPJ text style cannot declare both color and gradient.");
            output.GradientFill = BuildGradientFill(gradient, color => catalog.Color(color));
        }
        if (value.TryGetProperty("shadow", out var shadow)) output.Shadow = BuildShadow(shadow, catalog);
        if (value.TryGetProperty("highlight", out var highlight))
        {
            var resolved = catalog.Color(highlight);
            if (resolved.Alpha < 1)
                throw Unsupported("text", "highlight alpha is not part of the bounded DrawingML highlight profile");
            output.HighlightRgb = resolved.Rgb;
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
        if (value.TryGetProperty("language", out var language)) output.Language = language.GetString()!;
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
        JsonElement? direct,
        Catalog catalog)
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
        if (FirstProperty(direct, inline, named, "defaultText") is { } defaultText && defaultText.EnumerateObject().Any())
            target.DefaultRunProperties = BuildTextStyle(defaultText, catalog);
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
            if (bullet.TryGetProperty("color", out var bulletColor))
            {
                var color = catalog.Color(bulletColor);
                target.BulletColorRgb = color.Rgb;
                if (color.Alpha < 1) target.BulletColorOpacityThousandthPercent = Opacity(color.Alpha);
            }
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

    private static PresentationFrameTransform? BuildFrameTransform(PpjFrameModel frame)
    {
        if (frame.Rotation == 0 && !frame.FlipH && !frame.FlipV) return null;
        var transform = new PresentationFrameTransform();
        if (frame.Rotation != 0) transform.RotationAngle60000 = Angle(frame.Rotation);
        if (frame.FlipH) transform.FlipHorizontal = true;
        if (frame.FlipV) transform.FlipVertical = true;
        return transform;
    }

    private static void RejectUnsupportedFrameTransform(
        string elementId,
        PpjFrameModel frame,
        string elementKind)
    {
        if (frame.Rotation != 0 || frame.FlipH || frame.FlipV)
            throw Unsupported(elementId, $"{elementKind} frame rotation and flips are not yet compiler-owned");
    }

    private static PresentationBackground BuildBackground(
        JsonElement fill,
        Catalog catalog,
        double canvasWidth,
        double canvasHeight)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "solid")
        {
            var color = catalog.Color(fill.GetProperty("color"));
            var opacity = OptionalDouble(fill, "opacity") ?? color.Alpha;
            var background = new PresentationBackground { Solid = true, ColorRgb = color.Rgb };
            if (opacity < 1) background.OpacityThousandthPercent = Opacity(opacity);
            return background;
        }
        if (type == "gradient") return new PresentationBackground
        {
            GradientFill = BuildGradientFill(fill, color => catalog.Color(color)),
        };
        if (type == "image")
        {
            return new PresentationBackground
            {
                ImagePaint = PpjImagePaintLowering.Build(
                    fill,
                    canvasWidth,
                    canvasHeight,
                    catalog.NativeAssetId,
                    catalog.AssetDimensions,
                    "background"),
            };
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
        var paint = PpjImagePaintLowering.Build(
            raw,
            element.Frame.Width,
            element.Frame.Height,
            catalog.NativeAssetId,
            catalog.AssetDimensions,
            $"element {element.Id}");
        target.Crop = paint.Crop;
        target.Tiled = paint.Mode == PresentationImagePaint.Types.Mode.Tile;
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
            chart.TitleTextStyle = BuildChartTextStyle(titleTextStyle, catalog);
        }
        if (FirstProperty(inline, named, "legendTextStyle") is { } legendTextStyle)
        {
            if (!chart.HasLegend)
                throw Unsupported(elementId, "legendTextStyle requires a visible legend");
            chart.LegendTextStyle = BuildChartTextStyle(legendTextStyle, catalog);
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
            if (dataLabels.TryGetProperty("textStyle", out var textStyle))
                chart.DataLabels.TextStyle = BuildChartTextStyle(textStyle, catalog);
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
        => BuildChartFill(fill, catalog, subject);

    private static SpreadsheetChartSurfaceFill BuildChartFill(JsonElement fill, Catalog catalog, string subject)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return new SpreadsheetChartSurfaceFill { NoFill = true };
        if (type == "gradient") return new SpreadsheetChartSurfaceFill
        {
            GradientFill = BuildGradientFill(fill, color => catalog.Color(color)),
        };
        if (type != "solid") throw Unsupported(subject, "chart paint supports only none, solid or bounded gradient fills");
        var resolved = FillColor(fill, catalog) ?? throw new InvalidOperationException("Solid PPJ fill unexpectedly resolved to none.");
        var output = new SpreadsheetChartSurfaceFill { SolidRgb = resolved.Rgb };
        if (resolved.Opacity < 1) output.OpacityThousandthPercent = Opacity(resolved.Opacity);
        return output;
    }

    private static SpreadsheetChartTextStyleArtifact BuildChartTextStyle(JsonElement source, Catalog catalog)
    {
        var output = new SpreadsheetChartTextStyleArtifact();
        if (source.TryGetProperty("fontSize", out var fontSize)) output.FontSizePoints = fontSize.GetDouble();
        if (source.TryGetProperty("fontFamily", out var fontFamily)) output.FontFamily = fontFamily.GetString()!;
        if (source.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) output.FontFamilyEastAsia = eastAsia.GetString()!;
        if (source.TryGetProperty("bold", out var bold)) output.Bold = bold.GetBoolean();
        if (source.TryGetProperty("italic", out var italic)) output.Italic = italic.GetBoolean();
        if (source.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            output.ColorRgb = resolved.Rgb;
            if (resolved.Alpha != 1) output.OpacityThousandthPercent = Opacity(resolved.Alpha);
        }
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
            var hasCurrentPoint = false;
            var hasSubpathStart = false;
            foreach (var sourceCommand in sourcePath.GetProperty("commands").EnumerateArray())
            {
                var command = new PresentationCustomGeometryCommand();
                switch (sourceCommand.GetProperty("op").GetString())
                {
                    case "moveTo":
                        command.MoveTo = CustomPoint(sourceCommand, originX, originY, "x", "y");
                        hasCurrentPoint = true;
                        hasSubpathStart = true;
                        break;
                    case "lineTo":
                        command.LineTo = CustomPoint(sourceCommand, originX, originY, "x", "y");
                        hasCurrentPoint = true;
                        break;
                    case "quadraticTo":
                        command.QuadraticBezierTo = new PresentationCustomGeometryQuadraticBezier
                        {
                            Control = CustomPoint(sourceCommand, originX, originY, "x1", "y1"),
                            End = CustomPoint(sourceCommand, originX, originY, "x", "y"),
                        };
                        hasCurrentPoint = true;
                        break;
                    case "cubicTo":
                        command.CubicBezierTo = new PresentationCustomGeometryCubicBezier
                        {
                            Control1 = CustomPoint(sourceCommand, originX, originY, "x1", "y1"),
                            Control2 = CustomPoint(sourceCommand, originX, originY, "x2", "y2"),
                            End = CustomPoint(sourceCommand, originX, originY, "x", "y"),
                        };
                        hasCurrentPoint = true;
                        break;
                    case "arcTo":
                        if (!hasCurrentPoint)
                            throw new CodecException(
                                "ppj.geometry.arcCurrentPoint",
                                $"PPJ custom geometry {elementId} has an arc without an established current point.");
                        command.ArcTo = new PresentationCustomGeometryArc
                        {
                            WidthRadius = CustomPathCoordinate(sourceCommand.GetProperty("radiusX").GetDouble()),
                            HeightRadius = CustomPathCoordinate(sourceCommand.GetProperty("radiusY").GetDouble()),
                            StartAngle = Angle(sourceCommand.GetProperty("startAngle").GetDouble()),
                            SweepAngle = Angle(sourceCommand.GetProperty("sweepAngle").GetDouble()),
                        };
                        hasCurrentPoint = true;
                        break;
                    case "close":
                        command.Close = true;
                        hasCurrentPoint = hasSubpathStart;
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
        "centerTitle" or "centered-title" => "ctrTitle",
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
        "radar" => SpreadsheetChartType.Radar,
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

    private static bool? OptionalBoolean(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) ? property.GetBoolean() : null;

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

    private static PresentationShadow BuildShadow(JsonElement value, Catalog catalog)
    {
        var colorValue = value.GetProperty("color");
        var schemeToken = colorValue.ValueKind == JsonValueKind.Object && colorValue.TryGetProperty("token", out var token) &&
                          PptxColor.TrySchemeToken(token.GetString() ?? string.Empty, out var recognizedScheme)
            ? recognizedScheme
            : null;
        var color = schemeToken is null
            ? catalog.Color(colorValue)
            : (Rgb: string.Empty, Alpha: colorValue.TryGetProperty("alpha", out var alpha) ? alpha.GetDouble() : 1d);
        var degrees = value.GetProperty("angle").GetDouble();
        var normalized = ((degrees % 360) + 360) % 360;
        var output = new PresentationShadow
        {
            ColorRgb = color.Rgb,
            BlurRadiusEmu = Emu(value.GetProperty("blur").GetDouble()),
            DistanceEmu = Emu(value.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = Angle(normalized),
            OpacityThousandthPercent = Opacity(OptionalDouble(value, "opacity") ?? color.Alpha),
        };
        if (schemeToken is not null)
            output.ColorScheme = schemeToken;
        if (value.TryGetProperty("alignment", out var alignment)) output.Alignment = alignment.GetString()!;
        if (value.TryGetProperty("rotateWithShape", out var rotateWithShape)) output.RotateWithShape = rotateWithShape.GetBoolean();
        return output;
    }

    private static (string Kind, JsonElement Value)? FirstTextPaint(params JsonElement?[] layers)
    {
        foreach (var layer in layers)
        {
            if (layer is not { ValueKind: JsonValueKind.Object } value) continue;
            var hasColor = value.TryGetProperty("color", out var color);
            var hasGradient = value.TryGetProperty("gradient", out var gradient);
            if (hasColor && hasGradient)
                throw new CodecException("ppj.text.paintConflict", "PPJ text style cannot declare both color and gradient.");
            if (hasColor) return ("color", color);
            if (hasGradient) return ("gradient", gradient);
        }
        return null;
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
