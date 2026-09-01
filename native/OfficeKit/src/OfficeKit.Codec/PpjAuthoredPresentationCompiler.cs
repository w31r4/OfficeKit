using System.Buffers;
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
internal static partial class PpjAuthoredPresentationCompiler
{
    private const double EmuPerPoint = 12_700d;
    private const double CustomPathUnitsPerPoint = 1_000d;

    internal static PpjCompileResult ValidateOnly(
        PresentationProgramRequest request,
        PpjValidationResult validation)
    {
        var program = validation.Program!;
        var assets = ValidateAssets(program, request.Assets);
        _ = new Catalog(program.Root, assets);
        var receipt = new PresentationProgramResult
        {
            ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            SourceBound = false,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        if (request.IncludeNodeMap)
            receipt.NodeMapJson = UnsafeByteOperations.UnsafeWrap(validation.Expansion.NodeMapJson);
        receipt.Assets.Add(assets.Select(asset => asset.Clone()));
        return new([], receipt, []);
    }

    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        EffectiveCodecLimits limits,
        PpjValidationResult validation)
    {
        var program = validation.Program!;
        if (program.Source is not null)
            throw new CodecException(
                "ppj.sourceBoundCompileRequired",
                "A source-bound PPJ must be compiled against its exact validated source PPTX.",
                "$.source");

        var assets = ValidateAssets(program, request.Assets);
        var catalog = new Catalog(program.Root, assets);
        var plan = new AuthoredSourceFreeBuildPlan(program, validation.Expansion!, catalog);
        var originalProgramJson = request.ProgramJson.ToByteArray();
        var exported = PptxCodec.ExportSourceFree(
            plan,
            assets,
            limits,
            parts => PpjEmbeddedProgramCodec.AddToSourceFreePackage(
                parts,
                originalProgramJson,
                validation,
                plan.NativeBindings,
                assets));
        var file = exported.File;
        PpjEmbeddedProgramCodec.ValidateEmbeddedOutput(file, limits);
        var fileSha256 = Sha256(file);
        var receipt = new PresentationProgramResult
        {
            ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            OutputSha256 = fileSha256,
            SourceBound = false,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        if (request.IncludeNodeMap)
            receipt.NodeMapJson = UnsafeByteOperations.UnsafeWrap(validation.Expansion.NodeMapJson);
        receipt.Assets.Add(assets.Select(asset => asset.Clone()));
        receipt.ChangedNodeIds.Add(validation.Expansion.Nodes.Select(node => node.Id));
        return new(file, receipt, exported.Diagnostics);
    }

    internal static PresentationElement BuildSourceBoundOverlayElement(
        PpjProgramModel program,
        PpjElementModel element)
    {
        if (element.NativeRef is not null)
            throw Unsupported(element.Id, "a new source-bound overlay cannot carry nativeRef authority");
        if (element is PpjImageElementModel { SvgAssetId: not null })
            throw Unsupported(element.Id, "paired SVG fallback images are outside the bounded source overlay relationship profile");
        var output = BuildElement(element, element.Raw, new Catalog(program.Root));
        if (!output.HasHidden) output.Hidden = false;
        if (!output.HasLocked) output.Locked = false;
        if (PptxCodec.BoundedAuthoredOverlayViolation(output) is { } violation)
            throw Unsupported(element.Id, violation);
        return output;
    }

    private sealed class AuthoredSourceFreeBuildPlan : IPptxSourceFreeBuildPlan
    {
        private readonly PpjProgramModel _program;
        private readonly Catalog _catalog;
        private readonly IReadOnlyDictionary<string, PpjExpandedPageModel> _expandedByPage;
        private readonly HashSet<string> _semanticIds;
        private readonly List<PptxNativeBinding> _nativeBindings = [];

        internal AuthoredSourceFreeBuildPlan(
            PpjProgramModel program,
            PpjExpansionResult expansion,
            Catalog catalog)
        {
            _program = program;
            _catalog = catalog;
            _expandedByPage = expansion.Pages.ToDictionary(page => page.Id, StringComparer.Ordinal);
            _semanticIds = expansion.Nodes.Select(node => node.Id).ToHashSet(StringComparer.Ordinal);
            Presentation = new PresentationArtifact
            {
                Id = program.Meta.Id,
                Name = program.Meta.Title,
                SlideWidthEmu = Emu(program.Design.Width),
                SlideHeightEmu = Emu(program.Design.Height),
                AuthoredTheme = catalog.Theme,
            };

            AddMasterLayoutState(Presentation, program, catalog);

            foreach (var page in program.Pages)
            {
                var slide = new PresentationSlide
                {
                    Id = page.Id,
                    Name = DisplayName(page.Name, page.Role, page.Id),
                    LayoutId = page.LayoutId ?? string.Empty,
                };
                if (page.Raw.TryGetProperty("hidden", out var hidden)) slide.Hidden = hidden.GetBoolean();
                if (page.Raw.TryGetProperty("notes", out var notes))
                    slide.SpeakerNotes = BuildNotes(notes, catalog);
                Presentation.Slides.Add(slide);
            }

            AddSections(Presentation, program);
            AddCustomShows(Presentation, program);
            AddComments(Presentation, program);
        }

        public PresentationArtifact Presentation { get; }
        internal IReadOnlyList<PptxNativeBinding> NativeBindings => _nativeBindings;

        public PresentationSlide MaterializeSlide(int pageIndex, PresentationSlide? previousSlide)
        {
            var page = _program.Pages[pageIndex];
            var expanded = _expandedByPage[page.Id];
            var slide = Presentation.Slides[pageIndex].Clone();
            if (page.Raw.TryGetProperty("background", out var background))
                slide.Background = BuildBackground(background, _catalog, _program.Design.Width, _program.Design.Height);
            for (var elementIndex = 0; elementIndex < expanded.Elements.Count; elementIndex++)
                slide.Elements.Add(BuildElement(expanded.Elements[elementIndex], expanded.ElementJson[elementIndex], _catalog));
            var loweredElements = WalkPresentation(slide.Elements)
                .ToDictionary(element => element.Id, StringComparer.Ordinal);
            foreach (var animation in page.Animations)
            {
                if (animation.ChartBuild is not null &&
                    loweredElements.TryGetValue(animation.TargetId, out var loweredTarget) &&
                    loweredTarget.ContentCase != PresentationElement.ContentOneofCase.Chart)
                    throw Unsupported(animation.TargetId, "chartBuild requires a native ChartPart target; vector-lowered charts support whole-object animation only");
                slide.Animations.Add(BuildAnimation(animation, expanded.Elements));
            }
            ApplyTransition(slide, page, previousSlide);
            return slide;
        }

        public void RecordNativeBindings(
            int pageIndex,
            PresentationSlide slide,
            IReadOnlyList<PresentationElement> flattenedElements)
        {
            for (var index = 0; index < flattenedElements.Count; index++)
            {
                var element = flattenedElements[index];
                if (!_semanticIds.Contains(element.Id)) continue;
                _nativeBindings.Add(new PptxNativeBinding(
                    slide.Id,
                    element.Id,
                    element.ContentCase.ToString(),
                    $"ppt/slides/slide{pageIndex + 1}.xml",
                    checked((uint)(index + 2))));
            }
        }
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
        var output = new PresentationElement
        {
            Id = element.Id,
            Name = DisplayName(element.Name, element.Role, element.Id),
        };
        if (element.Hidden is { } hidden) output.Hidden = hidden;
        if (element.Locked is { } locked) output.Locked = locked;
        switch (element)
        {
            case PpjTextElementModel:
                output.Shape = BuildTextShape(element, raw, catalog, "textbox");
                break;
            case PpjShapeElementModel shape:
                output.Shape = BuildShape(shape, raw, catalog);
                break;
            case PpjIconElementModel icon:
                output.Shape = BuildIcon(icon, raw, catalog);
                break;
            case PpjImageElementModel image:
                output.Image = BuildImage(image, raw, catalog);
                break;
            case PpjChartElementModel chart:
                ChartCompiler.BuildInto(output, chart, raw, catalog);
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
            case PpjMediaElementModel media:
                output.Media = BuildMedia(media, catalog);
                break;
            case PpjSmartArtElementModel { Mode: "authored" } diagram:
                output.Diagram = BuildAuthoredNativeDiagram(diagram, raw, catalog);
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

    internal static PresentationTextBody BuildChartTitleBody(
        PpjProgramModel program,
        PpjChartElementModel chart) => ChartCompiler.BuildTitleBody(program, chart);

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
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id);
        if (raw.TryGetProperty("text", out var text))
        {
            shape.TextBody = BuildTextBody(text, null, Property(raw, "textStyle"), catalog);
            shape.Text = Flatten(shape.TextBody);
        }
        if (FirstProperty(inlineStyle, namedStyle, "opacity") is { } opacity)
            ApplyCompoundShapeOpacity(shape, opacity.GetDouble(), element.Id);
        if (geometry == "custom") ApplyCustomGeometry(shape, raw.GetProperty("geometry"), element.Id);
        else shape.PresetAdjustments.Add(element.GeometryAdjustments);
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static PresentationShape BuildIcon(PpjIconElementModel element, JsonElement raw, Catalog catalog)
    {
        var definition = PpjIconCatalog.Resolve(element.IconName);
        var shape = ShapeFrame(element.Frame, "custom");
        shape.CatalogIconName = element.IconName;
        var namedStyle = catalog.ShapeStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        if (FirstProperty(inlineStyle, namedStyle, "fill") is null) shape.FillRgb = "000000";
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id);
        if (FirstProperty(inlineStyle, namedStyle, "opacity") is { } opacity)
            ApplyCompoundShapeOpacity(shape, opacity.GetDouble(), element.Id);
        ApplyIconGeometry(shape, definition, element.Frame, element.Id);
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static void ApplyIconGeometry(
        PresentationShape target,
        PpjIconDefinition definition,
        PpjFrameModel frame,
        string elementId)
    {
        var pathWidth = CustomPathCoordinate(frame.Width);
        var pathHeight = CustomPathCoordinate(frame.Height);
        var scale = Math.Min(pathWidth / definition.Width, pathHeight / definition.Height);
        var offsetX = (pathWidth - definition.Width * scale) / 2d;
        var offsetY = (pathHeight - definition.Height * scale) / 2d;
        var path = new PresentationCustomGeometryPath
        {
            Width = pathWidth,
            Height = pathHeight,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
        };
        foreach (var source in definition.Commands)
        {
            PresentationCustomGeometryPoint Point(int offset) => new()
            {
                X = IconCoordinate(offsetX + source.Values[offset] * scale, elementId),
                Y = IconCoordinate(offsetY + source.Values[offset + 1] * scale, elementId),
            };
            var command = source.Operation switch
            {
                'M' => new PresentationCustomGeometryCommand { MoveTo = Point(0) },
                'L' => new PresentationCustomGeometryCommand { LineTo = Point(0) },
                'C' => new PresentationCustomGeometryCommand
                {
                    CubicBezierTo = new PresentationCustomGeometryCubicBezier
                    {
                        Control1 = Point(0),
                        Control2 = Point(2),
                        End = Point(4),
                    },
                },
                'Z' => new PresentationCustomGeometryCommand { Close = true },
                _ => throw Unsupported(elementId, "icon catalog contains an unsupported path command"),
            };
            path.Commands.Add(command);
        }
        target.CustomPaths.Add(path);
    }

    private static long IconCoordinate(double value, string elementId)
    {
        if (!double.IsFinite(value) || value < -int.MaxValue || value > int.MaxValue)
            throw Unsupported(elementId, "icon geometry exceeds the native coordinate range");
        return checked((long)Math.Round(value, MidpointRounding.AwayFromZero));
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

    private static PresentationMedia BuildMedia(PpjMediaElementModel element, Catalog catalog)
    {
        var media = new PresentationMedia
        {
            MediaType = element.MediaType,
            AssetId = catalog.NativeAssetId(element.AssetId),
            PosterAssetId = catalog.NativeAssetId(element.PosterAssetId),
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Accessibility = Accessibility(element.Accessibility),
        };
        if (element.Frame.Rotation != 0 || element.Frame.FlipH || element.Frame.FlipV)
        {
            media.Transform = new PresentationImageTransform();
            if (element.Frame.Rotation != 0) media.Transform.RotationAngle60000 = Angle(element.Frame.Rotation);
            if (element.Frame.FlipH) media.Transform.FlipHorizontal = true;
            if (element.Frame.FlipV) media.Transform.FlipVertical = true;
        }
        if (element.StartAtMs is { } startAtMs) media.StartAtMs = startAtMs;
        if (element.EndAtMs is { } endAtMs) media.EndAtMs = endAtMs;
        if (element.Loop is { } loop) media.Loop = loop;
        if (element.Mute is { } mute) media.Mute = mute;
        return media;
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
        var headerTextStyle = FirstProperty(inlineStyle, namedStyle, "headerTextStyle");
        var headerCellFill = FirstProperty(inlineStyle, namedStyle, "headerCellFill");
        var cellStyle = FirstProperty(inlineStyle, namedStyle, "cellStyle");
        var bodyStyles = FirstProperty(inlineStyle, namedStyle, "bodyStyles");
        var firstRowStyle = FirstProperty(inlineStyle, namedStyle, "firstRowStyle");
        var lastRowStyle = FirstProperty(inlineStyle, namedStyle, "lastRowStyle");
        var firstColumnStyle = FirstProperty(inlineStyle, namedStyle, "firstColumnStyle");
        var lastColumnStyle = FirstProperty(inlineStyle, namedStyle, "lastColumnStyle");
        var rowOverColumn = FirstProperty(inlineStyle, namedStyle, "rowOverColumn")?.GetBoolean() ?? true;
        var declaredHeaderRows = FirstProperty(inlineStyle, namedStyle, "headerRows");
        var headerRowCount = declaredHeaderRows?.GetInt32() ?? 0;
        if (headerRowCount > element.Rows.Count)
            throw Unsupported(element.Id, $"table headerRows {headerRowCount} exceeds the {element.Rows.Count}-row physical grid");
        if (headerRowCount == 0 && (headerTextStyle is not null || headerCellFill is not null))
            throw Unsupported(element.Id, "headerTextStyle and headerCellFill require headerRows greater than zero");
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
            var headerRow = rowIndex < headerRowCount;
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
                var bodyStyle = rowIndex > 0 && rowIndex < element.Rows.Count - 1 && bodyStyles is { } declaredBodyStyles
                    ? declaredBodyStyles[checked((rowIndex - 1) % declaredBodyStyles.GetArrayLength())]
                    : (JsonElement?)null;
                var rowStyle = MergeJsonObjects(
                    rowIndex == 0 ? firstRowStyle : null,
                    rowIndex == element.Rows.Count - 1 ? lastRowStyle : null);
                var columnStyle = MergeJsonObjects(
                    cursor == 0 ? firstColumnStyle : null,
                    cursor + cell.ColumnSpan == element.Columns.Count ? lastColumnStyle : null);
                var structuralStyle = rowOverColumn
                    ? MergeJsonObjects(columnStyle, rowStyle)
                    : MergeJsonObjects(rowStyle, columnStyle);
                var baseTextStyle = MergeJsonObjects(defaultTextStyle, Property(cellStyle, "textStyle"));
                var bodyTextStyle = MergeJsonObjects(
                    Property(bodyStyle, "textStyle"),
                    headerRow ? headerTextStyle : null);
                var directTextStyle = MergeJsonObjects(
                    Property(structuralStyle, "textStyle"),
                    Property(rawCell, "textStyle"));
                targetCell.TextBody = BuildTextBody(
                    rawCell.GetProperty("text"),
                    baseTextStyle,
                    bodyTextStyle,
                    directTextStyle,
                    catalog);
                targetCell.Text = PptxTextCodec.Flatten(targetCell.TextBody);
                if ((Property(rawCell, "fill") ??
                     Property(structuralStyle, "fill") ??
                     (headerRow ? headerCellFill : null) ??
                     Property(bodyStyle, "fill") ??
                     Property(cellStyle, "fill") ??
                     defaultCellFill) is { } cellFill)
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
                if (MergeJsonObjects(
                        Property(cellStyle, "borders"),
                        Property(bodyStyle, "borders"),
                        Property(structuralStyle, "borders"),
                        Property(rawCell, "borders")) is { } borders)
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
        if (declaredHeaderRows is not null) table.FirstRow = headerRowCount > 0;
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
        Catalog catalog,
        PpjSmartArtDefinition definition)
    {
        var resolvedLayout = definition.LayoutProfile;
        var nodeJson = raw.GetProperty("nodes").EnumerateArray().ToDictionary(
            node => node.GetProperty("id").GetString()!,
            node => node,
            StringComparer.Ordinal);
        var layout = LayoutDiagramNodes(element, definition.Execution);
        if (layout.Any(item => item.Frame.Width < 8 || item.Frame.Height < 8) ||
            resolvedLayout == "picture" && layout.Any(item =>
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
            foreach (var edge in DiagramEdges(element, layout, resolvedLayout))
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
                    resolvedLayout);
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
            if (resolvedLayout == "picture")
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

    private static PresentationDiagram BuildAuthoredNativeDiagram(
        PpjSmartArtElementModel element,
        JsonElement raw,
        Catalog catalog)
    {
        var definition = element.Layout is not null
            ? PpjSmartArtDefinitionCodec.BuiltIn(element.Layout)
            : catalog.SmartArtDefinition(element.DefinitionAssetId!);
        var resolvedLayout = definition.LayoutProfile;
        ValidateResolvedDiagramProfile(element, raw, definition);
        var rawNodes = raw.GetProperty("nodes").EnumerateArray().ToDictionary(
            node => node.GetProperty("id").GetString()!,
            node => node,
            StringComparer.Ordinal);
        var diagram = new PresentationDiagram
        {
            Layout = resolvedLayout,
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Drawing = BuildAuthoredDiagram(element, raw, catalog, definition),
            DrawingCacheVerified = true,
            DefinitionAssetId = element.DefinitionAssetId is null
                ? string.Empty
                : catalog.NativeAssetId(element.DefinitionAssetId),
        };
        ApplyAccessibility(diagram, element.Accessibility);
        foreach (var node in element.Nodes)
        {
            var rawNode = rawNodes[node.Id];
            var textStyle = catalog.TextStyle(node.StyleRef ?? element.TextStyleRef);
            diagram.Nodes.Add(new PresentationDiagramNode
            {
                Id = node.Id,
                TextBody = BuildTextBody(rawNode.GetProperty("text"), textStyle, null, catalog),
                AssetId = node.AssetId is null ? string.Empty : catalog.NativeAssetId(node.AssetId),
            });
        }
        diagram.Connections.Add(ResolvedDiagramConnections(element, resolvedLayout));
        return diagram;
    }

    private static void ValidateResolvedDiagramProfile(
        PpjSmartArtElementModel element,
        JsonElement raw,
        PpjSmartArtDefinition definition)
    {
        var resolvedLayout = definition.LayoutProfile;
        var placement = definition.Execution.Placement;
        var hasParentConnections = element.Connections.Any(connection => connection.Role == "parent");
        if (hasParentConnections && resolvedLayout is "list" or "process" or "cycle" or "matrix" or "pyramid" or "picture")
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot declare parent connections");
        if (resolvedLayout == "hierarchy" && element.Nodes.Count > 1 && !hasParentConnections)
            throw Unsupported(element.Id, "authored hierarchy diagrams require parent connections");

        var connected = element.Nodes.Count > 1 && resolvedLayout is "process" or "cycle" or "hierarchy" or "relationship";
        var hasConnector = raw.TryGetProperty("connector", out _);
        if (connected && !hasConnector)
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams require explicit connector styling");
        if (!connected && hasConnector)
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams do not emit connector edges for this node set");

        if (resolvedLayout == "picture" && element.Nodes.Any(node => node.AssetId is null))
            throw Unsupported(element.Id, "every authored picture-diagram node requires an image asset");
        if (resolvedLayout != "picture" && element.Nodes.Any(node => node.AssetId is not null))
            throw Unsupported(element.Id, "diagram node assets are only valid for the picture layout");
        if ((resolvedLayout == "picture") != (placement == "square-grid-picture"))
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot execute the {placement} placement");
        if ((resolvedLayout == "hierarchy") != (placement == "depth-levels"))
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot execute the {placement} placement");
        if ((resolvedLayout == "pyramid") != (placement == "stacked-width"))
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot execute the {placement} placement");
        if (resolvedLayout is "cycle" or "relationship" && placement is not ("radial" or "center-radial"))
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot execute the {placement} placement");
        if (resolvedLayout is "list" or "process" or "matrix" && placement is not ("grid" or "horizontal-grid" or "square-grid"))
            throw Unsupported(element.Id, $"authored {resolvedLayout} diagrams cannot execute the {placement} placement");
        if (definition.Execution.Reverse && placement == "depth-levels")
            throw Unsupported(element.Id, "reverse is not executable for a hierarchy placement");
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
        ApplyShapeStyle(shape, shapeStyle, null, catalog, item.Node.Id);
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
        if (FirstProperty(null, shapeStyle, "opacity") is { } opacity)
            ApplyCompoundShapeOpacity(shape, opacity.GetDouble(), item.Node.Id);
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
        IReadOnlyList<DiagramLayoutNode> layout,
        string resolvedLayout)
    {
        if (layout.Count < 2) return [];
        var indexes = layout.Select((item, index) => (item.Node.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        return ResolvedDiagramConnections(element, resolvedLayout)
            .Select(connection => new DiagramEdge(indexes[connection.FromId], indexes[connection.ToId]))
            .ToArray();
    }

    private static IReadOnlyList<PresentationDiagramConnection> ResolvedDiagramConnections(
        PpjSmartArtElementModel element,
        string resolvedLayout)
    {
        if (element.Connections.Count > 0)
            return element.Connections.Select(connection => new PresentationDiagramConnection
            {
                Id = connection.Id,
                FromId = connection.FromId,
                ToId = connection.ToId,
                Role = connection.Role,
                Order = connection.Order,
            }).ToArray();
        if (resolvedLayout == "process")
            return Enumerable.Range(0, element.Nodes.Count - 1).Select(index => new PresentationDiagramConnection
            {
                Id = $"sequence-{element.Nodes[index].Id}-{element.Nodes[index + 1].Id}",
                FromId = element.Nodes[index].Id,
                ToId = element.Nodes[index + 1].Id,
                Role = "sequence",
                Order = checked((uint)index),
            }).ToArray();
        if (resolvedLayout == "cycle")
            return Enumerable.Range(0, element.Nodes.Count).Select(index => new PresentationDiagramConnection
            {
                Id = $"sequence-{element.Nodes[index].Id}-{element.Nodes[(index + 1) % element.Nodes.Count].Id}",
                FromId = element.Nodes[index].Id,
                ToId = element.Nodes[(index + 1) % element.Nodes.Count].Id,
                Role = "sequence",
                Order = checked((uint)index),
            }).ToArray();
        if (resolvedLayout == "relationship")
            return Enumerable.Range(1, element.Nodes.Count - 1).Select(index => new PresentationDiagramConnection
            {
                Id = $"association-{element.Nodes[0].Id}-{element.Nodes[index].Id}",
                FromId = element.Nodes[0].Id,
                ToId = element.Nodes[index].Id,
                Role = "association",
                Order = checked((uint)(index - 1)),
            }).ToArray();
        return [];
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramNodes(
        PpjSmartArtElementModel element,
        PpjSmartArtExecutionPlan execution)
    {
        var layout = execution.Placement switch
        {
            "grid" => LayoutDiagramGrid(element, execution.Columns ?? 1, execution.GapPoints),
            "horizontal-grid" => LayoutDiagramGrid(element, execution.Columns ?? element.Nodes.Count, execution.GapPoints),
            "radial" => LayoutDiagramRadial(element, centerFirst: false),
            "depth-levels" => LayoutDiagramHierarchy(element, execution.GapPoints),
            "center-radial" => LayoutDiagramRadial(element, centerFirst: true),
            "square-grid" or "square-grid-picture" => LayoutDiagramGrid(
                element,
                execution.Columns ?? (int)Math.Ceiling(Math.Sqrt(element.Nodes.Count)),
                execution.GapPoints),
            "stacked-width" => LayoutDiagramPyramid(element, execution.GapPoints),
            _ => throw Unsupported(element.Id, $"authored diagram placement {execution.Placement} is not compiler-owned"),
        };
        if (!execution.Reverse) return layout;
        var reversed = element.Nodes.Reverse().ToArray();
        return layout.Select((item, index) => new DiagramLayoutNode(reversed[index], item.Frame)).ToArray();
    }

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramGrid(
        PpjSmartArtElementModel element,
        int columns,
        double? gapPoints)
    {
        columns = Math.Max(1, Math.Min(columns, element.Nodes.Count));
        var rows = (int)Math.Ceiling(element.Nodes.Count / (double)columns);
        var gap = DiagramGap(element.Frame, Math.Max(rows, columns), gapPoints);
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

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramHierarchy(
        PpjSmartArtElementModel element,
        double? gapPoints)
    {
        var byId = element.Nodes.ToDictionary(node => node.Id, StringComparer.Ordinal);
        var parentIds = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var connection in element.Connections.Where(connection => connection.Role == "parent"))
            parentIds[connection.ToId] = connection.FromId;
        var depths = new Dictionary<string, int>(StringComparer.Ordinal);
        int depth(PpjSmartArtNodeModel node)
        {
            if (depths.TryGetValue(node.Id, out var found)) return found;
            var value = parentIds.TryGetValue(node.Id, out var parentId) ? depth(byId[parentId]) + 1 : 0;
            depths[node.Id] = value;
            return value;
        }
        foreach (var node in element.Nodes) depth(node);
        var levels = element.Nodes.GroupBy(node => depths[node.Id]).OrderBy(group => group.Key).ToArray();
        var verticalGap = DiagramGap(element.Frame, levels.Length, gapPoints);
        var rowHeight = (element.Frame.Height - verticalGap * (levels.Length - 1)) / levels.Length;
        var result = new List<DiagramLayoutNode>(element.Nodes.Count);
        foreach (var level in levels)
        {
            var nodes = level.ToArray();
            var horizontalGap = DiagramGap(element.Frame, nodes.Length, gapPoints);
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

    private static IReadOnlyList<DiagramLayoutNode> LayoutDiagramPyramid(
        PpjSmartArtElementModel element,
        double? gapPoints)
    {
        var gap = DiagramGap(element.Frame, element.Nodes.Count, gapPoints);
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

    private static double DiagramGap(PpjFrameModel frame, int count, double? explicitGap = null) =>
        count <= 1 ? 0 : explicitGap ?? Math.Min(18, Math.Max(4, Math.Min(frame.Width, frame.Height) * 0.025));

    private static string DiagramChildId(string owner, string kind, string suffix)
    {
        var candidate = $"{owner}.{kind}.{suffix}";
        if (candidate.Length <= 255) return candidate;
        var ownerPrefix = owner.Length <= 180 ? owner : owner[..180];
        return $"{ownerPrefix}.{kind}.{Sha256(Encoding.UTF8.GetBytes(candidate))[..24]}";
    }

    private static string FlattenText(PpjTextContentModel text) => text.PlainText ??
        string.Join("\n", text.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(PpjRunText))));

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
        string elementId)
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
    }

    private static void ApplyCompoundShapeOpacity(PresentationShape target, double multiplier, string elementId)
    {
        if (multiplier >= 1) return;

        if (target.FillRgb.Length > 0 || target.FillScheme.Length > 0)
            target.FillOpacityThousandthPercent = MultiplyOpacity(
                target.HasFillOpacityThousandthPercent,
                target.FillOpacityThousandthPercent,
                multiplier);
        if (target.GradientFill is not null)
            MultiplyGradientOpacity(target.GradientFill, multiplier);
        if (target.ImageFill is not null)
            target.ImageFill.OpacityThousandthPercent = MultiplyOpacity(
                target.ImageFill.HasOpacityThousandthPercent,
                target.ImageFill.OpacityThousandthPercent,
                multiplier);
        if (target.LineStyle != "none" && (target.LineRgb.Length > 0 || target.LineScheme.Length > 0))
            target.LineOpacityThousandthPercent = MultiplyOpacity(
                target.HasLineOpacityThousandthPercent,
                target.LineOpacityThousandthPercent,
                multiplier);
        if (target.Shadow is not null)
            MultiplyShadowOpacity(target.Shadow, multiplier);
        if (target.TextBody is not null)
            MultiplyTextBodyOpacity(target.TextBody, multiplier, elementId);
    }

    private static void MultiplyTextBodyOpacity(PresentationTextBody body, double multiplier, string elementId)
    {
        foreach (var paragraph in body.Paragraphs)
        {
            if ((paragraph.BulletColorCase is PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb or
                 PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme))
                paragraph.BulletColorOpacityThousandthPercent = MultiplyOpacity(
                    paragraph.HasBulletColorOpacityThousandthPercent,
                    paragraph.BulletColorOpacityThousandthPercent,
                    multiplier);
            if (paragraph.DefaultRunProperties is not null)
                MultiplyTextStyleOpacity(paragraph.DefaultRunProperties, multiplier, elementId);
            foreach (var run in paragraph.Runs)
                MultiplyTextRunOpacity(run, multiplier, elementId);
        }
    }

    private static void MultiplyTextRunOpacity(PresentationTextRun run, double multiplier, string elementId)
    {
        if (run.HighlightCase != PresentationTextRun.HighlightOneofCase.None)
            throw Unsupported(elementId, "shape opacity cannot preserve text highlight because the bounded highlight profile has no alpha");
        if (PresentationRunText(run).Length > 0 && !run.HasColorRgb && !run.HasColorScheme && run.GradientFill is null)
            throw Unsupported(elementId, "shape opacity requires explicit text color or gradient paint");
        if (run.HasColorRgb || run.HasColorScheme)
            run.ColorOpacityThousandthPercent = MultiplyOpacity(
                run.HasColorOpacityThousandthPercent,
                run.ColorOpacityThousandthPercent,
                multiplier);
        if (run.GradientFill is not null)
            MultiplyGradientOpacity(run.GradientFill, multiplier);
        if (run.Shadow is not null)
            MultiplyShadowOpacity(run.Shadow, multiplier);
    }

    private static void MultiplyTextStyleOpacity(PresentationTextStyle style, double multiplier, string elementId)
    {
        if (style.HighlightCase != PresentationTextStyle.HighlightOneofCase.None)
            throw Unsupported(elementId, "shape opacity cannot preserve default text highlight because the bounded highlight profile has no alpha");
        if (style.ColorCase != PresentationTextStyle.ColorOneofCase.None)
            style.ColorOpacityThousandthPercent = MultiplyOpacity(
                style.HasColorOpacityThousandthPercent,
                style.ColorOpacityThousandthPercent,
                multiplier);
        if (style.GradientFill is not null)
            MultiplyGradientOpacity(style.GradientFill, multiplier);
        if (style.Shadow is not null)
            MultiplyShadowOpacity(style.Shadow, multiplier);
    }

    private static void MultiplyGradientOpacity(PresentationGradientFill fill, double multiplier)
    {
        foreach (var stop in fill.Stops)
            stop.OpacityThousandthPercent = MultiplyOpacity(
                stop.HasOpacityThousandthPercent,
                stop.OpacityThousandthPercent,
                multiplier);
    }

    private static void MultiplyShadowOpacity(PresentationShadow shadow, double multiplier) =>
        shadow.OpacityThousandthPercent = MultiplyOpacity(
            shadow.HasOpacityThousandthPercent,
            shadow.OpacityThousandthPercent,
            multiplier);

    private static uint MultiplyOpacity(bool hasLocalOpacity, uint localOpacity, double multiplier) =>
        Opacity((hasLocalOpacity ? localOpacity / 100_000d : 1d) * multiplier);

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
        Catalog catalog) => BuildTextBody(text, namedStyle, null, inlineStyle, catalog);

    private static PresentationTextBody BuildTextBody(
        JsonElement text,
        JsonElement? namedStyle,
        JsonElement? middleStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var body = new PresentationTextBody();
        ApplyBodyProperties(body.BodyProperties = new PresentationTextBodyProperties(), namedStyle, middleStyle, inlineStyle);
        if (text.ValueKind == JsonValueKind.String)
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(middleStyle, "paragraph"),
                Property(inlineStyle, "paragraph"),
                null,
                catalog);
            paragraph.Runs.Add(BuildRun(text.GetString()!, namedStyle, middleStyle, inlineStyle, null, null, catalog));
            body.Paragraphs.Add(paragraph);
            return body;
        }
        foreach (var paragraphJson in text.GetProperty("paragraphs").EnumerateArray())
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(middleStyle, "paragraph"),
                Property(inlineStyle, "paragraph"),
                Property(paragraphJson, "style"),
                catalog);
            foreach (var run in paragraphJson.GetProperty("runs").EnumerateArray())
            {
                paragraph.Runs.Add(run.TryGetProperty("formula", out var formula)
                    ? BuildFormulaRun(
                        formula.GetProperty("source").GetString()!,
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        catalog)
                    : BuildRun(
                        run.GetProperty("text").GetString()!,
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        Property(run, "hyperlink"),
                        catalog));
            }
            body.Paragraphs.Add(paragraph);
        }
        return body;
    }

    private static PresentationTextRun BuildFormulaRun(
        string source,
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        Catalog catalog)
    {
        var run = new PresentationTextRun { Formula = PpjLatexCompiler.Compile(source) };
        var inlineDefault = FirstProperty(inlineBox, null, "defaultText");
        var middleDefault = FirstProperty(middleBox, null, "defaultText");
        var namedDefault = FirstProperty(namedBox, null, "defaultText");
        if (FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "size") is { } size)
            run.FontSizePoints = size.GetDouble();
        if (FirstTextPaint(inlineRun, inlineDefault, middleDefault, namedDefault) is { Kind: "color" } paint)
        {
            var resolved = catalog.Color(paint.Value);
            run.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) run.ColorOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        return run;
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
        Catalog catalog) => BuildRun(text, namedBox, null, inlineBox, inlineRun, hyperlink, catalog);

    private static PresentationTextRun BuildRun(
        string text,
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        JsonElement? hyperlink,
        Catalog catalog)
    {
        var run = new PresentationTextRun { Text = text };
        var inlineDefault = FirstProperty(inlineBox, null, "defaultText");
        var middleDefault = FirstProperty(middleBox, null, "defaultText");
        var namedDefault = FirstProperty(namedBox, null, "defaultText");
        var bold = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "bold");
        var italic = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "italic");
        var size = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "size");
        var paint = FirstTextPaint(inlineRun, inlineDefault, middleDefault, namedDefault);
        var shadow = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "shadow");
        var highlight = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "highlight");
        var font = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "font");
        var family = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "fontFamily");
        var eastAsia = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "fontFamilyEastAsia");
        var underline = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "underline");
        var strike = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "strike");
        var kerning = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "kerning");
        var spacing = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "letterSpacing");
        var baseline = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "baseline");
        var capitalization = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "capitalization");
        var language = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "language");
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
        JsonElement? inline) => ApplyBodyProperties(target, named, null, inline);

    private static void ApplyBodyProperties(
        PresentationTextBodyProperties target,
        JsonElement? named,
        JsonElement? middle,
        JsonElement? inline)
    {
        var vertical = FirstProperty(inline, middle, named, "verticalAlignment");
        if (vertical is { } verticalValue)
            target.VerticalAnchor = verticalValue.GetString() == "middle" ? "center" : verticalValue.GetString()!;
        var autoFit = FirstProperty(inline, middle, named, "autoFit");
        if (autoFit is { } autoFitValue)
            target.AutoFitMode = autoFitValue.GetString() switch
            {
                "shrink-text" => "shrinkText",
                "resize-shape" => "resizeShape",
                _ => "none",
            };
        var wrap = FirstProperty(inline, middle, named, "wrap");
        if (wrap is { } wrapValue) target.Wrap = wrapValue.GetString()!;
        var margins = FirstProperty(inline, middle, named, "margins");
        if (margins is { } inset)
        {
            if (inset.TryGetProperty("left", out var left)) target.LeftInsetEmu = Emu(left.GetDouble());
            if (inset.TryGetProperty("top", out var top)) target.TopInsetEmu = Emu(top.GetDouble());
            if (inset.TryGetProperty("right", out var right)) target.RightInsetEmu = Emu(right.GetDouble());
            if (inset.TryGetProperty("bottom", out var bottom)) target.BottomInsetEmu = Emu(bottom.GetDouble());
        }
        var columns = FirstProperty(inline, middle, named, "columns");
        if (columns is { } columnCount) target.Columns = checked((uint)columnCount.GetInt32());
        var gap = FirstProperty(inline, middle, named, "columnGap");
        if (gap is { } columnGap) target.ColumnSpacingEmu = Emu(columnGap.GetDouble());
        var columnDirection = FirstProperty(inline, middle, named, "columnDirection");
        if (columnDirection is { } direction) target.RightToLeftColumns = direction.GetString() == "right-to-left";
        var verticalText = FirstProperty(inline, middle, named, "verticalText");
        if (verticalText is { } textMode) target.VerticalTextMode = textMode.GetString()!;
    }

    private static void ApplyParagraphStyle(
        PresentationTextParagraph target,
        JsonElement? named,
        JsonElement? inline,
        JsonElement? direct,
        Catalog catalog) => ApplyParagraphStyle(target, named, null, inline, direct, catalog);

    private static void ApplyParagraphStyle(
        PresentationTextParagraph target,
        JsonElement? named,
        JsonElement? middle,
        JsonElement? inline,
        JsonElement? direct,
        Catalog catalog)
    {
        if (FirstProperty(direct, inline, middle, named, "alignment") is { } alignment)
            target.Alignment = alignment.GetString()!;
        if (FirstProperty(direct, inline, middle, named, "level") is { } level)
            target.Level = checked((uint)level.GetInt32());
        if (FirstProperty(direct, inline, middle, named, "indent") is { } indent)
            target.MarginLeftEmu = Emu(indent.GetDouble());
        if (FirstProperty(direct, inline, middle, named, "hanging") is { } hanging)
            target.IndentEmu = -Emu(hanging.GetDouble());
        if (FirstProperty(direct, inline, middle, named, "spaceBefore") is { } before)
            target.SpaceBeforePoints = before.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "spaceBeforeMultiplier") is { } beforeMultiplier)
            target.SpaceBeforeMultiplier = beforeMultiplier.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "spaceAfter") is { } after)
            target.SpaceAfterPoints = after.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "spaceAfterMultiplier") is { } afterMultiplier)
            target.SpaceAfterMultiplier = afterMultiplier.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "lineSpacing") is { } spacing)
            target.LineSpacingPoints = spacing.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "lineSpacingMultiplier") is { } spacingMultiplier)
            target.LineSpacingMultiplier = spacingMultiplier.GetDouble();
        if (FirstProperty(direct, inline, middle, named, "defaultText") is { } defaultText && defaultText.EnumerateObject().Any())
            target.DefaultRunProperties = BuildTextStyle(defaultText, catalog);
        if (FirstProperty(direct, inline, middle, named, "bullet") is { } bullet)
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

    internal static PresentationAnimation BuildAnimation(PpjAnimationModel source, IReadOnlyList<PpjElementModel> elements)
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
                DurationMs = checked((uint)(page.Transition.DurationMs ?? 800)),
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
        slide.Transition = PpjTransitionLowering.BuildBase(page.Transition);
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
        if (FirstProperty(inline, named, "startAngle") is { } startAngle)
        {
            if (chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
                throw Unsupported(elementId, "startAngle applies only to pie and doughnut charts");
            chart.FirstSliceAngle = checked((uint)startAngle.GetInt32());
        }
        if (FirstProperty(inline, named, "holeSize") is { } holeSize)
        {
            if (chart.Type != SpreadsheetChartType.Doughnut)
                throw Unsupported(elementId, "holeSize applies only to doughnut charts");
            chart.DoughnutHoleSize = checked((uint)holeSize.GetInt32());
        }
        if (FirstProperty(inline, named, "bubbleScale") is { } bubbleScale)
        {
            if (chart.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(elementId, "bubbleScale applies only to bubble charts");
            chart.BubbleScale = checked((uint)bubbleScale.GetInt32());
        }
        if (FirstProperty(inline, named, "bubbleSizeMode") is { } bubbleSizeMode)
        {
            if (chart.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(elementId, "bubbleSizeMode applies only to bubble charts");
            chart.BubbleSizeMode = bubbleSizeMode.GetString()!;
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
            if (dataLabels.TryGetProperty("numberFormat", out var numberFormat)) chart.DataLabels.NumberFormatCode = numberFormat.GetString()!;
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

    private static void ApplyAccessibility(PresentationDiagram target, PpjAccessibilityModel? source) =>
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
        text.PlainText ?? string.Join('\n', text.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(PpjRunText))));

    private static string Flatten(PresentationTextBody body) =>
        string.Join('\n', body.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(PresentationRunText))));

    private static string PpjRunText(PpjRunModel run) => run.Text ??
        (run.Formula is null ? string.Empty : PpjLatexCompiler.Compile(run.Formula.Source).PlainText);

    private static string PresentationRunText(PresentationTextRun run) => run.ContentCase switch
    {
        PresentationTextRun.ContentOneofCase.Text => run.Text,
        PresentationTextRun.ContentOneofCase.Formula => run.Formula.PlainText,
        PresentationTextRun.ContentOneofCase.LineBreak => "\n",
        PresentationTextRun.ContentOneofCase.Field => run.Field.Text,
        _ => string.Empty,
    };

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

    private static JsonElement? FirstProperty(
        JsonElement? first,
        JsonElement? second,
        JsonElement? third,
        JsonElement? fourth,
        string name) =>
        Property(first, name) ?? Property(second, name) ?? Property(third, name) ?? Property(fourth, name);

    private static JsonElement? Property(JsonElement? value, string name) =>
        value is { ValueKind: JsonValueKind.Object } element && element.TryGetProperty(name, out var property) ? property : null;

    private static JsonElement? MergeJsonObjects(params JsonElement?[] sources)
    {
        Dictionary<string, JsonElement>? merged = null;
        foreach (var source in sources)
        {
            if (source is not { ValueKind: JsonValueKind.Object } value) continue;
            merged ??= new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
                merged[property.Name] = property.Value.Clone();
        }
        if (merged is null) return null;

        // Keep this path NativeAOT-safe: serializing JsonObject would require
        // reflection metadata that the standalone codec deliberately omits.
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            foreach (var (name, value) in merged)
            {
                writer.WritePropertyName(name);
                value.WriteTo(writer);
            }
            writer.WriteEndObject();
        }
        using var document = JsonDocument.Parse(buffer.WrittenMemory);
        return document.RootElement.Clone();
    }

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
        private readonly Dictionary<string, PpjSmartArtDefinition> _smartArtDefinitions;

        internal PresentationThemeArtifact Theme { get; }

        internal Catalog(JsonElement root, IReadOnlyList<Asset>? compiledAssets = null)
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
            _smartArtDefinitions = new Dictionary<string, PpjSmartArtDefinition>(StringComparer.Ordinal);
            if (compiledAssets is not null && root.TryGetProperty("assets", out assets))
            {
                var compiledById = compiledAssets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
                foreach (var declaration in assets.EnumerateArray().Where(asset =>
                             asset.GetProperty("mimeType").GetString() ==
                             PptxAssetCatalog.SmartArtDefinitionContentType))
                {
                    var programId = declaration.GetProperty("id").GetString()!;
                    var nativeId = _nativeAssetIds[programId];
                    if (!compiledById.TryGetValue(nativeId, out var asset))
                        throw new CodecException(
                            "ppj.asset.missing",
                            $"PPJ SmartArt definition asset {programId} was not supplied.",
                            "$.assets");
                    _smartArtDefinitions[programId] = PpjSmartArtDefinitionCodec.Parse(
                        asset,
                        $"$.assets[{programId}]");
                }
            }
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
        internal PpjSmartArtDefinition SmartArtDefinition(string definitionAssetId) =>
            _smartArtDefinitions.TryGetValue(definitionAssetId, out var definition)
                ? definition
                : throw new CodecException(
                    "ppj.smartArt.definitionUnavailable",
                    $"PPJ SmartArt definition asset {definitionAssetId} is not available to the authored compiler.");

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

    private static string NativeAssetId(string mimeType, string sha256) =>
        PptxAssetCatalog.NativeAssetIdFor(mimeType, sha256);
}
