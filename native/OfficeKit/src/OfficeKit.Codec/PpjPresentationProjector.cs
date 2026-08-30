using System.Buffers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjProjectionResult(
    PresentationProgramResult Program,
    IReadOnlyList<Diagnostic> Diagnostics,
    ArtifactEnvelope? SourceArtifact,
    IReadOnlyDictionary<string, PpjNativeLeafBinding> NativeLeafBindings);

/// <summary>
/// Projects a validated PPTX package into the bounded public PPJ language.
/// Native package locators and raw OOXML deliberately stay behind opaque,
/// source-bound handles; the program exposes only semantic state and
/// hash-bound capabilities that the existing source-preserving writer can
/// independently prove again during compilation.
/// </summary>
internal static partial class PpjPresentationProjector
{
    private const double EmuPerPoint = 12_700d;

    internal static PpjProjectionResult Project(
        byte[] sourceBytes,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits)
    {
        if (PpjEmbeddedProgramCodec.TryRecover(sourceBytes, request, limits) is { } recovered)
            return new(recovered.Program, recovered.Diagnostics, null, new Dictionary<string, PpjNativeLeafBinding>(StringComparer.Ordinal));

        var imported = PptxCodec.Import(sourceBytes, limits);
        var envelope = imported.Artifact;
        var presentation = envelope.Presentation ??
            throw new CodecException("ppj.projection.presentation", "The imported package did not produce a Presentation artifact.", "$");
        var sourceSha256 = Sha256(sourceBytes);
        var revision = $"pptx-{sourceSha256[..16]}";
        var sourceUri = string.IsNullOrWhiteSpace(request.SourceUri)
            ? $"deck.assets/source/{sourceSha256}.pptx"
            : request.SourceUri;
        var assetRoot = string.IsNullOrWhiteSpace(request.AssetRootUri)
            ? "deck.assets/media"
            : request.AssetRootUri.TrimEnd('/');

        var context = new ProjectionContext(sourceSha256, revision, assetRoot, envelope.Assets);
        RegisterIds(presentation, context);

        var pages = new JsonArray();
        foreach (var slide in presentation.Slides)
            pages.Add(ProjectPage(slide, presentation, context));

        var assets = context.ProgramAssets;
        var sections = ProjectSections(presentation, context);
        var customShows = ProjectCustomShows(presentation, context);
        var comments = ProjectComments(presentation, context);
        var nodeMap = context.BuildNodeMap();
        var nodeMapBytes = CanonicalBytes(nodeMap);
        var projectionPayload = new JsonObject
        {
            ["canvas"] = FrameDimensions(presentation),
            ["assets"] = assets.DeepClone(),
            ["pages"] = pages.DeepClone(),
            ["sections"] = sections.DeepClone(),
            ["customShows"] = customShows.DeepClone(),
            ["comments"] = comments.DeepClone(),
        };
        var projectionSha256 = Sha256(CanonicalBytes(projectionPayload));

        var root = new JsonObject
        {
            ["schema"] = "office-kit/ppj/v1",
            ["meta"] = new JsonObject
            {
                ["id"] = StableDocumentId(presentation.Id, sourceSha256),
                ["title"] = string.IsNullOrWhiteSpace(presentation.Name) ? "Imported presentation" : presentation.Name,
                ["language"] = "und",
                ["version"] = 1,
                ["description"] = "Source-derived PPJ projection. Unmodeled native content remains in the hash-bound PPTX source package.",
            },
            ["intent"] = ImportedIntent(),
            ["design"] = ImportedDesign(presentation),
            ["assets"] = assets,
            ["source"] = new JsonObject
            {
                ["kind"] = "pptx",
                ["uri"] = sourceUri,
                ["sha256"] = sourceSha256,
                ["revision"] = revision,
                ["projection"] = new JsonObject
                {
                    ["version"] = 1,
                    ["sha256"] = projectionSha256,
                    ["nodeMapSha256"] = Sha256(nodeMapBytes),
                    ["visibleObjectCount"] = context.VisibleObjectCount,
                },
            },
            ["pages"] = pages,
        };
        if (sections.Count > 0) root["sections"] = sections;
        if (customShows.Count > 0) root["customShows"] = customShows;
        if (comments.Count > 0) root["comments"] = comments;

        var candidateBytes = CanonicalBytes(root);
        var validation = PpjProgramValidator.Validate(candidateBytes);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        var result = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? ByteString.CopyFrom(nodeMapBytes) : ByteString.Empty,
            SourceSha256 = sourceSha256,
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        result.Assets.Add(context.ResultAssets.Select(asset => asset.Clone()));
        return new(result, imported.Diagnostics, envelope, context.NativeLeafBindings);
    }

    private static JsonObject ImportedIntent() => new()
    {
        ["brief"] = new JsonObject
        {
            ["primaryJob"] = "handoff",
            ["expectedOutcome"] = "Continue editing this presentation while preserving source-owned native content.",
            ["evidenceBoundary"] = "The source presentation does not declare its audience, factual authority, or intended outcome.",
        },
        ["audience"] = new JsonObject
        {
            ["description"] = "Imported presentation audience was not declared.",
        },
        ["narrative"] = new JsonObject
        {
            ["thesis"] = "Preserve and continue the imported presentation without inventing its original intent.",
        },
        ["editorial"] = new JsonObject
        {
            ["tone"] = new JsonArray("source-derived"),
            ["avoid"] = new JsonArray("Inventing missing facts or silently rewriting source-owned content"),
        },
        ["delivery"] = new JsonObject
        {
            ["mode"] = "hybrid",
            ["mediumFit"] = "acceptable",
            ["mediumFitNote"] = "Delivery intent was not declared in the imported file.",
        },
    };

    private static JsonObject ImportedDesign(PresentationArtifact presentation) => new()
    {
        ["canvas"] = FrameDimensions(presentation),
        ["theme"] = new JsonObject
        {
            ["name"] = "Source-owned presentation theme",
            // Imported run and bullet styles may retain a direct theme token
            // even though the source theme graph is opaque to the bounded
            // writer. Keep every standard token addressable with a neutral
            // fallback so the PPJ projection stays valid without claiming
            // that these fallback RGB values replace the source theme.
            ["colors"] = ImportedThemeColors(),
        },
        ["fonts"] = new JsonArray(new JsonObject
        {
            ["id"] = "source-font",
            ["family"] = "Arial",
            ["language"] = "und",
        }),
        ["styles"] = new JsonObject(),
        ["grammar"] = new JsonObject
        {
            ["name"] = "Source-derived projection",
            ["rationale"] = "The original PPTX remains the authority for native visual styling that PPJ does not model.",
            ["visualThesis"] = "Preserve the imported design and expose only bounded semantic edits.",
            ["surfaceHierarchy"] = new JsonArray("Keep source-owned surfaces unchanged unless a capability explicitly permits an edit."),
            ["typographyRhythm"] = new JsonArray("Retain imported typography through the source package."),
            ["geometryRules"] = new JsonArray("Retain imported geometry and z-order unless a nativeRef capability permits a change."),
            ["densityRhythm"] = new JsonArray("Retain the source page density."),
            ["carrierRules"] = new JsonArray("Use projected typed objects where safe and opaque native objects everywhere else."),
            ["forbiddenPatterns"] = new JsonArray("Rebuilding opaque content", "Guessing unsupported native semantics"),
        },
        ["motionPolicy"] = "explicit",
    };

    private static JsonObject FrameDimensions(PresentationArtifact presentation) => new()
    {
        ["width"] = Points(presentation.SlideWidthEmu),
        ["height"] = Points(presentation.SlideHeightEmu),
        ["unit"] = "pt",
    };

    private static void RegisterIds(PresentationArtifact presentation, ProjectionContext context)
    {
        foreach (var master in presentation.Masters) context.RegisterMaster(master.Id);
        foreach (var layout in presentation.Layouts) context.RegisterLayout(layout.Id);
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.RegisterPage(slide.Id);
            RegisterElementIds(slide.Elements, pageId, context);
        }
    }

    private static void RegisterElementIds(IEnumerable<PresentationElement> elements, string pageId, ProjectionContext context)
    {
        foreach (var element in elements)
        {
            context.RegisterElement(pageId, element.Id);
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                RegisterElementIds(element.Group.Children, pageId, context);
        }
    }

    private static JsonObject ProjectPage(
        PresentationSlide slide,
        PresentationArtifact presentation,
        ProjectionContext context)
    {
        var pageId = context.PageId(slide.Id);
        var pageHash = HashOrFallback(slide.Source?.SlideXmlSha256, slide.ToByteArray());
        var pageCapabilities = new List<CapabilitySpec>();
        if (slide.Source?.DeletionCapability?.Supported == true)
            pageCapabilities.Add(new("delete", ["element"]));
        if (slide.Source?.BackgroundEditable == true)
            pageCapabilities.Add(new("setBackground", ["background"]));
        if (slide.Source?.TransitionEditable == true || slide.Source?.TransitionAddable == true)
            pageCapabilities.Add(new("setTransition", ["transition"]));
        // A native slide clone needs fresh page/element identities and a
        // complete source-owned subtree mapping. Do not issue the underlying
        // codec capability until PPJ can represent that bounded clone request.

        var elements = new JsonArray();
        for (var elementIndex = 0; elementIndex < slide.Elements.Count; elementIndex++)
        {
            var element = slide.Elements[elementIndex];
            elements.Add(ProjectElement(
                element,
                slide,
                pageId,
                context,
                [element.Source?.ShapeTreeIndex ?? checked((uint)elementIndex)]));
        }

        var page = new JsonObject
        {
            ["id"] = pageId,
            ["role"] = "source continuation",
            ["elements"] = elements,
            ["nativeRef"] = NativeRef(context, $"page:{pageId}", pageHash, pageCapabilities),
        };
        if (!string.IsNullOrWhiteSpace(slide.Name)) page["name"] = slide.Name;
        if (!string.IsNullOrWhiteSpace(slide.LayoutId) && context.TryLayoutId(slide.LayoutId, out var layoutId))
            page["layout"] = layoutId;
        if (slide.HasHidden) page["hidden"] = slide.Hidden;
        if (!string.IsNullOrEmpty(slide.SpeakerNotes?.Text)) page["notes"] = slide.SpeakerNotes.Text;
        if (ProjectBackground(slide.Background, context) is { } background) page["background"] = background;

        var animations = ProjectAnimations(slide, context);
        if (animations.Count > 0) page["animations"] = animations;
        if (ProjectTransition(slide, presentation, context) is { } transition) page["transition"] = transition;
        return page;
    }

    private static JsonObject ProjectElement(
        PresentationElement element,
        PresentationSlide slide,
        string pageId,
        ProjectionContext context,
        IReadOnlyList<uint> shapeTreePath)
    {
        var id = context.ElementId(pageId, element.Id);
        var hash = HashOrFallback(element.Source?.ElementSha256, element.ToByteArray());
        var capabilities = Capabilities(element);
        var leaves = PpjNativeLeafProjection.Describe(
            context.SourceSha256,
            pageId,
            id,
            element,
            shapeTreePath,
            context.RecordNativeLeaf);
        var nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, capabilities, leaves);
        JsonObject projected = element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => ProjectShape(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Image => ProjectImage(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Table => ProjectTable(element, id, nativeRef),
            PresentationElement.ContentOneofCase.Connector => ProjectConnector(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Chart => ProjectChart(element, id, nativeRef),
            PresentationElement.ContentOneofCase.Group => ProjectGroup(element, id, nativeRef, slide, pageId, context, shapeTreePath),
            PresentationElement.ContentOneofCase.Opaque => ProjectOpaque(element, id, nativeRef),
            _ => ProjectOpaque(element, id, nativeRef, "unknown"),
        };
        if (element.HasHidden && element.Hidden) projected["hidden"] = true;
        if (element.HasLocked && element.Locked) projected["locked"] = true;
        if (projected["type"]!.GetValue<string>() == "opaque")
        {
            // ProjectShape/ProjectImage may conservatively fall back to opaque
            // after the first nativeRef already owns this JsonArray. Clone the
            // issued leaf descriptors; a JsonNode cannot have two parents.
            nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, OpaqueCapabilities(element), leaves.DeepClone().AsArray());
            projected["nativeRef"] = nativeRef;
        }
        context.RecordNode(pageId, id, projected["type"]!.GetValue<string>(), nativeRef);
        return projected;
    }

    private static JsonObject ProjectShape(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var shape = element.Shape;
        var frame = ShapeFrame(shape);
        var text = TextContent(shape.TextBody, shape.Text);
        var hasText = !string.IsNullOrEmpty(shape.Text) || shape.TextBody?.Paragraphs.Count > 0;
        var isPlaceholder = shape.Placeholder is not null;
        var isTextBox = shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry);
        // The importer deliberately leaves unsupported custom path graphs
        // empty.  The source-bound projection can still expose the owning
        // shape (frame, native leaves, and safe paint fields) without
        // pretending that its path topology is editable.
        var sourceCustomGeometry = shape.Geometry == "custom";
        var sourceImageFill = !string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
            context.TryMaterializeAsset(shape.ImageFillAssetId, out var sourceImageAssetId);
        if (shape.ImageFill is not null && !context.TryMaterializeAsset(shape.ImageFill.AssetId, out _))
            return ProjectOpaque(element, id, nativeRef, "shape", "Preserved source shape whose image fill cannot be materialized safely.");
        if (!isPlaceholder && !isTextBox && !PptxPresetGeometryAdjustmentCodec.HasProfile(shape.Geometry) &&
            !CanProjectCustomGeometry(shape) && !sourceImageFill && !sourceCustomGeometry)
            return ProjectOpaque(element, id, nativeRef, "shape", $"Preserved source shape with unsupported geometry '{shape.Geometry}'.");
        // A legacy line stored as p:sp has no connector endpoints in the
        // public PPJ model. Keep it source-owned (with the generic bounded
        // frame/reorder capabilities) instead of emitting an invalid shape
        // preset that cannot round-trip as a connector.
        if (shape.Geometry == "line")
            return ProjectOpaque(element, id, nativeRef, "shape", "Preserved source line without connector endpoint semantics.");

        var common = ElementBase(id, element.Name, frame, Accessibility(shape.Accessibility), nativeRef);
        if (shape.Placeholder is not null)
        {
            common["type"] = "placeholder";
            common["placeholderType"] = PlaceholderType(shape.Placeholder.Type);
            common["index"] = shape.Placeholder.Index;
            if (hasText) common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } placeholderStyle) common["style"] = placeholderStyle;
            return common;
        }
        if (shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry))
        {
            common["type"] = "text";
            common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } textStyle) common["style"] = textStyle;
            ApplyTextContainerStyle(common, shape, context);
            return common;
        }
        common["type"] = "shape";
        if (shape.Geometry == "custom")
            common["geometry"] = CanProjectCustomGeometry(shape)
                ? ProjectCustomGeometry(shape)
                : new JsonObject
                {
                    // The path graph is intentionally not guessed when a
                    // source-bound image-filled custom shape falls outside
                    // the bounded custom-geometry profile. Keep the shape's
                    // native geometry behind an explicit source-bound marker
                    // while exposing its frame, image asset and capabilities.
                    ["kind"] = "source-custom",
                    ["sourceBound"] = true,
                };
        else
        {
            var geometry = new JsonObject { ["kind"] = "preset", ["preset"] = shape.Geometry };
            if (shape.PresetAdjustments.Count > 0)
                geometry["adjustments"] = new JsonArray(shape.PresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            common["geometry"] = geometry;
        }
        if (hasText) common["text"] = text;
        if (TextBoxStyle(shape.TextBody) is { Count: > 0 } shapeTextStyle) common["textStyle"] = shapeTextStyle;
        var style = ShapeStyle(shape, context);
        if (style.Count > 0) common["style"] = style;
        return common;
    }

    private static bool CanProjectCustomGeometry(PresentationShape shape)
    {
        if (shape.Geometry != "custom" || shape.CustomPaths.Count == 0 ||
            shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
            shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 ||
            shape.TextRectangle is not null)
            return false;
        var width = shape.CustomPaths[0].Width;
        var height = shape.CustomPaths[0].Height;
        if (width <= 0 || height <= 0 || shape.CustomPaths.Any(path => path.Width != width || path.Height != height))
            return false;
        return shape.CustomPaths.SelectMany(path => path.Commands).All(command => command.CommandCase switch
        {
            PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => Literal(command.MoveTo),
            PresentationCustomGeometryCommand.CommandOneofCase.LineTo => Literal(command.LineTo),
            PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo =>
                Literal(command.QuadraticBezierTo.Control) && Literal(command.QuadraticBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo =>
                Literal(command.CubicBezierTo.Control1) && Literal(command.CubicBezierTo.Control2) && Literal(command.CubicBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => Literal(command.ArcTo),
            PresentationCustomGeometryCommand.CommandOneofCase.Close => true,
            _ => false,
        });
    }

    private static bool Literal(PresentationCustomGeometryPoint point) =>
        !point.HasXReference && !point.HasYReference;

    private static bool Literal(PresentationCustomGeometryArc arc) =>
        !arc.HasWidthRadiusReference && !arc.HasHeightRadiusReference &&
        !arc.HasStartAngleReference && !arc.HasSweepAngleReference;

    private static JsonObject ProjectCustomGeometry(PresentationShape shape)
    {
        var width = shape.CustomPaths[0].Width / 1_000d;
        var height = shape.CustomPaths[0].Height / 1_000d;
        var paths = new JsonArray();
        foreach (var source in shape.CustomPaths)
        {
            var commands = new JsonArray();
            foreach (var command in source.Commands) commands.Add(ProjectCustomCommand(command));
            var path = new JsonObject { ["commands"] = commands };
            if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.Normal) path["fill"] = true;
            else if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.None) path["fill"] = false;
            if (source.HasStroke) path["stroke"] = source.Stroke;
            paths.Add(path);
        }
        return new JsonObject
        {
            ["kind"] = "custom",
            ["viewBox"] = new JsonObject { ["x"] = 0, ["y"] = 0, ["width"] = width, ["height"] = height },
            ["paths"] = paths,
        };
    }

    private static JsonObject ProjectCustomCommand(PresentationCustomGeometryCommand command) => command.CommandCase switch
    {
        PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => ProjectCustomPoint("moveTo", command.MoveTo),
        PresentationCustomGeometryCommand.CommandOneofCase.LineTo => ProjectCustomPoint("lineTo", command.LineTo),
        PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo => new JsonObject
        {
            ["op"] = "quadraticTo",
            ["x1"] = CustomPathPoint(command.QuadraticBezierTo.Control.X),
            ["y1"] = CustomPathPoint(command.QuadraticBezierTo.Control.Y),
            ["x"] = CustomPathPoint(command.QuadraticBezierTo.End.X),
            ["y"] = CustomPathPoint(command.QuadraticBezierTo.End.Y),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new JsonObject
        {
            ["op"] = "cubicTo",
            ["x1"] = CustomPathPoint(command.CubicBezierTo.Control1.X),
            ["y1"] = CustomPathPoint(command.CubicBezierTo.Control1.Y),
            ["x2"] = CustomPathPoint(command.CubicBezierTo.Control2.X),
            ["y2"] = CustomPathPoint(command.CubicBezierTo.Control2.Y),
            ["x"] = CustomPathPoint(command.CubicBezierTo.End.X),
            ["y"] = CustomPathPoint(command.CubicBezierTo.End.Y),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => new JsonObject
        {
            ["op"] = "arcTo",
            ["radiusX"] = CustomPathPoint(command.ArcTo.WidthRadius),
            ["radiusY"] = CustomPathPoint(command.ArcTo.HeightRadius),
            ["startAngle"] = NormalizeCustomPathStartAngle(command.ArcTo.StartAngle),
            ["sweepAngle"] = CustomPathAngle(command.ArcTo.SweepAngle),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.Close => new JsonObject { ["op"] = "close" },
        _ => throw new InvalidOperationException("Unsupported PPJ custom path command passed the projection gate."),
    };

    private static JsonObject ProjectCustomPoint(string operation, PresentationCustomGeometryPoint point) => new()
    {
        ["op"] = operation,
        ["x"] = CustomPathPoint(point.X),
        ["y"] = CustomPathPoint(point.Y),
    };

    private static double CustomPathPoint(long value) => value / 1_000d;

    private static double CustomPathAngle(int value) => value / 60_000d;

    private static double NormalizeCustomPathStartAngle(int value)
    {
        var degrees = CustomPathAngle(value);
        return ((degrees % 360) + 360) % 360;
    }

    private static JsonObject ProjectImage(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var image = element.Image;
        if (!context.TryMaterializeAsset(image.AssetId, out var assetId))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose media payload cannot be materialized safely.");
        var customMask = image.CustomMaskPaths.Count > 0 ? ImageMaskShape(image) : null;
        if (customMask is not null && !CanProjectCustomGeometry(customMask))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose custom mask cannot be represented exactly in PPJ.");
        var output = ElementBase(id, element.Name, ImageFrame(image), ImageAccessibility(image), nativeRef);
        output["type"] = "image";
        output["asset"] = assetId;
        output["fit"] = image.Tiled ? "tile" : "stretch";
        if (image.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = Crop(image.Crop.LeftThousandthPercent),
                ["top"] = Crop(image.Crop.TopThousandthPercent),
                ["right"] = Crop(image.Crop.RightThousandthPercent),
                ["bottom"] = Crop(image.Crop.BottomThousandthPercent),
            };
        }
        if (image.HasOpacityThousandthPercent)
            output["opacity"] = Unit(image.OpacityThousandthPercent);
        if (customMask is not null)
            output["mask"] = ProjectCustomGeometry(customMask);
        else if (!string.IsNullOrEmpty(image.MaskPreset) && PptxPresetGeometryAdjustmentCodec.HasProfile(image.MaskPreset))
        {
            var mask = new JsonObject { ["kind"] = "preset", ["preset"] = image.MaskPreset };
            if (image.MaskPresetAdjustments.Count > 0)
                mask["adjustments"] = new JsonArray(image.MaskPresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            output["mask"] = mask;
        }
        if (image.Border is not null && !string.IsNullOrEmpty(image.Border.ColorRgb))
            output["border"] = Stroke(image.Border.ColorRgb, image.Border.WidthEmu, image.Border.Style, image.Border.Cap, image.Border.Join,
                image.Border.HasOpacityThousandthPercent ? Unit(image.Border.OpacityThousandthPercent) : null);
        if (image.Shadow is not null && (!string.IsNullOrEmpty(image.Shadow.ColorRgb) || !string.IsNullOrEmpty(image.Shadow.ColorScheme)))
            output["shadow"] = Shadow(image.Shadow);
        return output;
    }

    private static PresentationShape ImageMaskShape(PresentationImage image)
    {
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = image.WidthEmu,
            HeightEmu = image.HeightEmu,
        };
        shape.CustomPaths.Add(image.CustomMaskPaths);
        return shape;
    }

    private static JsonObject ProjectChart(PresentationElement element, string id, JsonObject nativeRef)
    {
        var chart = element.Chart;
        var type = ChartType(chart.Type);
        var series = chart.Type == SpreadsheetChartType.Combo
            ? chart.ComboSeries.Select(item => item.Series).ToArray()
            : chart.Series.ToArray();
        var numericX = chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;
        var invalidSeries = numericX
            ? chart.Categories.Count != 0 || series.Any(item =>
                item.Values.Count != item.XValues.Count ||
                chart.Type == SpreadsheetChartType.Bubble && item.Values.Count != item.BubbleSizes.Count ||
                chart.Type == SpreadsheetChartType.Scatter && item.BubbleSizes.Count != 0)
            : series.Any(item => item.Values.Count != chart.Categories.Count || item.XValues.Count != 0 || item.BubbleSizes.Count != 0);
        if (type is null || series.Length == 0 || invalidSeries)
            return ProjectOpaque(element, id, nativeRef, "chart", "Preserved source chart outside the bounded PPJ data profile.");

        var output = ElementBase(id, element.Name, ChartFrame(chart), Accessibility(chart.Accessibility), nativeRef);
        output["type"] = "chart";
        output["chartType"] = chart.Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar" ? "bar" : type;
        if (!string.IsNullOrEmpty(chart.Title))
            output["title"] = chart.TitleBody is null
                ? StringNode(chart.Title)
                : TextContent(chart.TitleBody, chart.Title);
        var categories = new JsonArray();
        foreach (var value in chart.Categories) categories.Add(StringNode(value));
        var seriesJson = new JsonArray();
        for (var index = 0; index < series.Length; index++)
        {
            var item = series[index];
            var values = new JsonArray();
            var missingValueIndexes = item.MissingValueIndexes.ToHashSet();
            for (var valueIndex = 0; valueIndex < item.Values.Count; valueIndex++)
                values.Add(missingValueIndexes.Contains((uint)valueIndex) ? null : NumberNode(item.Values[valueIndex]));
            var entry = new JsonObject
            {
                ["id"] = $"series-{index + 1}",
                ["name"] = item.Name ?? string.Empty,
                ["values"] = values,
            };
            if (item.XValues.Count > 0)
            {
                var xValues = new JsonArray();
                foreach (var value in item.XValues) xValues.Add(NumberNode(value));
                entry["xValues"] = xValues;
            }
            if (item.BubbleSizes.Count > 0)
            {
                var bubbleSizes = new JsonArray();
                foreach (var value in item.BubbleSizes) bubbleSizes.Add(NumberNode(value));
                entry["bubbleSizes"] = bubbleSizes;
            }
            if (chart.Type == SpreadsheetChartType.Combo)
            {
                entry["chartType"] = chart.ComboSeries[index].Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar"
                    ? "bar"
                    : ChartType(chart.ComboSeries[index].Type) ?? "line";
                entry["axis"] = chart.ComboSeries[index].AxisGroup == PresentationChartAxisGroup.Secondary ? "secondary" : "primary";
            }
            ProjectChartSeriesStyle(entry, item);
            seriesJson.Add(entry);
        }
        output["data"] = new JsonObject { ["categories"] = categories, ["series"] = seriesJson };
        if (chart.XAxis is not null) output["xAxis"] = ProjectChartAxis(chart.XAxis);
        if (chart.YAxis is not null) output["yAxis"] = ProjectChartAxis(chart.YAxis);
        if (chart.SecondaryXAxis is not null) output["secondaryXAxis"] = ProjectChartAxis(chart.SecondaryXAxis);
        if (chart.SecondaryYAxis is not null) output["secondaryYAxis"] = ProjectChartAxis(chart.SecondaryYAxis);
        var style = new JsonObject
        {
            ["legend"] = chart.HasLegend
                ? chart.LegendPosition.Length == 0 ? "right" : chart.LegendPosition
                : "none",
        };
        if (chart.Grouping.Length > 0) style["stacking"] = chart.Grouping;
        if (chart.HasGapWidth) style["gapWidth"] = chart.GapWidth;
        if (chart.XAxis is null && chart.HasShowCategoryAxis) style["showCategoryAxis"] = chart.ShowCategoryAxis;
        if (chart.YAxis is null && chart.HasShowValueAxis) style["showValueAxis"] = chart.ShowValueAxis;
        if (chart.HasShowGridlines) style["showGridlines"] = chart.ShowGridlines;
        if (chart.ChartAreaFill is not null) style["chartAreaFill"] = ProjectChartSurfaceFill(chart.ChartAreaFill);
        if (chart.PlotAreaFill is not null) style["plotAreaFill"] = ProjectChartSurfaceFill(chart.PlotAreaFill);
        if (chart.TitleTextStyle is not null)
            style["titleTextStyle"] = ProjectChartTextStyle(chart.TitleTextStyle);
        if (chart.LegendTextStyle is not null)
            style["legendTextStyle"] = ProjectChartTextStyle(chart.LegendTextStyle);
        if (chart.LineOptions?.HasSmooth == true) style["smooth"] = chart.LineOptions.Smooth;
        if (chart.LineOptions?.VaryColors == true) style["varyColors"] = true;
        if (chart.DataLabels is not null)
        {
            var labels = new JsonObject
            {
                ["showValue"] = chart.DataLabels.ShowValue,
                ["showCategory"] = chart.DataLabels.ShowCategoryName,
            };
            if (chart.DataLabels.HasShowSeriesName) labels["showSeries"] = chart.DataLabels.ShowSeriesName;
            if (chart.DataLabels.HasShowPercent) labels["showPercent"] = chart.DataLabels.ShowPercent;
            if (chart.DataLabels.HasPosition && DataLabelPosition(chart.DataLabels.Position) is { } position)
                labels["position"] = position;
            if (chart.DataLabels.TextStyle is not null)
                labels["textStyle"] = ProjectChartTextStyle(chart.DataLabels.TextStyle);
            style["dataLabels"] = labels;
        }
        output["style"] = style;
        return output;
    }

    private static void ProjectChartSeriesStyle(JsonObject output, SpreadsheetChartSeriesArtifact series)
    {
        if (series.SeriesFill is not null)
            output["fill"] = ProjectChartSurfaceFill(series.SeriesFill);
        else if (series.Fill is not null && !string.IsNullOrEmpty(series.Fill.Rgb))
            output["fill"] = new JsonObject { ["type"] = "solid", ["color"] = Color(series.Fill.Rgb) };
        if (series.Line is not null && !string.IsNullOrEmpty(series.Line.Color?.Rgb))
            output["stroke"] = ProjectChartLine(series.Line);
        if (series.Marker is not null && series.Marker.Symbol != SpreadsheetChartMarkerSymbol.Unspecified)
        {
            var symbol = Marker(series.Marker.Symbol);
            if (symbol is not null)
            {
                if (!series.Marker.HasSize && series.Marker.Fill is null && series.Marker.Line is null)
                    output["marker"] = symbol;
                else
                {
                    var marker = new JsonObject { ["symbol"] = symbol };
                    if (series.Marker.HasSize) marker["size"] = series.Marker.Size;
                    if (series.Marker.Fill is not null && !string.IsNullOrEmpty(series.Marker.Fill.Rgb))
                    {
                        var color = Color(series.Marker.Fill.Rgb);
                        if (series.Marker.HasFillOpacityThousandthPercent)
                        {
                            var alpha = Math.Clamp((int)Math.Round(Unit(series.Marker.FillOpacityThousandthPercent) * 255), 0, 255);
                            color += $"{alpha:X2}";
                        }
                        marker["fill"] = color;
                    }
                    if (series.Marker.Line is not null && !string.IsNullOrEmpty(series.Marker.Line.Color?.Rgb))
                        marker["stroke"] = ProjectChartLine(series.Marker.Line);
                    output["marker"] = marker;
                }
            }
        }
        if (series.Trendlines.Count > 0)
        {
            var trendlines = new JsonArray();
            foreach (var item in series.Trendlines)
            {
                if (TrendlineType(item.Type) is not { } type) continue;
                var trendline = new JsonObject { ["type"] = type };
                if (!string.IsNullOrEmpty(item.Name)) trendline["name"] = item.Name;
                if (item.HasPolynomialOrder) trendline["order"] = item.PolynomialOrder;
                if (item.HasPeriod) trendline["period"] = item.Period;
                if (item.HasForward) trendline["forward"] = item.Forward;
                if (item.HasBackward) trendline["backward"] = item.Backward;
                if (item.HasIntercept) trendline["intercept"] = item.Intercept;
                if (item.DisplayEquation) trendline["displayEquation"] = true;
                if (item.DisplayRSquared) trendline["displayRSquared"] = true;
                if (item.Line is not null && !string.IsNullOrEmpty(item.Line.Color?.Rgb))
                    trendline["stroke"] = ProjectChartLine(item.Line);
                trendlines.Add(trendline);
            }
            if (trendlines.Count > 0) output["trendlines"] = trendlines;
        }
        if (series.ErrorBars is { } errorBars &&
            ErrorBarDirection(errorBars.Direction) is { } direction &&
            ErrorBarType(errorBars.Type) is { } barType &&
            ErrorBarValueType(errorBars.ValueType) is { } valueType)
        {
            var projected = new JsonObject
            {
                ["direction"] = direction,
                ["type"] = barType,
                ["valueType"] = valueType,
            };
            if (errorBars.HasValue) projected["value"] = errorBars.Value;
            if (errorBars.NoEndCap) projected["noEndCap"] = true;
            if (errorBars.Line is not null && !string.IsNullOrEmpty(errorBars.Line.Color?.Rgb))
                projected["stroke"] = ProjectChartLine(errorBars.Line);
            output["errorBars"] = projected;
        }
    }

    private static JsonObject ProjectChartAxis(SpreadsheetChartAxisArtifact axis)
    {
        var output = new JsonObject();
        if (!string.IsNullOrEmpty(axis.Title)) output["title"] = axis.Title;
        if (!string.IsNullOrEmpty(axis.NumberFormatCode)) output["numberFormat"] = axis.NumberFormatCode;
        if (axis.HasTickLabelInterval) output["tickLabelInterval"] = axis.TickLabelInterval;
        if (axis.HasMinimum) output["min"] = axis.Minimum;
        if (axis.HasMaximum) output["max"] = axis.Maximum;
        if (axis.HasMajorUnit) output["majorUnit"] = axis.MajorUnit;
        if (axis.HasVisible) output["visible"] = axis.Visible;
        if (axis.TextStyle is not null)
            output["textStyle"] = ProjectChartTextStyle(axis.TextStyle);
        if (axis.TitleTextStyle is not null)
            output["titleTextStyle"] = ProjectChartTextStyle(axis.TitleTextStyle);
        return output;
    }

    private static JsonObject ProjectChartTextStyle(SpreadsheetChartTextStyleArtifact source)
    {
        var output = new JsonObject();
        if (source.HasFontSizePoints) output["fontSize"] = source.FontSizePoints;
        if (source.FontFamily.Length > 0) output["fontFamily"] = source.FontFamily;
        if (source.FontFamilyEastAsia.Length > 0) output["fontFamilyEastAsia"] = source.FontFamilyEastAsia;
        if (source.HasBold) output["bold"] = source.Bold;
        if (source.HasItalic) output["italic"] = source.Italic;
        if (source.ColorRgb.Length > 0)
            output["color"] = TextColor(source.ColorRgb, null,
                source.HasOpacityThousandthPercent, source.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject ProjectChartLine(SpreadsheetChartLineStyleArtifact line)
    {
        var output = new JsonObject
        {
            ["color"] = Color(line.Color.Rgb),
            ["width"] = line.HasWidthPoints ? line.WidthPoints : 0.75,
        };
        if (ChartDash(line.DashStyle) is { } dash) output["dash"] = dash;
        if (line.Cap is "flat" or "round" or "square") output["cap"] = line.Cap;
        if (line.Join is "miter" or "round" or "bevel") output["join"] = line.Join;
        if (line.HasOpacityThousandthPercent) output["opacity"] = Unit(line.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject ProjectChartSurfaceFill(SpreadsheetChartSurfaceFill fill)
    {
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            return new JsonObject { ["type"] = "none" };
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            return Gradient(fill.GradientFill);
        var output = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = Color(fill.SolidRgb),
        };
        if (fill.HasOpacityThousandthPercent) output["opacity"] = Unit(fill.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject ProjectTable(PresentationElement element, string id, JsonObject nativeRef)
    {
        var table = element.Table;
        if (table.ColumnWidthsEmu.Count == 0 || table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != table.ColumnWidthsEmu.Count))
            return ProjectOpaque(element, id, nativeRef, "table", "Preserved source table outside the bounded rectangular grid profile.");

        var output = ElementBase(id, element.Name, TableFrame(table), Accessibility(table.Accessibility), nativeRef);
        output["type"] = "table";
        var columns = new JsonArray();
        for (var index = 0; index < table.ColumnWidthsEmu.Count; index++)
            columns.Add(new JsonObject { ["id"] = $"column-{index + 1}", ["width"] = Points(table.ColumnWidthsEmu[index]) });
        output["columns"] = columns;
        output["rows"] = ProjectTableRows(table);
        var style = new JsonObject();
        if (table.HasFirstRow) style["headerRows"] = table.FirstRow ? 1 : 0;
        if (table.HasBandedRows) style["bandedRows"] = table.BandedRows;
        if (table.HasBandedColumns) style["bandedColumns"] = table.BandedColumns;
        if (table.HasFirstColumn) style["firstColumnEmphasis"] = table.FirstColumn;
        if (table.HasLastColumn) style["lastColumnEmphasis"] = table.LastColumn;
        if (style.Count > 0) output["style"] = style;
        return output;
    }

    private static JsonArray ProjectTableRows(PresentationTable table)
    {
        var mergeByOrigin = table.MergeRanges.ToDictionary(
            item => (Row: (int)item.StartRow, Column: (int)item.StartColumn));
        var covered = new HashSet<(int Row, int Column)>();
        foreach (var merge in table.MergeRanges)
            for (var row = (int)merge.StartRow; row <= merge.EndRow; row++)
                for (var column = (int)merge.StartColumn; column <= merge.EndColumn; column++)
                    if (row != merge.StartRow || column != merge.StartColumn) covered.Add((row, column));

        var rows = new JsonArray();
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var cells = new JsonArray();
            for (var columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++)
            {
                if (covered.Contains((rowIndex, columnIndex))) continue;
                var cell = new JsonObject
                {
                    ["id"] = $"cell-{rowIndex + 1}-{columnIndex + 1}",
                    ["text"] = table.Rows[rowIndex].Cells[columnIndex].Text ?? string.Empty,
                };
                if (mergeByOrigin.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    cell["rowSpan"] = checked((int)(merge.EndRow - merge.StartRow + 1));
                    cell["columnSpan"] = checked((int)(merge.EndColumn - merge.StartColumn + 1));
                }
                cells.Add(cell);
            }
            rows.Add(new JsonObject
            {
                ["id"] = $"row-{rowIndex + 1}",
                ["height"] = Points(table.Rows[rowIndex].HeightEmu),
                ["cells"] = cells,
            });
        }
        return rows;
    }

    private static JsonObject ProjectConnector(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context)
    {
        var connector = element.Connector;
        var output = ElementBase(id, element.Name, ConnectorFrame(connector), Accessibility(connector.Accessibility), nativeRef);
        output["type"] = "connector";
        output["connectorType"] = connector.ConnectorType is "elbow" or "curved" ? connector.ConnectorType : "straight";
        output["from"] = ConnectorEndpoint(connector.StartTargetId, connector.StartXEmu, connector.StartYEmu, pageId, context);
        output["to"] = ConnectorEndpoint(connector.EndTargetId, connector.EndXEmu, connector.EndYEmu, pageId, context);
        output["stroke"] = Stroke(
            connector.LineRgb,
            connector.LineWidthEmu,
            connector.LineStyle,
            connector.LineCap,
            connector.LineJoin,
            connector.HasLineOpacityThousandthPercent ? Unit(connector.LineOpacityThousandthPercent) : null,
            connector.LineScheme);
        if (Arrow(connector.StartArrow) is { } startArrow) output["startArrow"] = startArrow;
        if (Arrow(connector.EndArrow) is { } endArrow) output["endArrow"] = endArrow;
        return output;
    }

    private static JsonObject ProjectGroup(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        PresentationSlide slide,
        string pageId,
        ProjectionContext context,
        IReadOnlyList<uint> shapeTreePath)
    {
        var group = element.Group;
        if (group.Children.Count == 0)
            return ProjectOpaque(element, id, nativeRef, "group", "Preserved empty or unsupported native group.");
        var output = ElementBase(id, element.Name, GroupFrame(group), Accessibility(group.Accessibility), nativeRef);
        output["type"] = "group";
        var children = new JsonArray();
        for (var index = 0; index < group.Children.Count; index++)
        {
            var child = group.Children[index];
            children.Add(ProjectElement(
                child,
                slide,
                pageId,
                context,
                shapeTreePath.Concat([child.Source?.ShapeTreeIndex ?? checked((uint)index)]).ToArray()));
        }
        output["elements"] = children;
        return output;
    }

    private static JsonObject ProjectOpaque(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string? kind = null,
        string? summary = null)
    {
        var opaque = element.ContentCase == PresentationElement.ContentOneofCase.Opaque ? element.Opaque : null;
        var nativeKind = kind ?? opaque?.NativeKind;
        if (string.IsNullOrWhiteSpace(nativeKind)) nativeKind = element.ContentCase.ToString();
        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = "opaque";
        output["nativeKind"] = StringNode(nativeKind);
        output["summary"] = StringNode(summary ?? $"Preserved source-owned {nativeKind} object; only issued nativeRef capabilities are editable.");
        if (opaque is not null && PpjNativeTextProjection.TryRead(opaque.RawXml, out var nativeLeaves))
        {
            var visible = new JsonArray();
            foreach (var leaf in nativeLeaves) visible.Add(StringNode(leaf));
            output["visibleText"] = visible;
        }
        else if (!string.IsNullOrWhiteSpace(opaque?.Text))
        {
            var visible = new JsonArray();
            foreach (var line in opaque.Text.Split('\n').Where(line => line.Length > 0)) visible.Add(StringNode(line));
            if (visible.Count > 0) output["visibleText"] = visible;
        }
        return output;
    }

    private static JsonObject ElementBase(
        string id,
        string? name,
        JsonObject frame,
        JsonObject? accessibility,
        JsonObject nativeRef)
    {
        var output = new JsonObject
        {
            ["id"] = id,
            ["frame"] = frame,
            ["nativeRef"] = nativeRef,
        };
        if (!string.IsNullOrWhiteSpace(name)) output["name"] = StringNode(name);
        if (accessibility is not null) output["accessibility"] = accessibility;
        return output;
    }

    private static JsonObject ShapeStyle(PresentationShape shape, ProjectionContext context)
    {
        var style = new JsonObject();
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = "solid", ["color"] = Color(shape.FillRgb) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = Unit(shape.FillOpacityThousandthPercent);
            style["fill"] = fill;
        }
        else if (shape.GradientFill is not null)
        {
            style["fill"] = Gradient(shape.GradientFill);
        }
        else if (shape.ImageFill is not null && ProjectImagePaint(shape.ImageFill, context) is { } imageFill)
        {
            style["fill"] = imageFill;
        }
        else if (!string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
                 context.TryMaterializeAsset(shape.ImageFillAssetId, out var sourceImageAssetId))
        {
            style["fill"] = new JsonObject
            {
                ["type"] = "image",
                ["asset"] = sourceImageAssetId,
                ["fit"] = "stretch",
            };
        }
        if ((!string.IsNullOrEmpty(shape.LineRgb) || !string.IsNullOrEmpty(shape.LineScheme)) && shape.LineStyle != "none")
            style["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin,
                shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                shape.LineScheme);
        if (shape.Shadow is not null && (!string.IsNullOrEmpty(shape.Shadow.ColorRgb) || !string.IsNullOrEmpty(shape.Shadow.ColorScheme)))
            style["shadow"] = Shadow(shape.Shadow);
        return style;
    }

    private static JsonNode TextContent(PresentationTextBody? body, string? fallback)
    {
        if (body is null || body.Paragraphs.Count == 0 ||
            body.Paragraphs.Any(paragraph => paragraph.Runs.Count == 0 ||
                paragraph.Runs.Any(run => run.ContentCase != PresentationTextRun.ContentOneofCase.Text)))
            return StringNode(fallback ?? string.Empty);

        var paragraphs = new JsonArray();
        for (var paragraphIndex = 0; paragraphIndex < body.Paragraphs.Count; paragraphIndex++)
        {
            var source = body.Paragraphs[paragraphIndex];
            var paragraph = new JsonObject
            {
                ["id"] = $"paragraph-{paragraphIndex + 1}",
            };
            var paragraphStyle = new JsonObject();
            if (source.HasLevel) paragraphStyle["level"] = checked((int)source.Level);
            if (source.HasAlignment && ParagraphAlignment(source.Alignment) is { } alignment)
                paragraphStyle["alignment"] = alignment;
            if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                paragraphStyle["indent"] = Points(source.MarginLeftEmu);
            if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                paragraphStyle["hanging"] = -Points(source.IndentEmu);
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                paragraphStyle["lineSpacing"] = Math.Max(0.001, source.LineSpacingPoints);
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                paragraphStyle["lineSpacingMultiplier"] = Math.Max(0.00001, source.LineSpacingMultiplier);
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                paragraphStyle["spaceBefore"] = Math.Max(0, source.SpaceBeforePoints);
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                paragraphStyle["spaceBeforeMultiplier"] = Math.Max(0, source.SpaceBeforeMultiplier);
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                paragraphStyle["spaceAfter"] = Math.Max(0, source.SpaceAfterPoints);
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                paragraphStyle["spaceAfterMultiplier"] = Math.Max(0, source.SpaceAfterMultiplier);
            if (source.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                ProjectTextStyle(source.DefaultRunProperties) is { Count: > 0 } defaultText)
                paragraphStyle["defaultText"] = defaultText;
            if (ProjectBullet(source) is { } bullet) paragraphStyle["bullet"] = bullet;
            if (paragraphStyle.Count > 0) paragraph["style"] = paragraphStyle;

            var runs = new JsonArray();
            for (var runIndex = 0; runIndex < source.Runs.Count; runIndex++)
            {
                var sourceRun = source.Runs[runIndex];
                var run = new JsonObject
                {
                    ["id"] = $"run-{paragraphIndex + 1}-{runIndex + 1}",
                    ["text"] = sourceRun.Text,
                };
                var style = RunStyle(sourceRun);
                if (style.Count > 0) run["style"] = style;
                runs.Add(run);
            }
            paragraph["runs"] = runs;
            paragraphs.Add(paragraph);
        }
        return new JsonObject { ["paragraphs"] = paragraphs };
    }

    private static JsonObject RunStyle(PresentationTextRun run)
    {
        var style = new JsonObject();
        if (run.HasFontFamily) style["fontFamily"] = run.FontFamily;
        if (run.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = run.FontFamilyEastAsia;
        if (run.HasFontSizePoints && run.FontSizePoints > 0) style["size"] = run.FontSizePoints;
        if (run.HasBold) style["bold"] = run.Bold;
        if (run.HasItalic) style["italic"] = run.Italic;
        if (run.HasColorRgb && !string.IsNullOrEmpty(run.ColorRgb))
            style["color"] = TextColor(run.ColorRgb, null, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.HasColorScheme && !string.IsNullOrEmpty(run.ColorScheme))
            style["color"] = TextColor(null, run.ColorScheme, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.GradientFill is not null)
            style["gradient"] = TextGradient(run.GradientFill);
        if (run.Shadow is not null) style["shadow"] = Shadow(run.Shadow);
        if (run.HighlightCase == PresentationTextRun.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(run.HighlightRgb))
            style["highlight"] = Color(run.HighlightRgb);
        if (run.HasUnderline) style["underline"] = run.Underline switch { "sng" => "single", "dbl" => "double", _ => run.Underline };
        if (run.HasStrike) style["strike"] = run.Strike;
        if (run.HasFontKerningPoints) style["kerning"] = run.FontKerningPoints;
        if (run.HasFontBaselinePercent) style["baseline"] = run.FontBaselinePercent;
        if (run.HasFontSpacingPoints) style["letterSpacing"] = run.FontSpacingPoints;
        if (run.HasFontCaps) style["capitalization"] = run.FontCaps;
        if (run.HasLanguage) style["language"] = run.Language;
        return style;
    }

    private static JsonObject ProjectTextStyle(PresentationTextStyle source)
    {
        var style = new JsonObject();
        if (source.HasFontFamily) style["fontFamily"] = source.FontFamily;
        if (source.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = source.FontFamilyEastAsia;
        if (source.HasFontSizePoints && source.FontSizePoints > 0) style["size"] = source.FontSizePoints;
        if (source.HasBold) style["bold"] = source.Bold;
        if (source.HasItalic) style["italic"] = source.Italic;
        if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorRgb && !string.IsNullOrEmpty(source.ColorRgb))
            style["color"] = TextColor(source.ColorRgb, null, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorScheme && !string.IsNullOrEmpty(source.ColorScheme))
            style["color"] = TextColor(null, source.ColorScheme, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.GradientFill is not null)
            style["gradient"] = TextGradient(source.GradientFill);
        if (source.Shadow is not null) style["shadow"] = Shadow(source.Shadow);
        if (source.HighlightCase == PresentationTextStyle.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(source.HighlightRgb))
            style["highlight"] = Color(source.HighlightRgb);
        if (source.HasUnderline) style["underline"] = source.Underline switch { "sng" => "single", "dbl" => "double", _ => source.Underline };
        if (source.HasStrike) style["strike"] = source.Strike;
        if (source.HasFontKerningPoints) style["kerning"] = source.FontKerningPoints;
        if (source.HasFontBaselinePercent) style["baseline"] = source.FontBaselinePercent;
        if (source.HasFontSpacingPoints) style["letterSpacing"] = source.FontSpacingPoints;
        if (source.HasFontCaps) style["capitalization"] = source.FontCaps;
        if (source.HasLanguage) style["language"] = source.Language;
        return style;
    }

    private static JsonObject? ProjectBullet(PresentationTextParagraph paragraph)
    {
        JsonObject? bullet = paragraph.BulletCase switch
        {
            PresentationTextParagraph.BulletOneofCase.NoBullet => new JsonObject { ["type"] = "none" },
            PresentationTextParagraph.BulletOneofCase.BulletCharacter => new JsonObject
            {
                ["type"] = "character",
                ["character"] = paragraph.BulletCharacter,
            },
            PresentationTextParagraph.BulletOneofCase.AutoNumber => new JsonObject
            {
                ["type"] = "number",
                ["scheme"] = paragraph.AutoNumber.Scheme,
            },
            _ => null,
        };
        if (bullet is null) return null;
        // Font, color, and size metadata on a noBullet paragraph are inert
        // native residue. Do not emit them into the closed `none` PPJ shape;
        // the source-bound leaves still preserve and edit those tokens.
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.NoBullet)
            return bullet;
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber && paragraph.AutoNumber.HasStartAt)
            bullet["startAt"] = checked((int)paragraph.AutoNumber.StartAt);
        if (paragraph.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily)
            bullet["fontFamily"] = paragraph.BulletFontFamily;
        if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb)
            bullet["color"] = TextColor(
                paragraph.BulletColorRgb,
                null,
                paragraph.HasBulletColorOpacityThousandthPercent,
                paragraph.BulletColorOpacityThousandthPercent);
        else if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme)
            bullet["color"] = TextColor(
                null,
                paragraph.BulletColorScheme,
                paragraph.HasBulletColorOpacityThousandthPercent,
                paragraph.BulletColorOpacityThousandthPercent);
        if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints)
            bullet["size"] = paragraph.BulletSizePoints;
        else if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent)
            bullet["sizePercent"] = paragraph.BulletSizePercent;
        return bullet;
    }

    private static JsonArray ImportedThemeColors()
    {
        var output = new JsonArray();
        foreach (var item in new[]
        {
            (Id: "source-neutral", Value: "#000000"),
            (Id: "bg1", Value: "#FFFFFF"),
            (Id: "tx1", Value: "#000000"),
            (Id: "bg2", Value: "#EEECE1"),
            (Id: "tx2", Value: "#1F497D"),
            (Id: "accent1", Value: "#4F81BD"),
            (Id: "accent2", Value: "#C0504D"),
            (Id: "accent3", Value: "#9BBB59"),
            (Id: "accent4", Value: "#8064A2"),
            (Id: "accent5", Value: "#4BACC6"),
            (Id: "accent6", Value: "#F79646"),
            (Id: "hlink", Value: "#0000FF"),
            (Id: "folHlink", Value: "#800080"),
            (Id: "dk1", Value: "#000000"),
            (Id: "lt1", Value: "#FFFFFF"),
            (Id: "dk2", Value: "#1F497D"),
            (Id: "lt2", Value: "#EEECE1"),
        })
            output.Add(new JsonObject
            {
                ["id"] = item.Id,
                ["value"] = item.Value,
                ["role"] = "Fallback only; imported native styling remains source-owned",
            });
        return output;
    }

    private static JsonObject TextBoxStyle(PresentationTextBody? body)
    {
        var output = new JsonObject();
        var properties = body?.BodyProperties;
        if (properties is null) return output;
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor)
            output["verticalAlignment"] = properties.VerticalAnchor == "center" ? "middle" : properties.VerticalAnchor;
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
            output["wrap"] = properties.Wrap;
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode)
            output["autoFit"] = properties.AutoFitMode switch { "shrinkText" => "shrink-text", "resizeShape" => "resize-shape", _ => "none" };
        var margins = new JsonObject();
        if (properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) margins["left"] = Points(properties.LeftInsetEmu);
        if (properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) margins["top"] = Points(properties.TopInsetEmu);
        if (properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) margins["right"] = Points(properties.RightInsetEmu);
        if (properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) margins["bottom"] = Points(properties.BottomInsetEmu);
        if (margins.Count > 0) output["margins"] = margins;
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
            output["columns"] = checked((int)properties.Columns);
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu)
            output["columnGap"] = Points(properties.ColumnSpacingEmu);
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns)
            output["columnDirection"] = properties.RightToLeftColumns ? "right-to-left" : "left-to-right";
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode)
            output["verticalText"] = properties.VerticalTextMode;
        return output;
    }

    private static string? ParagraphAlignment(string value) => value switch
    {
        "l" or "left" => "left",
        "ctr" or "center" => "center",
        "r" or "right" => "right",
        "just" or "justify" => "justify",
        "dist" or "distributed" => "distributed",
        _ => null,
    };

    private static void ApplyTextContainerStyle(JsonObject output, PresentationShape shape, ProjectionContext context)
    {
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = "solid", ["color"] = Color(shape.FillRgb) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = Unit(shape.FillOpacityThousandthPercent);
            output["fill"] = fill;
        }
        else if (shape.GradientFill is not null)
        {
            output["fill"] = Gradient(shape.GradientFill);
        }
        else if (shape.ImageFill is not null && ProjectImagePaint(shape.ImageFill, context) is { } imageFill)
        {
            output["fill"] = imageFill;
        }
        if ((!string.IsNullOrEmpty(shape.LineRgb) || !string.IsNullOrEmpty(shape.LineScheme)) && shape.LineStyle != "none")
            output["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin,
                shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                shape.LineScheme);
    }

    private static JsonObject? ProjectBackground(PresentationBackground? background, ProjectionContext context)
    {
        if (background is null) return null;
        if (background.ImagePaint is not null)
            return ProjectImagePaint(background.ImagePaint, context);
        if (!string.IsNullOrEmpty(background.ImageAssetId) && context.TryMaterializeAsset(background.ImageAssetId, out var assetId))
            return new JsonObject { ["type"] = "image", ["asset"] = assetId, ["fit"] = "stretch" };
        if (background.GradientFill is not null)
            return Gradient(background.GradientFill);
        if (!string.IsNullOrEmpty(background.ColorRgb))
        {
            var output = new JsonObject { ["type"] = "solid", ["color"] = Color(background.ColorRgb) };
            if (background.HasOpacityThousandthPercent)
                output["opacity"] = Unit(background.OpacityThousandthPercent);
            return output;
        }
        return null;
    }

    private static JsonObject? ProjectImagePaint(PresentationImagePaint paint, ProjectionContext context)
    {
        if (!context.TryMaterializeAsset(paint.AssetId, out var assetId)) return null;
        var output = new JsonObject
        {
            ["type"] = "image",
            ["asset"] = assetId,
            ["fit"] = paint.Mode == PresentationImagePaint.Types.Mode.Tile ? "tile" : "stretch",
        };
        if (paint.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = Crop(paint.Crop.LeftThousandthPercent),
                ["top"] = Crop(paint.Crop.TopThousandthPercent),
                ["right"] = Crop(paint.Crop.RightThousandthPercent),
                ["bottom"] = Crop(paint.Crop.BottomThousandthPercent),
            };
        }
        if (paint.HasOpacityThousandthPercent)
            output["opacity"] = Unit(paint.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject Gradient(PresentationGradientFill source)
    {
        var stops = new JsonArray();
        foreach (var stop in source.Stops)
        {
            var item = new JsonObject
            {
                ["offset"] = Unit(stop.PositionThousandthPercent),
                ["color"] = Color(stop.ColorRgb),
            };
            if (stop.HasOpacityThousandthPercent)
                item["opacity"] = Unit(stop.OpacityThousandthPercent);
            stops.Add(item);
        }
        var output = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = source.Kind == PresentationGradientFill.Types.Kind.Radial ? "radial" : "linear",
            ["stops"] = stops,
        };
        if (source.Kind == PresentationGradientFill.Types.Kind.Linear && source.HasAngle60000)
            output["angle"] = source.Angle60000 / 60_000d;
        return output;
    }

    private static JsonObject TextGradient(PresentationGradientFill source)
    {
        var output = Gradient(source);
        output.Remove("type");
        return output;
    }

    private static JsonArray ProjectAnimations(PresentationSlide slide, ProjectionContext context)
    {
        var output = new JsonArray();
        var pageId = context.PageId(slide.Id);
        foreach (var animation in slide.Animations)
        {
            if (!context.TryElementId(pageId, animation.TargetId, out var targetId)) continue;
            var item = new JsonObject
            {
                ["id"] = context.UniqueId($"animation-{animation.Id}"),
                ["target"] = targetId,
                ["phase"] = animation.Phase,
                ["effect"] = animation.Effect,
                ["start"] = animation.Start,
                ["durationMs"] = animation.HasDurationMs ? checked((int)animation.DurationMs) : 500,
            };
            if (!string.IsNullOrEmpty(animation.Direction)) item["direction"] = animation.Direction;
            if (animation.HasDelayMs) item["delayMs"] = checked((int)animation.DelayMs);
            if (!string.IsNullOrEmpty(animation.TextBuild)) item["textBuild"] = animation.TextBuild;
            if (!string.IsNullOrEmpty(animation.ChartBuild)) item["chartBuild"] = animation.ChartBuild;
            if (animation.HasStaggerMs) item["staggerMs"] = checked((int)animation.StaggerMs);
            if (animation.HasAnimateChartBackground) item["animateChartBackground"] = animation.AnimateChartBackground;
            output.Add(item);
        }
        return output;
    }

    private static JsonObject? ProjectTransition(
        PresentationSlide slide,
        PresentationArtifact presentation,
        ProjectionContext context)
    {
        if (slide.Morph is not null && slide.Morph.Pairs.Count > 0)
        {
            var fromPage = presentation.Slides.FirstOrDefault(item => item.Id == slide.Morph.FromSlideId);
            if (fromPage is null) return null;
            var pairs = new JsonArray();
            foreach (var pair in slide.Morph.Pairs)
            {
                if (!context.TryElementId(context.PageId(fromPage.Id), pair.FromId, out var fromId) ||
                    !context.TryElementId(context.PageId(slide.Id), pair.ToId, out var toId)) continue;
                pairs.Add(new JsonObject { ["key"] = context.UniqueId($"morph-{pair.Key}"), ["from"] = fromId, ["to"] = toId });
            }
            if (pairs.Count == 0) return null;
            return new JsonObject
            {
                ["type"] = "morph",
                ["durationMs"] = slide.Morph.HasDurationMs ? checked((int)slide.Morph.DurationMs) : 800,
                ["fromPage"] = context.PageId(fromPage.Id),
                ["morphPairs"] = pairs,
            };
        }
        var transition = slide.Transition;
        if (transition is null || !PpjTransitionLowering.IsBaseEffect(transition.Effect)) return null;
        var output = new JsonObject
        {
            ["type"] = transition.Effect,
            ["speed"] = transition.Speed,
            ["advanceOnClick"] = transition.AdvanceOnClick,
        };
        if (transition.HasDurationMs) output["durationMs"] = checked((int)transition.DurationMs);
        if (!string.IsNullOrEmpty(transition.Direction)) output["direction"] = transition.Direction;
        if (!string.IsNullOrEmpty(transition.Orientation)) output["orientation"] = transition.Orientation;
        if (transition.HasThroughBlack) output["throughBlack"] = transition.ThroughBlack;
        if (transition.HasSpokes) output["spokes"] = checked((int)transition.Spokes);
        if (transition.HasAdvanceAfterMs) output["advanceAfterMs"] = checked((int)transition.AdvanceAfterMs);
        return output;
    }

    private static JsonArray ProjectSections(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var section in presentation.Sections)
        {
            var pages = new JsonArray();
            foreach (var id in section.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            output.Add(new JsonObject
            {
                ["id"] = context.UniqueId($"section-{section.Id}"),
                ["name"] = string.IsNullOrWhiteSpace(section.Name) ? "Section" : section.Name,
                ["pages"] = pages,
            });
        }
        return output;
    }

    private static JsonArray ProjectCustomShows(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var show in presentation.CustomShows)
        {
            var pages = new JsonArray();
            foreach (var id in show.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            output.Add(new JsonObject
            {
                ["id"] = context.UniqueId($"show-{show.Id}"),
                ["name"] = string.IsNullOrWhiteSpace(show.Name) ? "Custom show" : show.Name,
                ["pages"] = pages,
            });
        }
        return output;
    }

    private static JsonArray ProjectComments(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.PageId(slide.Id);
            foreach (var comment in slide.LegacyComments)
            {
                if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Text)) continue;
                output.Add(new JsonObject
                {
                    ["id"] = context.UniqueId($"comment-{pageId}-{comment.Id}"),
                    ["page"] = pageId,
                    ["author"] = string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author,
                    ["text"] = comment.Text,
                    ["createdAt"] = createdAt.ToUniversalTime().ToString("O"),
                    ["resolved"] = false,
                    ["position"] = new JsonObject { ["x"] = Points(comment.PositionXEmu), ["y"] = Points(comment.PositionYEmu) },
                });
            }
        }
        return output;
    }

    private static IReadOnlyList<CapabilitySpec> Capabilities(PresentationElement element)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                if (source.TextEditable && TextTopologyRepresentable(element.Shape.TextBody))
                    output.Add(new("replaceText", ["text"]));
                if (source.Editable)
                {
                    // Source-bound image-filled custom geometry is projected
                    // for discovery and frame/stroke edits, but its native
                    // fill graph is not represented by a lossless PPJ
                    // replacement operation. Do not advertise setFill for
                    // that bounded shape profile.
                    if (string.IsNullOrWhiteSpace(element.Shape.ImageFillAssetId) ||
                        element.Shape.Geometry != "custom")
                        output.Add(new("setFill", ["fill"]));
                    output.Add(new("setStroke", ["stroke"]));
                    output.Add(new("setFrame", element.Shape.Placeholder is null ? EditableFrameFields : PositionFrameFields));
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or "custom") &&
                        PptxPresetGeometryAdjustmentCodec.TryExpectedCount(element.Shape.Geometry, out var adjustmentCount) &&
                        adjustmentCount > 0)
                        output.Add(new("setGeometry", ["geometry.adjustments"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Image when source.Editable:
                output.Add(new("replaceImage", ["image.asset"]));
                output.Add(new("setImageCrop", ["image.crop"]));
                output.Add(new("setImageFit", ["image.fit"]));
                output.Add(new("setFrame", EditableFrameFields));
                output.Add(new("setOpacity", ["opacity"]));
                if (!string.IsNullOrEmpty(element.Image.MaskPreset) &&
                    PptxPresetGeometryAdjustmentCodec.TryExpectedCount(element.Image.MaskPreset, out var maskAdjustmentCount) &&
                    maskAdjustmentCount > 0)
                    output.Add(new("setImageMask", ["image.mask.adjustments"]));
                break;
            case PresentationElement.ContentOneofCase.Chart when source.Editable:
                output.Add(new("setChartTitle", ["chart.title"]));
                output.Add(new("setChartData", ["chart.data"]));
                output.Add(new("setChartTextStyle", ["chart.textStyle"]));
                output.Add(new("setChartFill", ["chart.fill"]));
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Table when source.Editable:
                output.Add(new("replaceText", ["text"]));
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Connector when source.Editable:
                output.Add(new("setStroke", ["stroke"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Group when source.Editable:
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
        }
        if (source.VisibilityEditable) output.Add(new("setHidden", ["hidden"]));
        if (source.LockingEditable) output.Add(new("setLocked", ["locked"]));
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static bool TextTopologyRepresentable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Text));

    private static IReadOnlyList<CapabilitySpec> OpaqueCapabilities(PresentationElement element)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        if (element.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
            PpjNativeTextProjection.TryRead(element.Opaque.RawXml, out _))
            output.Add(new("replaceText", ["visibleText"]));
        if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
        if (source.VisibilityEditable) output.Add(new("setHidden", ["hidden"]));
        if (source.LockingEditable) output.Add(new("setLocked", ["locked"]));
        // Diagram, OLE, and source-owned chart payloads stay opaque until the
        // corresponding PPJ typed state is projected. A runtime can edit such
        // objects only after the public language has somewhere to represent
        // both the old and requested value.
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static JsonObject NativeRef(
        ProjectionContext context,
        string scope,
        string objectHash,
        IEnumerable<CapabilitySpec> capabilities,
        JsonArray? leaves = null)
    {
        var capabilityArray = new JsonArray();
        foreach (var capability in capabilities
                     .OrderBy(item => item.Operation, StringComparer.Ordinal)
                     .ThenBy(item => string.Join("\0", item.Fields), StringComparer.Ordinal))
        {
            var fields = new JsonArray();
            foreach (var field in capability.Fields) fields.Add(StringNode(field));
            capabilityArray.Add(new JsonObject
            {
                ["id"] = $"cap-{capability.Operation}-{Sha256(Encoding.UTF8.GetBytes(scope + capability.Operation))[..10]}",
                ["operation"] = capability.Operation,
                ["expectedHash"] = objectHash,
                ["fields"] = fields,
            });
        }
        var output = new JsonObject
        {
            ["handle"] = $"nr-{Sha256(Encoding.UTF8.GetBytes(context.SourceSha256 + "\0" + scope))}",
            ["sourceSha256"] = context.SourceSha256,
            ["revision"] = context.Revision,
            ["objectHash"] = objectHash,
            ["capabilitySetSha256"] = Sha256(CanonicalBytes(capabilityArray)),
            ["capabilities"] = capabilityArray,
        };
        if (leaves is { Count: > 0 }) output["leaves"] = leaves;
        return output;
    }

    private static JsonObject ShapeFrame(PresentationShape shape)
    {
        var frame = Frame(shape.LeftEmu, shape.TopEmu, shape.WidthEmu, shape.HeightEmu);
        if (shape.Transform?.HasRotationAngle60000 == true) frame["rotation"] = shape.Transform.RotationAngle60000 / 60_000d;
        if (shape.Transform?.HasFlipHorizontal == true) frame["flipH"] = shape.Transform.FlipHorizontal;
        if (shape.Transform?.HasFlipVertical == true) frame["flipV"] = shape.Transform.FlipVertical;
        return frame;
    }

    private static readonly string[] EditableFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height", "frame.rotation", "frame.flipH", "frame.flipV"];

    private static readonly string[] PositionFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height"];

    private static JsonObject ImageFrame(PresentationImage image)
    {
        var frame = Frame(image.LeftEmu, image.TopEmu, image.WidthEmu, image.HeightEmu);
        if (image.Transform?.HasRotationAngle60000 == true) frame["rotation"] = image.Transform.RotationAngle60000 / 60_000d;
        if (image.Transform?.HasFlipHorizontal == true) frame["flipH"] = image.Transform.FlipHorizontal;
        if (image.Transform?.HasFlipVertical == true) frame["flipV"] = image.Transform.FlipVertical;
        return frame;
    }

    private static JsonObject TableFrame(PresentationTable table) =>
        Frame(table.LeftEmu, table.TopEmu, table.WidthEmu, table.HeightEmu, table.FrameTransform);

    private static JsonObject ChartFrame(PresentationChart chart) =>
        Frame(chart.LeftEmu, chart.TopEmu, chart.WidthEmu, chart.HeightEmu, chart.FrameTransform);

    private static JsonObject GroupFrame(PresentationGroup group) =>
        Frame(group.LeftEmu, group.TopEmu, group.WidthEmu, group.HeightEmu, group.FrameTransform);

    private static JsonObject ConnectorFrame(PresentationConnector connector)
    {
        var left = Math.Min(connector.StartXEmu, connector.EndXEmu);
        var top = Math.Min(connector.StartYEmu, connector.EndYEmu);
        return Frame(left, top, Math.Abs(connector.EndXEmu - connector.StartXEmu), Math.Abs(connector.EndYEmu - connector.StartYEmu));
    }

    private static JsonObject ElementFrame(PresentationElement element) => element.ContentCase switch
    {
        PresentationElement.ContentOneofCase.Shape => ShapeFrame(element.Shape),
        PresentationElement.ContentOneofCase.Image => ImageFrame(element.Image),
        PresentationElement.ContentOneofCase.Table => TableFrame(element.Table),
        PresentationElement.ContentOneofCase.Connector => ConnectorFrame(element.Connector),
        PresentationElement.ContentOneofCase.Chart => ChartFrame(element.Chart),
        PresentationElement.ContentOneofCase.Group => GroupFrame(element.Group),
        PresentationElement.ContentOneofCase.Opaque => Frame(element.Opaque.LeftEmu, element.Opaque.TopEmu, element.Opaque.WidthEmu, element.Opaque.HeightEmu),
        _ => Frame(0, 0, 1, 1),
    };

    private static JsonObject Frame(long left, long top, long width, long height) => new()
    {
        ["x"] = Points(left),
        ["y"] = Points(top),
        ["width"] = Math.Max(0.001, Points(width)),
        ["height"] = Math.Max(0.001, Points(height)),
    };

    private static JsonObject Frame(
        long left,
        long top,
        long width,
        long height,
        PresentationFrameTransform? transform)
    {
        var frame = Frame(left, top, width, height);
        if (transform?.HasRotationAngle60000 == true) frame["rotation"] = transform.RotationAngle60000 / 60_000d;
        if (transform?.HasFlipHorizontal == true) frame["flipH"] = transform.FlipHorizontal;
        if (transform?.HasFlipVertical == true) frame["flipV"] = transform.FlipVertical;
        return frame;
    }

    private static JsonObject ConnectorEndpoint(
        string targetId,
        long x,
        long y,
        string pageId,
        ProjectionContext context)
    {
        if (!string.IsNullOrEmpty(targetId) && context.TryElementId(pageId, targetId, out var projected))
            return new JsonObject { ["element"] = projected, ["anchor"] = "auto" };
        return new JsonObject { ["x"] = Points(x), ["y"] = Points(y) };
    }

    private static JsonObject Stroke(
        string rgb,
        long widthEmu,
        string? style,
        string? cap,
        string? join,
        double? opacity,
        string? scheme = null)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(scheme)
                ? new JsonObject { ["token"] = scheme }
                : Color(string.IsNullOrEmpty(rgb) ? "000000" : rgb),
            ["width"] = Math.Max(0, Points(widthEmu)),
        };
        var dash = Dash(style);
        if (dash is not null) output["dash"] = StringNode(dash);
        if (cap is "flat" or "round" or "square") output["cap"] = StringNode(cap);
        if (join is "miter" or "round" or "bevel") output["join"] = StringNode(join);
        if (opacity is not null) output["opacity"] = opacity.Value;
        return output;
    }

    private static JsonObject Shadow(PresentationShadow shadow)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(shadow.ColorScheme)
                ? new JsonObject { ["token"] = shadow.ColorScheme }
                : Color(string.IsNullOrEmpty(shadow.ColorRgb) ? "000000" : shadow.ColorRgb),
            ["opacity"] = shadow.HasOpacityThousandthPercent ? Unit(shadow.OpacityThousandthPercent) : 1,
            ["blur"] = Math.Max(0, Points(shadow.HasBlurRadiusEmu ? shadow.BlurRadiusEmu : 0)),
            ["distance"] = Math.Max(0, Points(shadow.HasDistanceEmu ? shadow.DistanceEmu : 0)),
            ["angle"] = (shadow.HasDirectionAngle60000 ? shadow.DirectionAngle60000 : 0) / 60_000d,
        };
        if (shadow.HasAlignment) output["alignment"] = shadow.Alignment;
        if (shadow.HasRotateWithShape) output["rotateWithShape"] = shadow.RotateWithShape;
        return output;
    }

    private static JsonObject? Accessibility(PresentationNonVisualAccessibility? value)
    {
        if (value is null) return null;
        var output = new JsonObject { ["decorative"] = value.HasDecorative && value.Decorative };
        if (!string.IsNullOrEmpty(value.Title)) output["title"] = value.Title;
        if (!string.IsNullOrEmpty(value.Description)) output["description"] = value.Description;
        return output;
    }

    private static JsonObject? ImageAccessibility(PresentationImage image)
    {
        if (!image.HasAccessibilityDecorative && string.IsNullOrEmpty(image.AccessibilityTitle) && string.IsNullOrEmpty(image.AltText)) return null;
        var output = new JsonObject { ["decorative"] = image.HasAccessibilityDecorative && image.AccessibilityDecorative };
        if (!string.IsNullOrEmpty(image.AccessibilityTitle)) output["title"] = image.AccessibilityTitle;
        if (!string.IsNullOrEmpty(image.AltText)) output["description"] = image.AltText;
        return output;
    }

    private static string PlaceholderType(string value) => value switch
    {
        "ctrTitle" or "title" => "title",
        "subTitle" or "subtitle" => "subtitle",
        "body" => "body",
        "obj" or "content" => "content",
        "pic" or "picture" => "picture",
        "chart" => "chart",
        "tbl" or "table" => "table",
        "dt" or "date" => "date",
        "ftr" or "footer" => "footer",
        "sldNum" or "slide-number" => "slide-number",
        _ => "other",
    };

    private static string? ChartType(SpreadsheetChartType type) => type switch
    {
        SpreadsheetChartType.Bar => "column",
        SpreadsheetChartType.Line => "line",
        SpreadsheetChartType.Pie => "pie",
        SpreadsheetChartType.Area => "area",
        SpreadsheetChartType.Doughnut => "doughnut",
        SpreadsheetChartType.Scatter => "scatter",
        SpreadsheetChartType.Bubble => "bubble",
        SpreadsheetChartType.Radar => "radar",
        SpreadsheetChartType.Combo => "combo",
        _ => null,
    };

    private static string? Marker(SpreadsheetChartMarkerSymbol value) => value switch
    {
        SpreadsheetChartMarkerSymbol.None => "none",
        SpreadsheetChartMarkerSymbol.Dot => "dot",
        SpreadsheetChartMarkerSymbol.Circle => "circle",
        SpreadsheetChartMarkerSymbol.Square => "square",
        SpreadsheetChartMarkerSymbol.Diamond => "diamond",
        SpreadsheetChartMarkerSymbol.Triangle => "triangle",
        SpreadsheetChartMarkerSymbol.X => "x",
        SpreadsheetChartMarkerSymbol.Star => "star",
        SpreadsheetChartMarkerSymbol.Plus => "plus",
        SpreadsheetChartMarkerSymbol.Dash => "dash",
        _ => null,
    };

    private static string? ChartDash(SpreadsheetChartLineDashStyle value) => value switch
    {
        SpreadsheetChartLineDashStyle.Solid => "solid",
        SpreadsheetChartLineDashStyle.Dashed => "dash",
        SpreadsheetChartLineDashStyle.Dotted => "dot",
        SpreadsheetChartLineDashStyle.DashDot => "dash-dot",
        _ => null,
    };

    private static string? DataLabelPosition(SpreadsheetChartDataLabelPosition value) => value switch
    {
        SpreadsheetChartDataLabelPosition.BestFit => "best-fit",
        SpreadsheetChartDataLabelPosition.Bottom => "bottom",
        SpreadsheetChartDataLabelPosition.Center => "center",
        SpreadsheetChartDataLabelPosition.InsideBase => "inside-base",
        SpreadsheetChartDataLabelPosition.InsideEnd => "inside-end",
        SpreadsheetChartDataLabelPosition.Left => "left",
        SpreadsheetChartDataLabelPosition.OutsideEnd => "outside-end",
        SpreadsheetChartDataLabelPosition.Right => "right",
        SpreadsheetChartDataLabelPosition.Top => "top",
        _ => null,
    };

    private static string? TrendlineType(SpreadsheetChartTrendlineType value) => value switch
    {
        SpreadsheetChartTrendlineType.Exponential => "exponential",
        SpreadsheetChartTrendlineType.Linear => "linear",
        SpreadsheetChartTrendlineType.Logarithmic => "logarithmic",
        SpreadsheetChartTrendlineType.MovingAverage => "moving-average",
        SpreadsheetChartTrendlineType.Polynomial => "polynomial",
        SpreadsheetChartTrendlineType.Power => "power",
        _ => null,
    };

    private static string? ErrorBarDirection(SpreadsheetChartErrorBarDirection value) => value switch
    {
        SpreadsheetChartErrorBarDirection.X => "x",
        SpreadsheetChartErrorBarDirection.Y => "y",
        _ => null,
    };

    private static string? ErrorBarType(SpreadsheetChartErrorBarType value) => value switch
    {
        SpreadsheetChartErrorBarType.Both => "both",
        SpreadsheetChartErrorBarType.Minus => "minus",
        SpreadsheetChartErrorBarType.Plus => "plus",
        _ => null,
    };

    private static string? ErrorBarValueType(SpreadsheetChartErrorBarValueType value) => value switch
    {
        SpreadsheetChartErrorBarValueType.FixedValue => "fixed-value",
        SpreadsheetChartErrorBarValueType.Percentage => "percentage",
        SpreadsheetChartErrorBarValueType.StandardDeviation => "standard-deviation",
        SpreadsheetChartErrorBarValueType.StandardError => "standard-error",
        _ => null,
    };

    private static string? Dash(string? value) => value switch
    {
        null or "" or "solid" => "solid",
        "dashed" or "dash" => "dash",
        "dotted" or "dot" => "dot",
        "dash-dot" or "dashDot" => "dash-dot",
        "long-dash" or "longDash" => "long-dash",
        _ => null,
    };

    private static string? Arrow(string? value) => value switch
    {
        "triangle" or "stealth" or "diamond" or "oval" or "open" => value,
        _ => null,
    };

    private static string Color(string rgb) => $"#{rgb.TrimStart('#').ToUpperInvariant()}";

    private static JsonNode TextColor(string? rgb, string? scheme, bool hasOpacity, uint opacity)
    {
        if (!string.IsNullOrEmpty(rgb))
        {
            var value = Color(rgb);
            if (!hasOpacity) return value;
            var alpha = Math.Clamp((int)Math.Round(Unit(opacity) * 255), 0, 255);
            return $"{value}{alpha:X2}";
        }
        var output = new JsonObject { ["token"] = scheme };
        if (hasOpacity) output["alpha"] = Unit(opacity);
        return output;
    }
    private static double Crop(int value) => Math.Clamp(value / 100_000d, -1, 1);
    private static double Unit(uint value) => Math.Clamp(value / 100_000d, 0, 1);
    private static double Points(long emu) => Math.Round(emu / EmuPerPoint, 6, MidpointRounding.AwayFromZero);

    private static string HashOrFallback(string? hash, byte[] fallback) =>
        hash is { Length: 64 } && hash.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f')
            ? hash
            : Sha256(fallback);

    private static string StableDocumentId(string? candidate, string sha256)
    {
        var normalized = NormalizeId(candidate, "presentation");
        return normalized.Length <= 100 ? normalized : $"presentation-{sha256[..24]}";
    }

    private static string NormalizeId(string? value, string fallback)
    {
        var normalized = InvalidIdCharacters().Replace(value ?? string.Empty, "-").Trim('-', '.', ':', '_');
        if (normalized.Length == 0 || !char.IsAsciiLetterOrDigit(normalized[0])) normalized = $"{fallback}-{normalized}".TrimEnd('-');
        if (normalized.Length > 112)
            normalized = $"{normalized[..87]}-{Sha256(Encoding.UTF8.GetBytes(value ?? fallback))[..24]}";
        return normalized;
    }

    private static byte[] CanonicalBytes(JsonNode node)
    {
        using var document = JsonDocument.Parse(node.ToJsonString());
        return PpjCanonicalJson.Write(document.RootElement);
    }

    private static JsonNode StringNode(string value)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            writer.WriteStringValue(value);
        return JsonNode.Parse(buffer.WrittenSpan) ?? throw new InvalidOperationException("String JSON primitive could not be created.");
    }

    private static JsonNode NumberNode(double value)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            writer.WriteNumberValue(value);
        return JsonNode.Parse(buffer.WrittenSpan) ?? throw new InvalidOperationException("Number JSON primitive could not be created.");
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    [GeneratedRegex("[^A-Za-z0-9._:-]+")]
    private static partial Regex InvalidIdCharacters();

    private sealed record CapabilitySpec(string Operation, IReadOnlyList<string> Fields);

    private sealed class ProjectionContext
    {
        private readonly IReadOnlyDictionary<string, Asset> sourceAssets;
        private readonly Dictionary<string, string> pageIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> masterIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> layoutIds = new(StringComparer.Ordinal);
        private readonly Dictionary<(string Page, string Element), string> elementIds = new();
        private readonly HashSet<string> usedIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdBySourceId = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdByHash = new(StringComparer.Ordinal);
        private readonly JsonArray programAssets = new();
        private readonly List<Asset> resultAssets = [];
        private readonly JsonArray nodes = new();
        private readonly Dictionary<string, PpjNativeLeafBinding> nativeLeafBindings = new(StringComparer.Ordinal);

        internal ProjectionContext(
            string sourceSha256,
            string revision,
            string assetRoot,
            IEnumerable<Asset> assets)
        {
            SourceSha256 = sourceSha256;
            Revision = revision;
            AssetRoot = assetRoot;
            sourceAssets = assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        }

        internal string SourceSha256 { get; }
        internal string Revision { get; }
        internal IReadOnlyDictionary<string, PpjNativeLeafBinding> NativeLeafBindings => nativeLeafBindings;

        internal void RecordNativeLeaf(PpjNativeLeafBinding binding)
        {
            if (!nativeLeafBindings.TryAdd(binding.Id, binding))
                throw new CodecException("ppj.nativeRef.leafId", $"Duplicate projected native leaf ID {binding.Id}.");
        }
        internal string AssetRoot { get; }
        internal int VisibleObjectCount { get; private set; }
        internal JsonArray ProgramAssets => programAssets;
        internal IReadOnlyList<Asset> ResultAssets => resultAssets;

        internal string RegisterPage(string sourceId)
        {
            var id = UniqueId($"page-{NormalizeId(sourceId, "slide")}");
            pageIds[sourceId] = id;
            return id;
        }

        internal string RegisterMaster(string sourceId)
        {
            var id = UniqueId($"master-{NormalizeId(sourceId, "master")}");
            masterIds[sourceId] = id;
            return id;
        }

        internal string RegisterLayout(string sourceId)
        {
            var id = UniqueId($"layout-{NormalizeId(sourceId, "layout")}");
            layoutIds[sourceId] = id;
            return id;
        }

        internal string RegisterElement(string pageId, string sourceId)
        {
            var id = UniqueId($"{pageId}-{NormalizeId(sourceId, "element")}");
            elementIds[(pageId, sourceId)] = id;
            return id;
        }

        internal string PageId(string sourceId) => pageIds[sourceId];
        internal bool TryPageId(string sourceId, out string id) => pageIds.TryGetValue(sourceId, out id!);
        internal bool TryLayoutId(string sourceId, out string id) => layoutIds.TryGetValue(sourceId, out id!);
        internal string ElementId(string pageId, string sourceId) => elementIds[(pageId, sourceId)];
        internal bool TryElementId(string pageId, string sourceId, out string id) => elementIds.TryGetValue((pageId, sourceId), out id!);

        internal string UniqueId(string candidate)
        {
            var normalized = NormalizeId(candidate, "id");
            if (usedIds.Add(normalized)) return normalized;
            var suffix = Sha256(Encoding.UTF8.GetBytes(candidate))[..12];
            var shortened = normalized.Length > 114 ? normalized[..114] : normalized;
            var unique = $"{shortened}-{suffix}";
            var ordinal = 2;
            while (!usedIds.Add(unique)) unique = $"{shortened}-{suffix}-{ordinal++}";
            return unique;
        }

        internal bool TryMaterializeAsset(string sourceId, out string programAssetId)
        {
            if (assetIdBySourceId.TryGetValue(sourceId, out programAssetId!)) return true;
            if (!sourceAssets.TryGetValue(sourceId, out var source) || source.Data.IsEmpty) return false;
            var hash = HashOrFallback(source.Sha256, source.Data.ToByteArray());
            if (assetIdByHash.TryGetValue(hash, out programAssetId!))
            {
                assetIdBySourceId[sourceId] = programAssetId;
                return true;
            }
            programAssetId = UniqueId($"asset-{hash[..20]}");
            assetIdBySourceId[sourceId] = programAssetId;
            assetIdByHash[hash] = programAssetId;
            var extension = Extension(source.ContentType, source.FileName);
            var fileName = $"{hash}.{extension}";
            programAssets.Add(new JsonObject
            {
                ["id"] = programAssetId,
                ["uri"] = $"{AssetRoot}/{fileName}",
                ["mimeType"] = source.ContentType,
                ["sha256"] = hash,
                ["rights"] = new JsonObject { ["status"] = "user-provided" },
                ["accessibility"] = new JsonObject { ["decorative"] = false, ["description"] = "Imported source media." },
            });
            var materialized = source.Clone();
            materialized.Id = programAssetId;
            materialized.FileName = fileName;
            materialized.Sha256 = hash;
            resultAssets.Add(materialized);
            return true;
        }

        internal void RecordNode(string pageId, string id, string type, JsonObject nativeRef)
        {
            VisibleObjectCount++;
            nodes.Add(new JsonObject
            {
                ["id"] = id,
                ["page"] = pageId,
                ["type"] = type,
                ["handle"] = nativeRef["handle"]!.GetValue<string>(),
                ["objectHash"] = nativeRef["objectHash"]!.GetValue<string>(),
                ["capabilitySetSha256"] = nativeRef["capabilitySetSha256"]!.GetValue<string>(),
            });
        }

        internal JsonObject BuildNodeMap() => new()
        {
            ["schema"] = "office-kit/ppj-node-map/v1",
            ["sourceSha256"] = SourceSha256,
            ["revision"] = Revision,
            ["nodes"] = nodes.DeepClone(),
        };

        private static string Extension(string contentType, string fileName)
        {
            var fromType = contentType.ToLowerInvariant() switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/svg+xml" => "svg",
                _ => string.Empty,
            };
            if (fromType.Length > 0) return fromType;
            var extension = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            return Regex.IsMatch(extension, "^[a-z0-9]{1,8}$") ? extension : "bin";
        }
    }
}
