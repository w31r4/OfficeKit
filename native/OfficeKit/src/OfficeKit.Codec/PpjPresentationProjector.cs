using System.Buffers;
using System.IO.Compression;
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
    IReadOnlyDictionary<string, PpjNativeLeafBinding> NativeLeafBindings,
    PpjValidationResult? Validation) : IDisposable
{
    public void Dispose() => Validation?.Dispose();
}

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
        EffectiveCodecLimits limits,
        bool retainSourceAssetData = true,
        string? verifiedSourceSha256 = null) => Project(
            new PptxPackageSource(sourceBytes),
            request,
            limits,
            retainSourceAssetData,
            verifiedSourceSha256);

    internal static PpjProjectionResult Project(
        PptxPackageSource source,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits,
        bool retainSourceAssetData = true,
        string? verifiedSourceSha256 = null)
    {
        if (PpjEmbeddedProgramCodec.TryRecover(source, request, limits) is { } recovered)
            return new(
                recovered.Program,
                recovered.Diagnostics,
                null,
                new Dictionary<string, PpjNativeLeafBinding>(StringComparer.Ordinal),
                null);

        var imported = PptxCodec.Import(
            source,
            limits,
            retainSourceAssetData,
            verifiedSourceSha256);
        var envelope = imported.Artifact;
        var presentation = envelope.Presentation ??
            throw new CodecException("ppj.projection.presentation", "The imported package did not produce a Presentation artifact.", "$");
        var sourceSha256 = envelope.Source?.PackageSha256;
        if (string.IsNullOrEmpty(sourceSha256))
            sourceSha256 = source.Sha256();
        var revision = $"pptx-{sourceSha256[..16]}";
        var sourceUri = string.IsNullOrWhiteSpace(request.SourceUri)
            ? $"deck.assets/source/{sourceSha256}.pptx"
            : request.SourceUri;
        var assetRoot = string.IsNullOrWhiteSpace(request.AssetRootUri)
            ? "deck.assets/media"
            : request.AssetRootUri.TrimEnd('/');

        var context = new ProjectionContext(sourceSha256, revision, assetRoot, envelope.Assets, envelope.OpaqueOpc, source);
        RegisterIds(presentation, context);

        var pages = new JsonArray();
        foreach (var slide in presentation.Slides)
            pages.Add(ProjectPage(slide, presentation, context));

        foreach (var nativeAsset in context.NativeSourceAssets)
            if (!envelope.Assets.Any(asset => asset.Id.Equals(nativeAsset.Id, StringComparison.Ordinal)))
                envelope.Assets.Add(nativeAsset.Clone());

        var assets = context.ProgramAssets;
        var sections = ProjectSections(presentation, context);
        var customShows = ProjectCustomShows(presentation, context);
        var comments = ProjectComments(presentation, context);
        var nodeMap = context.BuildNodeMap();
        var nodeMapBytes = CanonicalBytes(nodeMap);
        var projectionPayload = new JsonObject
        {
            ["canvas"] = FrameDimensions(presentation),
            ["assets"] = assets,
            ["pages"] = pages,
            ["sections"] = sections,
            ["customShows"] = customShows,
            ["comments"] = comments,
        };
        var projectionSha256 = Sha256(CanonicalBytes(projectionPayload));
        // JsonNode has single-parent ownership. The payload exists only to
        // bind the source-derived semantic graph, so release its children and
        // reuse those exact nodes in the public program instead of cloning the
        // full projection.
        projectionPayload.Clear();

        var root = new JsonObject
        {
            ["schema"] = StringNode("office-kit/ppj/v1"),
            ["meta"] = new JsonObject
            {
                ["id"] = StringNode(StableDocumentId(presentation.Id, sourceSha256)),
                ["title"] = StringNode(string.IsNullOrWhiteSpace(presentation.Name) ? "Imported presentation" : presentation.Name),
                ["language"] = StringNode("und"),
                ["version"] = JsonValue.Create(1),
                ["description"] = StringNode("Source-derived PPJ projection. Unmodeled native content remains in the hash-bound PPTX source package."),
            },
            ["intent"] = ImportedIntent(),
            ["design"] = ImportedDesign(presentation, context),
            ["assets"] = assets,
            ["source"] = new JsonObject
            {
                ["kind"] = StringNode("pptx"),
                ["uri"] = StringNode(sourceUri),
                ["sha256"] = StringNode(sourceSha256),
                ["revision"] = StringNode(revision),
                ["projection"] = new JsonObject
                {
                    ["version"] = JsonValue.Create(1),
                    ["sha256"] = StringNode(projectionSha256),
                    ["nodeMapSha256"] = StringNode(Sha256(nodeMapBytes)),
                    ["visibleObjectCount"] = JsonValue.Create(context.VisibleObjectCount),
                },
            },
            ["pages"] = pages,
        };
        if (sections.Count > 0) root["sections"] = sections;
        if (customShows.Count > 0) root["customShows"] = customShows;
        if (comments.Count > 0) root["comments"] = comments;

        var candidateBytes = CanonicalBytes(root);
        root.Clear();
        pages.Clear();
        sections.Clear();
        customShows.Clear();
        comments.Clear();
        nodeMap.Clear();
        context.ReleaseProjectionJson();
        var validation = PpjProgramValidator.Validate(candidateBytes);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            validation.Dispose();
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        var result = new PresentationProgramResult
        {
            ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? UnsafeByteOperations.UnsafeWrap(nodeMapBytes) : ByteString.Empty,
            SourceSha256 = sourceSha256,
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        result.Assets.Add(context.ResultAssets);
        return new(result, imported.Diagnostics, envelope, context.NativeLeafBindings, validation);
    }

    private static JsonObject ImportedIntent() => new()
    {
        ["brief"] = new JsonObject
        {
            ["primaryJob"] = StringNode("handoff"),
            ["expectedOutcome"] = StringNode("Continue editing this presentation while preserving source-owned native content."),
            ["evidenceBoundary"] = StringNode("The source presentation does not declare its audience, factual authority, or intended outcome."),
        },
        ["audience"] = new JsonObject
        {
            ["description"] = StringNode("Imported presentation audience was not declared."),
        },
        ["narrative"] = new JsonObject
        {
            ["thesis"] = StringNode("Preserve and continue the imported presentation without inventing its original intent."),
        },
        ["editorial"] = new JsonObject
        {
            ["tone"] = new JsonArray { StringNode("source-derived") },
            ["avoid"] = new JsonArray { StringNode("Inventing missing facts or silently rewriting source-owned content") },
        },
        ["delivery"] = new JsonObject
        {
            ["mode"] = StringNode("hybrid"),
            ["mediumFit"] = StringNode("acceptable"),
            ["mediumFitNote"] = StringNode("Delivery intent was not declared in the imported file."),
        },
    };

    private static JsonObject ImportedDesign(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonObject
        {
            ["canvas"] = ProjectCanvas(presentation, context),
            ["theme"] = new JsonObject
            {
                ["name"] = StringNode("Source-owned presentation theme"),
                // Imported run and bullet styles may retain a direct theme token
                // even though the source theme graph is opaque to the bounded
                // writer. Keep every standard token addressable with a neutral
                // fallback so the PPJ projection stays valid without claiming
                // that these fallback RGB values replace the source theme.
                ["colors"] = ImportedThemeColors(),
            },
            ["fonts"] = new JsonArray(new JsonObject
            {
                ["id"] = StringNode("source-font"),
                ["family"] = StringNode("Arial"),
                ["language"] = StringNode("und"),
            }),
            ["styles"] = new JsonObject(),
            ["grammar"] = new JsonObject
            {
                ["name"] = StringNode("Source-derived projection"),
                ["rationale"] = StringNode("The original PPTX remains the authority for native visual styling that PPJ does not model."),
                ["visualThesis"] = StringNode("Preserve and continue the imported presentation without inventing its original intent."),
                ["surfaceHierarchy"] = new JsonArray { StringNode("Keep source-owned surfaces unchanged unless a capability explicitly permits an edit.") },
                ["typographyRhythm"] = new JsonArray { StringNode("Retain imported typography through the source package.") },
                ["geometryRules"] = new JsonArray { StringNode("Retain imported geometry and z-order unless a nativeRef capability permits a change.") },
                ["densityRhythm"] = new JsonArray { StringNode("Retain the source page density.") },
                ["carrierRules"] = new JsonArray { StringNode("Use projected typed objects where safe and opaque native objects everywhere else.") },
                ["forbiddenPatterns"] = new JsonArray { StringNode("Rebuilding opaque content"), StringNode("Guessing unsupported native semantics") },
            },
            ["motionPolicy"] = StringNode("explicit"),
        };
        if (presentation.Masters.Count > 0)
            output["masters"] = ProjectMasters(presentation.Masters, context);
        if (presentation.Layouts.Count > 0)
            output["layouts"] = ProjectLayouts(presentation.Layouts, presentation.Masters, context);
        return output;
    }

    private static JsonArray ProjectMasters(
        IEnumerable<PresentationMaster> masters,
        ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var master in masters)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.MasterId(master.Id)),
                ["name"] = StringNode(master.Name),
            };
            if (master.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                if (master.Source.BackgroundEditable)
                    capabilities.Add(new("setBackground", ["background"]));
                if (master.Source.TextStylesEditable)
                    capabilities.Add(new("setTextParagraphStyle", ["textStyles"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"master:{master.Id}",
                    HashOrFallback(master.Source.MasterXmlSha256, master),
                    capabilities);
            }
            if (ProjectBackground(master.Background, context) is { } background)
                projected["background"] = background;
            var textStyles = new JsonObject();
            AddMasterTextLevels(textStyles, "title", master.TextStyles?.TitleLevels, context);
            AddMasterTextLevels(textStyles, "body", master.TextStyles?.BodyLevels, context);
            AddMasterTextLevels(textStyles, "other", master.TextStyles?.OtherLevels, context);
            if (textStyles.Count > 0) projected["textStyles"] = textStyles;
            var placeholders = ProjectLayoutPlaceholders(master.Placeholders, null, context);
            if (placeholders.Count > 0) projected["placeholders"] = placeholders;
            output.Add(projected);
        }
        return output;
    }

    private static JsonArray ProjectLayouts(
        IEnumerable<PresentationLayout> layouts,
        IEnumerable<PresentationMaster> masters,
        ProjectionContext context)
    {
        var mastersById = masters.ToDictionary(master => master.Id, StringComparer.Ordinal);
        var output = new JsonArray();
        foreach (var layout in layouts)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.LayoutId(layout.Id)),
                ["name"] = StringNode(layout.Name),
                ["master"] = StringNode(context.MasterId(layout.MasterId)),
                ["layoutType"] = StringNode(layout.Type),
            };
            if (layout.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                if (layout.Source.BackgroundEditable)
                    capabilities.Add(new("setBackground", ["background"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"layout:{layout.Id}",
                    HashOrFallback(layout.Source.LayoutXmlSha256, layout),
                    capabilities);
            }
            if (ProjectBackground(layout.Background, context) is { } background)
                projected["background"] = background;
            var inheritedPlaceholders = mastersById.TryGetValue(layout.MasterId, out var master)
                ? master.Placeholders
                : null;
            var placeholders = ProjectLayoutPlaceholders(layout.Placeholders, inheritedPlaceholders, context);
            if (placeholders.Count > 0) projected["placeholders"] = placeholders;
            output.Add(projected);
        }
        return output;
    }

    private static JsonArray ProjectLayoutPlaceholders(
        IEnumerable<PresentationPlaceholder> placeholders,
        IEnumerable<PresentationPlaceholder>? inheritedPlaceholders,
        ProjectionContext context)
    {
        var inheritedFrames = inheritedPlaceholders is null
            ? new Dictionary<(string Type, uint Index), PresentationPlaceholderFrame>()
            : inheritedPlaceholders
                .Where(placeholder => placeholder.DirectFrame is not null)
                .GroupBy(placeholder => (placeholder.Type, placeholder.Index))
                .Where(group => group.Count() == 1)
                .ToDictionary(
                    group => group.Key,
                    group => group.Single().DirectFrame!.Clone());
        var output = new JsonArray();
        foreach (var placeholder in placeholders)
        {
            var frameSource = placeholder.DirectFrame;
            if (frameSource is null &&
                inheritedFrames.TryGetValue((placeholder.Type, placeholder.Index), out var inheritedFrame))
                frameSource = inheritedFrame;
            // An inherited placeholder without a matching master frame, or an
            // irregular transform that was intentionally rejected by the
            // native reader, still has no safe PPJ frame. Keep it source-owned
            // instead of inventing coordinates.
            if (frameSource is null) continue;
            var frame = new JsonObject
            {
                ["x"] = JsonValue.Create(Points(frameSource.LeftEmu)),
                ["y"] = JsonValue.Create(Points(frameSource.TopEmu)),
                ["width"] = JsonValue.Create(Math.Max(0.001, Points(frameSource.WidthEmu))),
                ["height"] = JsonValue.Create(Math.Max(0.001, Points(frameSource.HeightEmu))),
            };
            if (frameSource.HasRotationAngle60000)
                frame["rotation"] = JsonValue.Create(frameSource.RotationAngle60000 / 60_000d);
            if (frameSource.HasFlipHorizontal)
                frame["flipH"] = JsonValue.Create(frameSource.FlipHorizontal);
            if (frameSource.HasFlipVertical)
                frame["flipV"] = JsonValue.Create(frameSource.FlipVertical);
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.UniqueId(placeholder.Id)),
                ["name"] = StringNode(placeholder.Name),
                ["placeholderType"] = StringNode(PlaceholderType(placeholder.Type)),
                ["index"] = JsonValue.Create(placeholder.Index),
                ["frame"] = frame,
            };
            if (placeholder.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                // A direct placeholder transform is an owner-local leaf.  The
                // source binding has already rejected inherited/irregular
                // geometry, so changing the existing frame cannot silently
                // move a slide placeholder or rewrite its master graph.
                if (placeholder.Source.DirectFramePresenceEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type))
                    capabilities.Add(new("setFrame", EditableFrameFields));
                // Text is exposed only when the projection can preserve its
                // paragraph/run topology.  A string fallback for a formula,
                // hyperlink, or opaque run graph must remain source-owned.
                if (placeholder.Source.TextEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type) &&
                    PlaceholderTextProjectionEditable(placeholder.TextBody))
                    capabilities.Add(new("replaceText", ["text"]));
                if (placeholder.Source.TextEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type) &&
                    PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(placeholder.TextBody?.BodyProperties))
                    capabilities.Add(new("setTextBodyStyle", ["text.style"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"placeholder:{placeholder.Id}",
                    HashOrFallback(placeholder.Source.ElementSha256, placeholder),
                    capabilities);
            }
            if (placeholder.TextBody is not null)
            {
                projected["text"] = TextContent(placeholder.TextBody, PptxTextCodec.Flatten(placeholder.TextBody), context);
                if (TextBoxStyle(placeholder.TextBody) is { Count: > 0 } style)
                    projected["style"] = style;
            }
            output.Add(projected);
        }
        return output;
    }

    private static bool PlaceholderTextProjectionEditable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase is
                PresentationTextRun.ContentOneofCase.Text or
                PresentationTextRun.ContentOneofCase.LineBreak or
                PresentationTextRun.ContentOneofCase.Field));

    private static bool BoundedLayoutPlaceholderType(string value) => value is
        "title" or "body" or "ctrTitle" or "subtitle" or "subTitle" or
        "content" or "obj" or "picture" or "pic" or "chart" or
        "table" or "tbl" or "date" or "dt" or "footer" or "ftr" or
        "slide-number" or "sldNum";

    private static void AddMasterTextLevels(
        JsonObject target,
        string name,
        IEnumerable<PresentationTextParagraph>? levels,
        ProjectionContext context)
    {
        if (levels is null) return;
        var projected = new JsonArray();
        foreach (var level in levels)
        {
            var item = new JsonObject();
            if (level.HasLevel) item["level"] = JsonValue.Create(checked((int)level.Level));
            if (level.HasAlignment && ParagraphAlignment(level.Alignment) is { } alignment)
                item["alignment"] = StringNode(alignment);
            if (level.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                item["indent"] = JsonValue.Create(Points(level.MarginLeftEmu));
            if (level.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                item["hanging"] = JsonValue.Create(-Points(level.IndentEmu));
            if (level.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                item["lineSpacing"] = JsonValue.Create(Math.Max(0.001, level.LineSpacingPoints));
            if (level.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                item["lineSpacingMultiplier"] = JsonValue.Create(Math.Max(0.00001, level.LineSpacingMultiplier));
            if (level.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                item["spaceBefore"] = JsonValue.Create(Math.Max(0, level.SpaceBeforePoints));
            if (level.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                item["spaceBeforeMultiplier"] = JsonValue.Create(Math.Max(0, level.SpaceBeforeMultiplier));
            if (level.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                item["spaceAfter"] = JsonValue.Create(Math.Max(0, level.SpaceAfterPoints));
            if (level.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                item["spaceAfterMultiplier"] = JsonValue.Create(Math.Max(0, level.SpaceAfterMultiplier));
            if (level.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                ProjectTextStyle(level.DefaultRunProperties) is { Count: > 0 } defaultText)
                item["defaultText"] = defaultText;
            if (ProjectBullet(level, context) is { } bullet) item["bullet"] = bullet;
            projected.Add(item);
        }
        if (projected.Count > 0) target[name] = projected;
    }

    private static JsonObject FrameDimensions(PresentationArtifact presentation) => new()
    {
        ["width"] = JsonValue.Create(Points(presentation.SlideWidthEmu)),
        ["height"] = JsonValue.Create(Points(presentation.SlideHeightEmu)),
        ["unit"] = StringNode("pt"),
    };

    private static JsonObject ProjectCanvas(PresentationArtifact presentation, ProjectionContext context)
    {
        var canvas = FrameDimensions(presentation);
        var objectHash = Sha256(CanonicalBytes(canvas));
        canvas["nativeRef"] = NativeRef(
            context,
            "canvas",
            objectHash,
            [new("setCanvas", ["canvas.width", "canvas.height"])]);
        return canvas;
    }

    private static void RegisterIds(PresentationArtifact presentation, ProjectionContext context)
    {
        foreach (var master in presentation.Masters) context.RegisterMaster(master.Id);
        foreach (var layout in presentation.Layouts) context.RegisterLayout(layout.Id);
        foreach (var customShow in presentation.CustomShows) context.RegisterCustomShow(customShow.Id);
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.RegisterPage(slide.Id, slide.Source?.PartPath);
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
        var pageHash = HashOrFallback(slide.Source?.SlideXmlSha256, slide);
        var pageCapabilities = new List<CapabilitySpec>();
        if (slide.Source is not null)
            pageCapabilities.Add(new("setName", ["name"]));
        if (slide.Source?.VisibilityEditable == true)
            pageCapabilities.Add(new("setHidden", ["hidden"]));
        if (slide.Source?.DeletionCapability?.Supported == true)
            pageCapabilities.Add(new("delete", ["element"]));
        if (slide.Source?.BackgroundEditable == true)
            pageCapabilities.Add(new("setBackground", ["background"]));
        if (slide.Source?.TransitionEditable == true || slide.Source?.TransitionAddable == true)
            pageCapabilities.Add(new("setTransition", ["transition"]));
        if (slide.Morph is null &&
            (slide.Source?.TimingEditable == true || slide.Source?.TimingAddable == true))
            pageCapabilities.Add(new("setAnimations", ["animations"]));
        if (slide.SpeakerNotes?.Source?.Editable == true || slide.Source?.SpeakerNotesAddable == true)
            pageCapabilities.Add(new("setNotes", ["notes"]));
        if (slide.Source is not null && presentation.Slides.Count > 1 && !presentation.SectionsOpaque)
            pageCapabilities.Add(new("reorder", ["pageOrder"]));
        if (slide.Source?.CloneCapability?.Supported == true)
            pageCapabilities.Add(new("duplicate", ["pageClone"]));
        if (slide.Source is not null)
            pageCapabilities.Add(new("appendElement", ["elements"]));

        var elements = new JsonArray();
        for (var elementIndex = 0; elementIndex < slide.Elements.Count; elementIndex++)
        {
            var element = slide.Elements[elementIndex];
            elements.Add(ProjectElement(
                element,
                slide,
                pageId,
                context,
                [element.Source?.ShapeTreeIndex ?? checked((uint)elementIndex)],
                element.ContentCase == PresentationElement.ContentOneofCase.Shape
                    ? EffectiveSlidePlaceholderFrame(element.Shape, slide, presentation)
                    : null));
        }

        var page = new JsonObject
        {
            ["id"] = StringNode(pageId),
            ["role"] = StringNode("source continuation"),
            ["elements"] = elements,
            ["nativeRef"] = NativeRef(context, $"page:{pageId}", pageHash, pageCapabilities),
        };
        if (slide.ReadingOrder.Count > 0)
        {
            var readingOrder = new JsonArray();
            foreach (var elementId in slide.ReadingOrder)
                if (context.TryElementId(pageId, elementId, out var projectedId))
                    readingOrder.Add(StringNode(projectedId));
            if (readingOrder.Count == slide.Elements.Count)
                page["readingOrder"] = readingOrder;
        }
        if (!string.IsNullOrWhiteSpace(slide.Name)) page["name"] = StringNode(slide.Name);
        if (!string.IsNullOrWhiteSpace(slide.LayoutId) && context.TryLayoutId(slide.LayoutId, out var layoutId))
            page["layout"] = StringNode(layoutId);
        if (slide.HasHidden) page["hidden"] = JsonValue.Create(slide.Hidden);
        if (!string.IsNullOrEmpty(slide.SpeakerNotes?.Text))
            page["notes"] = TextContent(slide.SpeakerNotes.TextBody, slide.SpeakerNotes.Text, context);
        if (ProjectBackground(slide.Background, context) is { } background) page["background"] = background;

        var animations = ProjectAnimations(slide, context);
        if (animations.Count > 0) page["animations"] = animations;
        if (ProjectTransition(slide, presentation, context) is { } transition) page["transition"] = transition;
        return page;
    }

    // Slide placeholders without a local a:xfrm render through the linked
    // layout, which in turn may inherit a matching master placeholder.  Keep
    // that effective frame in the PPJ view so the required public frame is
    // useful for inspection and editing.  The returned frame is only a
    // projection; the slide owner remains the source-bound edit boundary and
    // materialization is gated by the placeholder capability below.
    private static PresentationPlaceholderFrame? EffectiveSlidePlaceholderFrame(
        PresentationShape shape,
        PresentationSlide slide,
        PresentationArtifact presentation)
    {
        if (shape.Placeholder is null) return null;
        if (shape.DirectFrame is not null) return shape.DirectFrame.Clone();
        if (!shape.Placeholder.InheritsGeometry || string.IsNullOrWhiteSpace(slide.LayoutId)) return null;

        var layout = presentation.Layouts.FirstOrDefault(candidate =>
            candidate.Id.Equals(slide.LayoutId, StringComparison.Ordinal));
        if (layout is null) return null;
        var keyType = shape.Placeholder.Type;
        var keyIndex = shape.Placeholder.Index;
        var layoutMatches = layout.Placeholders
            .Where(candidate => candidate.Type.Equals(keyType, StringComparison.Ordinal) && candidate.Index == keyIndex)
            .ToArray();
        if (layoutMatches.Length > 1) return null;
        if (layoutMatches.Length == 1 && layoutMatches[0].DirectFrame is { } layoutFrame)
            return layoutFrame.Clone();

        var master = presentation.Masters.FirstOrDefault(candidate =>
            candidate.Id.Equals(layout.MasterId, StringComparison.Ordinal));
        if (master is null) return null;
        var masterMatches = master.Placeholders
            .Where(candidate => candidate.Type.Equals(keyType, StringComparison.Ordinal) && candidate.Index == keyIndex)
            .ToArray();
        return masterMatches.Length == 1 && masterMatches[0].DirectFrame is { } masterFrame
            ? masterFrame.Clone()
            : null;
    }

    private static JsonObject ProjectElement(
        PresentationElement element,
        PresentationSlide slide,
        string pageId,
        ProjectionContext context,
        IReadOnlyList<uint> shapeTreePath,
        PresentationPlaceholderFrame? effectivePlaceholderFrame = null)
    {
        var id = context.ElementId(pageId, element.Id);
        var hash = HashOrFallback(element.Source?.ElementSha256, element);
        var capabilities = Capabilities(element, effectivePlaceholderFrame is not null);
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
            PresentationElement.ContentOneofCase.Shape => ProjectShape(element, id, nativeRef, context, effectivePlaceholderFrame),
            PresentationElement.ContentOneofCase.Image => ProjectImage(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Table => ProjectTable(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Connector => ProjectConnector(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Chart => ProjectChart(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Diagram => ProjectNativeSmartArt(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Group => ProjectGroup(element, id, nativeRef, slide, pageId, context, shapeTreePath),
            PresentationElement.ContentOneofCase.Opaque when element.Opaque.DiagramText is not null =>
                ProjectSourceSmartArt(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Opaque when element.Opaque.OleWorkbook is not null || element.Opaque.OleOfficePackage is not null =>
                ProjectSourceOle(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Opaque => ProjectOpaque(element, id, nativeRef),
            _ => ProjectOpaque(element, id, nativeRef, "unknown"),
        };
        if (element.HasHidden && element.Hidden) projected["hidden"] = JsonValue.Create(true);
        if (element.HasLocked && element.Locked) projected["locked"] = JsonValue.Create(true);
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
        ProjectionContext context,
        PresentationPlaceholderFrame? effectivePlaceholderFrame = null)
    {
        var shape = element.Shape;
        var isPlaceholder = shape.Placeholder is not null;
        if (isPlaceholder && shape.DirectFrame is null && effectivePlaceholderFrame is null)
            return ProjectOpaque(element, id, nativeRef, "placeholder", "Preserved slide placeholder whose effective layout/master frame is ambiguous or unavailable.");
        var frame = isPlaceholder && (effectivePlaceholderFrame ?? shape.DirectFrame) is { } placeholderFrame
            ? ShapeFrame(placeholderFrame)
            : ShapeFrame(shape);
        var text = TextContent(shape.TextBody, shape.Text, context);
        var hasText = !string.IsNullOrEmpty(shape.Text) || shape.TextBody?.Paragraphs.Count > 0;
        var isTextBox = shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry);
        var lineLike = shape.Geometry == "line" || PpjLinePathCodec.IsLineLike(shape);
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
            !CanProjectCustomGeometry(shape) && !sourceImageFill && !sourceCustomGeometry && !lineLike)
            return ProjectOpaque(element, id, nativeRef, "shape", $"Preserved source shape with unsupported geometry '{shape.Geometry}'.");

        var common = ElementBase(id, element.Name, frame, Accessibility(shape.Accessibility), nativeRef);
        if (ProjectAction(shape.Action, context) is { } action)
            common["action"] = action;
        if (ProjectAction(shape.HoverAction, context) is { } hoverAction)
            common["hoverAction"] = hoverAction;
        if (shape.Placeholder is not null)
        {
            common["type"] = StringNode("placeholder");
            common["placeholderType"] = StringNode(PlaceholderType(shape.Placeholder.Type));
            common["index"] = JsonValue.Create(shape.Placeholder.Index);
            if (hasText) common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } placeholderStyle) common["style"] = placeholderStyle;
            return common;
        }
        if (shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry))
        {
            common["type"] = StringNode("text");
            common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } textStyle) common["style"] = textStyle;
            ApplyTextContainerStyle(common, shape, context);
            return common;
        }
        if (lineLike)
        {
            common["type"] = StringNode("line");
            if (shape.Geometry == "line")
                common["path"] = PpjLinePathCodec.Synthetic(shape);
            else if (PpjLinePathCodec.TryProjectKimi(shape) is { } kimiPath)
            {
                common["viewBox"] = kimiPath["viewBox"]!.DeepClone();
                common["points"] = kimiPath["points"]!.DeepClone();
                common["curve"] = kimiPath["curve"]!.DeepClone();
            }
            else
                common["path"] = PpjLinePathCodec.Project(shape);
            var lineStyle = ShapeStyle(shape, context);
            if (lineStyle.TryGetPropertyValue("stroke", out var lineStroke) && lineStroke is not null)
                common["stroke"] = lineStroke.DeepClone();
            else
                common["stroke"] = Stroke("000000", shape.LineWidthEmu, "solid", shape.LineCap, shape.LineJoin,
                    shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                    shape.LineScheme);
            if (Arrow(shape.StartArrow) is { } startArrow) common["startArrow"] = startArrow;
            if (Arrow(shape.EndArrow) is { } endArrow) common["endArrow"] = endArrow;
            if (shape.Shadow is not null) common["shadow"] = Shadow(shape.Shadow);
            if (shape.Glow is not null) common["glow"] = Glow(shape.Glow);
            if (shape.InnerShadow is not null) common["innerShadow"] = InnerShadow(shape.InnerShadow);
            if (shape.Reflection is not null) common["reflection"] = Reflection(shape.Reflection);
            if (shape.SoftEdge is not null) common["softEdge"] = SoftEdge(shape.SoftEdge);
            return common;
        }
        common["type"] = StringNode("shape");
        var compoundOpacity = 1d;
        var hasCompoundOpacity = element.Source?.Editable == true &&
            TryGetCompoundShapeOpacity(shape, out compoundOpacity);
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
                    ["kind"] = StringNode("source-custom"),
                    ["sourceBound"] = JsonValue.Create(true),
                };
        else
        {
            var geometry = new JsonObject { ["kind"] = StringNode("preset"), ["preset"] = StringNode(shape.Geometry) };
            if (PptxPresetGeometryAdjustmentCodec.IsCompleteValues(shape.Geometry, shape.PresetAdjustments))
                geometry["adjustments"] = new JsonArray(shape.PresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            common["geometry"] = geometry;
        }
        if (hasText) common["text"] = text;
        if (TextBoxStyle(shape.TextBody) is { Count: > 0 } shapeTextStyle) common["textStyle"] = shapeTextStyle;
        var style = ShapeStyle(shape, context, omitOwnerOpacity: hasCompoundOpacity);
        if (style.Count > 0) common["style"] = style;
        if (hasCompoundOpacity)
            common["compositing"] = new JsonObject { ["opacity"] = JsonValue.Create(compoundOpacity) };
        return common;
    }

    internal static bool TryGetCompoundShapeOpacity(PresentationShape shape, out double opacity)
    {
        opacity = 1;
        if (shape.Placeholder is not null || !string.IsNullOrEmpty(shape.Text) ||
            shape.Geometry is "line" or "custom" or "textbox" or "none" ||
            string.IsNullOrEmpty(shape.Geometry))
            return false;
        if (!string.IsNullOrEmpty(shape.FillRgb) && !string.IsNullOrEmpty(shape.FillScheme))
            return false;
        if (!string.IsNullOrWhiteSpace(shape.ImageFillAssetId) && shape.ImageFill is null)
            return false;

        var values = new List<double>();
        if (shape.FillRgb.Length > 0 || shape.FillScheme.Length > 0)
            values.Add(shape.HasFillOpacityThousandthPercent ? Unit(shape.FillOpacityThousandthPercent) : 1);
        if (shape.GradientFill is not null)
        {
            if (shape.GradientFill.Stops.Count == 0) return false;
            values.AddRange(shape.GradientFill.Stops.Select(stop =>
                stop.HasOpacityThousandthPercent ? Unit(stop.OpacityThousandthPercent) : 1));
        }
        if (shape.ImageFill is not null)
            values.Add(shape.ImageFill.HasOpacityThousandthPercent ? Unit(shape.ImageFill.OpacityThousandthPercent) : 1);
        if (shape.LineStyle != "none" && (shape.LineRgb.Length > 0 || shape.LineScheme.Length > 0))
            values.Add(shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : 1);
        if (shape.Shadow is not null)
            values.Add(shape.Shadow.HasOpacityThousandthPercent ? Unit(shape.Shadow.OpacityThousandthPercent) : 1);
        if (values.Count == 0) return false;
        var candidate = values[0];
        opacity = candidate;
        return values.All(value => Math.Abs(value - candidate) < 0.000005);
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
            if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.Normal) path["fill"] = JsonValue.Create(true);
            else if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.None) path["fill"] = JsonValue.Create(false);
            if (source.HasStroke) path["stroke"] = JsonValue.Create(source.Stroke);
            paths.Add(path);
        }
        return new JsonObject
        {
            ["kind"] = StringNode("custom"),
            ["viewBox"] = new JsonObject
            {
                ["x"] = JsonValue.Create(0),
                ["y"] = JsonValue.Create(0),
                ["width"] = JsonValue.Create(width),
                ["height"] = JsonValue.Create(height),
            },
            ["paths"] = paths,
        };
    }

    private static JsonObject ProjectCustomCommand(PresentationCustomGeometryCommand command) => command.CommandCase switch
    {
        PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => ProjectCustomPoint("moveTo", command.MoveTo),
        PresentationCustomGeometryCommand.CommandOneofCase.LineTo => ProjectCustomPoint("lineTo", command.LineTo),
        PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo => new JsonObject
        {
            ["op"] = StringNode("quadraticTo"),
            ["x1"] = JsonValue.Create(CustomPathPoint(command.QuadraticBezierTo.Control.X)),
            ["y1"] = JsonValue.Create(CustomPathPoint(command.QuadraticBezierTo.Control.Y)),
            ["x"] = JsonValue.Create(CustomPathPoint(command.QuadraticBezierTo.End.X)),
            ["y"] = JsonValue.Create(CustomPathPoint(command.QuadraticBezierTo.End.Y)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new JsonObject
        {
            ["op"] = StringNode("cubicTo"),
            ["x1"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.Control1.X)),
            ["y1"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.Control1.Y)),
            ["x2"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.Control2.X)),
            ["y2"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.Control2.Y)),
            ["x"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.End.X)),
            ["y"] = JsonValue.Create(CustomPathPoint(command.CubicBezierTo.End.Y)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => new JsonObject
        {
            ["op"] = StringNode("arcTo"),
            ["radiusX"] = JsonValue.Create(CustomPathPoint(command.ArcTo.WidthRadius)),
            ["radiusY"] = JsonValue.Create(CustomPathPoint(command.ArcTo.HeightRadius)),
            ["startAngle"] = JsonValue.Create(NormalizeCustomPathStartAngle(command.ArcTo.StartAngle)),
            ["sweepAngle"] = JsonValue.Create(CustomPathAngle(command.ArcTo.SweepAngle)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.Close => new JsonObject { ["op"] = StringNode("close") },
        _ => throw new InvalidOperationException("Unsupported PPJ custom path command passed the projection gate."),
    };

    private static JsonObject ProjectCustomPoint(string operation, PresentationCustomGeometryPoint point) => new()
    {
        ["op"] = StringNode(operation),
        ["x"] = JsonValue.Create(CustomPathPoint(point.X)),
        ["y"] = JsonValue.Create(CustomPathPoint(point.Y)),
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
        string? svgAssetId = null;
        if (!string.IsNullOrEmpty(image.SvgAssetId) && !context.TryMaterializeAsset(image.SvgAssetId, out svgAssetId))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose paired SVG payload cannot be materialized safely.");
        var customMask = image.CustomMaskPaths.Count > 0 ? ImageMaskShape(image) : null;
        if (customMask is not null && !CanProjectCustomGeometry(customMask))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose custom mask cannot be represented exactly in PPJ.");
        var output = ElementBase(id, element.Name, ImageFrame(image), ImageAccessibility(image), nativeRef);
        output["type"] = StringNode("image");
        output["asset"] = StringNode(assetId);
        if (svgAssetId is not null) output["svgAsset"] = StringNode(svgAssetId);
        output["fit"] = StringNode(image.Tiled ? "tile" : "stretch");
        if (image.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(image.Crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(image.Crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(image.Crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(image.Crop.BottomThousandthPercent)),
            };
        }
        if (image.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(image.OpacityThousandthPercent));
        if (customMask is not null)
            output["mask"] = ProjectCustomGeometry(customMask);
        else if (!string.IsNullOrEmpty(image.MaskPreset) && PptxPresetGeometryAdjustmentCodec.HasProfile(image.MaskPreset))
        {
            var mask = new JsonObject { ["kind"] = StringNode("preset"), ["preset"] = StringNode(image.MaskPreset) };
            if (image.MaskPresetAdjustments.Count > 0)
                mask["adjustments"] = new JsonArray(image.MaskPresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            output["mask"] = mask;
        }
        if (image.Border is not null &&
            (!string.IsNullOrEmpty(image.Border.ColorRgb) || !string.IsNullOrEmpty(image.Border.ColorScheme)))
            output["border"] = Stroke(image.Border.ColorRgb, image.Border.WidthEmu, image.Border.Style, image.Border.Cap, image.Border.Join,
                image.Border.HasOpacityThousandthPercent ? Unit(image.Border.OpacityThousandthPercent) : null,
                image.Border.ColorScheme);
        if (image.Shadow is not null && (!string.IsNullOrEmpty(image.Shadow.ColorRgb) || !string.IsNullOrEmpty(image.Shadow.ColorScheme)))
            output["shadow"] = Shadow(image.Shadow);
        if (image.Glow is not null && (!string.IsNullOrEmpty(image.Glow.ColorRgb) || !string.IsNullOrEmpty(image.Glow.ColorScheme)))
            output["glow"] = Glow(image.Glow);
        if (image.InnerShadow is not null && (!string.IsNullOrEmpty(image.InnerShadow.ColorRgb) || !string.IsNullOrEmpty(image.InnerShadow.ColorScheme)))
            output["innerShadow"] = InnerShadow(image.InnerShadow);
        if (image.Reflection is not null)
            output["reflection"] = Reflection(image.Reflection);
        if (image.SoftEdge is not null)
            output["softEdge"] = SoftEdge(image.SoftEdge);
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

    private static JsonObject ProjectChart(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
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
        output["type"] = StringNode("chart");
        output["chartType"] = StringNode(chart.Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar" ? "bar" : type!);
        if (!string.IsNullOrEmpty(chart.Title))
            output["title"] = chart.TitleBody is null
                ? StringNode(chart.Title)
                : TextContent(chart.TitleBody, chart.Title, context);
        if (chart.HasDisplayBlanksAs) output["displayBlanksAs"] = StringNode(chart.DisplayBlanksAs);
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
                ["id"] = StringNode($"series-{index + 1}"),
                ["name"] = StringNode(item.Name ?? string.Empty),
                ["values"] = values,
            };
            if (!string.IsNullOrWhiteSpace(item.CategoryFormula)) entry["categoryFormula"] = StringNode(item.CategoryFormula);
            if (!string.IsNullOrWhiteSpace(item.XValueFormula)) entry["xValueFormula"] = StringNode(item.XValueFormula);
            if (!string.IsNullOrWhiteSpace(item.ValueFormula)) entry["valueFormula"] = StringNode(item.ValueFormula);
            if (!string.IsNullOrWhiteSpace(item.BubbleSizeFormula)) entry["bubbleSizeFormula"] = StringNode(item.BubbleSizeFormula);
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
                entry["chartType"] = StringNode(chart.ComboSeries[index].Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar"
                    ? "bar"
                    : ChartType(chart.ComboSeries[index].Type) ?? "line");
                var axisIndex = chart.ComboSeries[index].AxisGroup == PresentationChartAxisGroup.Secondary ? 1 : 0;
                entry["axis"] = StringNode(axisIndex == 1 ? "secondary" : "primary");
                // The native chart model has two axis groups rather than the
                // Kimi array form. Emit both bounded indexes so a projected
                // combo series can be fed back into the dataset/encode
                // compatibility layer without losing its 0/1 placement.
                entry["xAxisIndex"] = JsonValue.Create(axisIndex);
                entry["yAxisIndex"] = JsonValue.Create(axisIndex);
            }
            ProjectChartSeriesStyle(entry, item);
            seriesJson.Add(entry);
        }
        var data = new JsonObject { ["categories"] = categories, ["series"] = seriesJson };
        if (CanProjectCanonicalDataset(chart, series))
        {
            // Keep the legacy categories/series projection for wire
            // compatibility, and add a deterministic dataset/encoding view
            // only when it is a lossless view of the same native vectors.
            // Complex formulas, sparse point overrides, and per-series
            // topology stay on the legacy path so a re-import cannot silently
            // discard their semantics.
            data["dataset"] = CanonicalDataset(chart, series);
            var encoding = chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble
                ? new JsonObject
                {
                    ["x"] = StringNode("x"),
                    ["y"] = StringNode("y"),
                    ["series"] = StringNode("series"),
                }
                : new JsonObject
                {
                    ["category"] = StringNode("category"),
                    ["series"] = StringNode("series"),
                    ["value"] = StringNode("value"),
                };
            if (chart.Type == SpreadsheetChartType.Bubble) encoding["size"] = StringNode("size");
            data["encoding"] = encoding;
        }
        output["data"] = data;
        if (chart.Type == SpreadsheetChartType.Radar && TryProjectRadarSpokeAxis(chart, out var spokeAxis))
            output["spokeAxis"] = spokeAxis;
        else
        {
            if (chart.XAxis is not null) output["xAxis"] = ProjectChartAxis(chart.XAxis);
            if (chart.YAxis is not null) output["yAxis"] = ProjectChartAxis(chart.YAxis);
        }
        if (chart.SecondaryXAxis is not null) output["secondaryXAxis"] = ProjectChartAxis(chart.SecondaryXAxis);
        if (chart.SecondaryYAxis is not null) output["secondaryYAxis"] = ProjectChartAxis(chart.SecondaryYAxis);
        var style = new JsonObject
        {
            ["legend"] = StringNode(chart.HasLegend
                ? chart.LegendPosition.Length == 0 ? "right" : chart.LegendPosition
                : "none"),
        };
        if (chart.HasLegendOverlay) style["legendOverlay"] = JsonValue.Create(chart.LegendOverlay);
        if (chart.LegendFill is not null) style["legendFill"] = ProjectChartSurfaceFill(chart.LegendFill);
        if (chart.LegendLine is not null && !string.IsNullOrEmpty(chart.LegendLine.Color?.Rgb))
            style["legendLine"] = ProjectChartLine(chart.LegendLine);
        if (chart.Grouping.Length > 0) style["stacking"] = StringNode(chart.Grouping);
        if (chart.HasGapWidth) style["gapWidth"] = JsonValue.Create(chart.GapWidth);
        if (chart.HasOverlap) style["overlap"] = JsonValue.Create(chart.Overlap);
        if (chart.HasVaryColors) style["varyColors"] = JsonValue.Create(chart.VaryColors);
        if (chart.HasFirstSliceAngle) style["startAngle"] = JsonValue.Create(chart.FirstSliceAngle);
        if (chart.HasDoughnutHoleSize) style["holeSize"] = JsonValue.Create(chart.DoughnutHoleSize);
        if (chart.HasBubbleScale) style["bubbleScale"] = JsonValue.Create(chart.BubbleScale);
        if (chart.BubbleSizeMode.Length > 0) style["bubbleSizeMode"] = StringNode(chart.BubbleSizeMode);
        if (chart.XAxis is null && chart.HasShowCategoryAxis) style["showCategoryAxis"] = JsonValue.Create(chart.ShowCategoryAxis);
        if (chart.YAxis is null && chart.HasShowValueAxis) style["showValueAxis"] = JsonValue.Create(chart.ShowValueAxis);
        if (chart.YAxis is null && chart.HasShowGridlines) style["showGridlines"] = JsonValue.Create(chart.ShowGridlines);
        if (chart.ChartAreaFill is not null) style["chartAreaFill"] = ProjectChartSurfaceFill(chart.ChartAreaFill);
        if (chart.PlotAreaFill is not null) style["plotAreaFill"] = ProjectChartSurfaceFill(chart.PlotAreaFill);
        if (chart.Frame is not null)
        {
            var frame = new JsonObject();
            if (chart.Frame.ImageFill is not null)
            {
                if (ProjectImagePaint(chart.Frame.ImageFill, context) is { } imageFill)
                    frame["fill"] = imageFill;
            }
            else if (chart.Frame.Fill is not null) frame["fill"] = ProjectChartSurfaceFill(chart.Frame.Fill);
            if (chart.Frame.Line is not null) frame["stroke"] = ProjectChartLine(chart.Frame.Line);
            if (chart.Frame.Shadow is not null) frame["shadow"] = Shadow(chart.Frame.Shadow);
            if (frame.Count > 0) style["frame"] = frame;
        }
        if (chart.TitleTextStyle is not null)
            style["titleTextStyle"] = ProjectChartTextStyle(chart.TitleTextStyle);
        if (chart.LegendTextStyle is not null)
            style["legendTextStyle"] = ProjectChartTextStyle(chart.LegendTextStyle);
        if (chart.LineOptions?.HasSmooth == true) style["smooth"] = JsonValue.Create(chart.LineOptions.Smooth);
        if (chart.LineOptions?.VaryColors == true) style["varyColors"] = JsonValue.Create(true);
        if (chart.DataLabels is not null)
        {
            var labels = new JsonObject
            {
                ["showValue"] = JsonValue.Create(chart.DataLabels.ShowValue),
                ["showCategory"] = JsonValue.Create(chart.DataLabels.ShowCategoryName),
            };
            if (chart.DataLabels.HasShowSeriesName) labels["showSeries"] = JsonValue.Create(chart.DataLabels.ShowSeriesName);
            if (chart.DataLabels.HasShowPercent) labels["showPercent"] = JsonValue.Create(chart.DataLabels.ShowPercent);
            if (chart.DataLabels.HasShowBubbleSize) labels["showBubbleSize"] = JsonValue.Create(chart.DataLabels.ShowBubbleSize);
            if (chart.DataLabels.HasShowLeaderLines) labels["showLeaderLines"] = JsonValue.Create(chart.DataLabels.ShowLeaderLines);
            if (chart.DataLabels.HasPosition && DataLabelPosition(chart.DataLabels.Position) is { } position)
                labels["position"] = StringNode(position);
            if (chart.DataLabels.TextStyle is not null)
                labels["textStyle"] = ProjectChartTextStyle(chart.DataLabels.TextStyle);
            if (chart.DataLabels.NumberFormatCode.Length > 0)
                labels["numberFormat"] = StringNode(chart.DataLabels.NumberFormatCode);
            if (chart.DataLabels.Fill is not null)
                labels["fill"] = ProjectChartSurfaceFill(chart.DataLabels.Fill);
            if (chart.DataLabels.Line is not null && !string.IsNullOrEmpty(chart.DataLabels.Line.Color?.Rgb))
                labels["line"] = ProjectChartLine(chart.DataLabels.Line);
            style["dataLabels"] = labels;
        }
        output["style"] = style;
        return output;
    }

    private static bool CanProjectCanonicalDataset(
        PresentationChart chart,
        IReadOnlyList<SpreadsheetChartSeriesArtifact> series)
    {
        if (series.Count == 0) return false;
        if (chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble)
        {
            if (chart.Categories.Count != 0) return false;
            return series.All(item =>
                item.Values.Count > 0 &&
                item.Values.Count == item.XValues.Count &&
                (chart.Type != SpreadsheetChartType.Bubble || item.Values.Count == item.BubbleSizes.Count) &&
                (chart.Type != SpreadsheetChartType.Scatter || item.BubbleSizes.Count == 0) &&
                string.IsNullOrWhiteSpace(item.CategoryFormula) &&
                string.IsNullOrWhiteSpace(item.XValueFormula) &&
                string.IsNullOrWhiteSpace(item.ValueFormula) &&
                string.IsNullOrWhiteSpace(item.BubbleSizeFormula) &&
                item.MissingValueIndexes.Count == 0 &&
                item.Trendlines.Count == 0 && item.ErrorBars is null &&
                item.Fill is null && item.Line is null && item.Marker is null &&
                item.SeriesFill is null && item.DataLabels is null && item.PointStyles.Count == 0);
        }
        if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area or
            SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut or SpreadsheetChartType.Radar) ||
            chart.Categories.Count == 0)
            return false;
        return series.All(item =>
            item.Values.Count == chart.Categories.Count &&
            string.IsNullOrWhiteSpace(item.CategoryFormula) &&
            string.IsNullOrWhiteSpace(item.XValueFormula) &&
            string.IsNullOrWhiteSpace(item.ValueFormula) &&
            string.IsNullOrWhiteSpace(item.BubbleSizeFormula) &&
            item.XValues.Count == 0 && item.BubbleSizes.Count == 0 &&
            item.MissingValueIndexes.Count == 0 &&
            item.Trendlines.Count == 0 && item.ErrorBars is null &&
            item.Fill is null && item.Line is null && item.Marker is null &&
            item.SeriesFill is null && item.DataLabels is null && item.PointStyles.Count == 0);
    }

    private static JsonObject CanonicalDataset(
        PresentationChart chart,
        IReadOnlyList<SpreadsheetChartSeriesArtifact> series)
    {
        var rows = new JsonArray();
        if (chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble)
        {
            foreach (var item in series)
            {
                for (var pointIndex = 0; pointIndex < item.Values.Count; pointIndex++)
                {
                    var row = new JsonArray
                    {
                        NumberNode(item.XValues[pointIndex]),
                        StringNode(item.Name ?? string.Empty),
                        NumberNode(item.Values[pointIndex]),
                    };
                    if (chart.Type == SpreadsheetChartType.Bubble)
                        row.Add(NumberNode(item.BubbleSizes[pointIndex]));
                    rows.Add(row);
                }
            }
            return new JsonObject
            {
                ["cols"] = chart.Type == SpreadsheetChartType.Bubble
                    ? new JsonArray { StringNode("x"), StringNode("series"), StringNode("y"), StringNode("size") }
                    : new JsonArray { StringNode("x"), StringNode("series"), StringNode("y") },
                ["rows"] = rows,
            };
        }
        for (var categoryIndex = 0; categoryIndex < chart.Categories.Count; categoryIndex++)
        {
            foreach (var item in series)
            {
                var row = new JsonObject
                {
                    ["category"] = StringNode(chart.Categories[categoryIndex]),
                    ["series"] = StringNode(item.Name ?? string.Empty),
                };
                var missing = item.MissingValueIndexes.Contains((uint)categoryIndex);
                row["value"] = missing ? null : NumberNode(item.Values[categoryIndex]);
                rows.Add(row);
            }
        }
        return new JsonObject
        {
            ["cols"] = new JsonArray { StringNode("category"), StringNode("series"), StringNode("value") },
            ["rows"] = rows,
        };
    }

    private static void ProjectChartSeriesStyle(JsonObject output, SpreadsheetChartSeriesArtifact series)
    {
        if (series.SeriesFill is not null)
            output["fill"] = ProjectChartSurfaceFill(series.SeriesFill);
        else if (series.Fill is not null && !string.IsNullOrEmpty(series.Fill.Rgb))
            output["fill"] = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(series.Fill.Rgb)) };
        if (series.Line is not null && !string.IsNullOrEmpty(series.Line.Color?.Rgb))
            output["stroke"] = ProjectChartLine(series.Line);
        if (series.PointStyles.Count > 0)
        {
            var points = new JsonArray();
            foreach (var point in series.PointStyles.OrderBy(point => point.Index))
            {
            var item = new JsonObject { ["index"] = JsonValue.Create(point.Index) };
                if (point.Fill is not null) item["fill"] = ProjectChartSurfaceFill(point.Fill);
                if (point.Line is not null && !string.IsNullOrEmpty(point.Line.Color?.Rgb))
                    item["stroke"] = ProjectChartLine(point.Line);
                if (point.HasExplosion) item["explosion"] = JsonValue.Create(point.Explosion);
                points.Add(item);
            }
            output["pointStyles"] = points;
        }
        if (series.Marker is not null && series.Marker.Symbol != SpreadsheetChartMarkerSymbol.Unspecified)
        {
            var symbol = Marker(series.Marker.Symbol);
            if (symbol is not null)
            {
                if (!series.Marker.HasSize && series.Marker.Fill is null && series.Marker.Line is null)
                    output["marker"] = StringNode(symbol);
                else
                {
                    var marker = new JsonObject { ["symbol"] = StringNode(symbol) };
                    if (series.Marker.HasSize) marker["size"] = JsonValue.Create(series.Marker.Size);
                    if (series.Marker.Fill is not null && !string.IsNullOrEmpty(series.Marker.Fill.Rgb))
                    {
                        var color = Color(series.Marker.Fill.Rgb);
                        if (series.Marker.HasFillOpacityThousandthPercent)
                        {
                            var alpha = Math.Clamp((int)Math.Round(Unit(series.Marker.FillOpacityThousandthPercent) * 255), 0, 255);
                            color += $"{alpha:X2}";
                        }
                        marker["fill"] = StringNode(color);
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
                var trendline = new JsonObject { ["type"] = StringNode(type) };
                if (!string.IsNullOrEmpty(item.Name)) trendline["name"] = StringNode(item.Name);
                if (item.HasPolynomialOrder) trendline["order"] = JsonValue.Create(item.PolynomialOrder);
                if (item.HasPeriod) trendline["period"] = JsonValue.Create(item.Period);
                if (item.HasForward) trendline["forward"] = JsonValue.Create(item.Forward);
                if (item.HasBackward) trendline["backward"] = JsonValue.Create(item.Backward);
                if (item.HasIntercept) trendline["intercept"] = JsonValue.Create(item.Intercept);
                if (item.DisplayEquation) trendline["displayEquation"] = JsonValue.Create(true);
                if (item.DisplayRSquared) trendline["displayRSquared"] = JsonValue.Create(true);
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
                ["direction"] = StringNode(direction),
                ["type"] = StringNode(barType),
                ["valueType"] = StringNode(valueType),
            };
            if (errorBars.HasValue) projected["value"] = JsonValue.Create(errorBars.Value);
            if (errorBars.NoEndCap) projected["noEndCap"] = JsonValue.Create(true);
            if (errorBars.Line is not null && !string.IsNullOrEmpty(errorBars.Line.Color?.Rgb))
                projected["stroke"] = ProjectChartLine(errorBars.Line);
            output["errorBars"] = projected;
        }
        if (series.DataLabels is not null)
            output["dataLabels"] = ProjectSeriesDataLabels(series.DataLabels);
    }

    private static JsonObject ProjectSeriesDataLabels(SpreadsheetChartSeriesDataLabelsArtifact source)
    {
        var output = source.Defaults is null ? new JsonObject() : ProjectChartLabelOverride(source.Defaults);
        if (source.Points.Count > 0)
        {
            var points = new JsonArray();
            foreach (var point in source.Points.OrderBy(point => point.Index))
            {
                var item = ProjectChartLabelOverride(point.Override!);
                item["index"] = JsonValue.Create(point.Index);
                points.Add(item);
            }
            output["points"] = points;
        }
        return output;
    }

    private static JsonObject ProjectChartLabelOverride(SpreadsheetChartDataLabelOverrideArtifact source)
    {
        var output = new JsonObject();
        if (source.Text.Length > 0) output["text"] = StringNode(source.Text);
        if (source.HasShowValue) output["showValue"] = JsonValue.Create(source.ShowValue);
        if (source.HasShowCategoryName) output["showCategory"] = JsonValue.Create(source.ShowCategoryName);
        if (source.HasShowSeriesName) output["showSeries"] = JsonValue.Create(source.ShowSeriesName);
        if (source.HasShowPercent) output["showPercent"] = JsonValue.Create(source.ShowPercent);
        if (source.HasShowBubbleSize) output["showBubbleSize"] = JsonValue.Create(source.ShowBubbleSize);
        if (source.HasShowLeaderLines) output["showLeaderLines"] = JsonValue.Create(source.ShowLeaderLines);
        if (source.HasPosition && DataLabelPosition(source.Position) is { } position) output["position"] = StringNode(position);
        if (source.TextStyle is not null) output["textStyle"] = ProjectChartTextStyle(source.TextStyle);
        if (source.NumberFormatCode.Length > 0) output["numberFormat"] = StringNode(source.NumberFormatCode);
        if (source.Fill is not null) output["fill"] = ProjectChartSurfaceFill(source.Fill);
        if (source.Line is not null) output["line"] = ProjectChartLine(source.Line);
        return output;
    }

    private static JsonObject ProjectChartAxis(SpreadsheetChartAxisArtifact axis)
    {
        var output = new JsonObject();
        if (!string.IsNullOrEmpty(axis.Title)) output["title"] = StringNode(axis.Title);
        if (!string.IsNullOrEmpty(axis.NumberFormatCode)) output["numberFormat"] = StringNode(axis.NumberFormatCode);
        if (axis.HasTickLabelInterval) output["tickLabelInterval"] = JsonValue.Create(axis.TickLabelInterval);
        if (axis.HasMinimum) output["min"] = JsonValue.Create(axis.Minimum);
        if (axis.HasMaximum) output["max"] = JsonValue.Create(axis.Maximum);
        if (axis.HasMajorUnit) output["majorUnit"] = JsonValue.Create(axis.MajorUnit);
        if (axis.HasMinorUnit) output["minorUnit"] = JsonValue.Create(axis.MinorUnit);
        if (axis.HasPosition) output["position"] = StringNode(axis.Position);
        if (axis.HasVisible) output["visible"] = JsonValue.Create(axis.Visible);
        if (axis.HasReverse) output["reverse"] = JsonValue.Create(axis.Reverse);
        if (axis.HasTickLabelsVisible) output["tickLabelsVisible"] = JsonValue.Create(axis.TickLabelsVisible);
        if (axis.HasTickLabelPosition) output["tickLabelPosition"] = StringNode(axis.TickLabelPosition);
        if (axis.HasMajorTickMark) output["majorTickMark"] = StringNode(axis.MajorTickMark);
        if (axis.HasMinorTickMark) output["minorTickMark"] = StringNode(axis.MinorTickMark);
        if (axis.AxisLine is not null && !string.IsNullOrEmpty(axis.AxisLine.Color?.Rgb))
            output["axisLine"] = ProjectChartLine(axis.AxisLine);
        else if (axis.HasAxisLineVisible)
            output["axisLine"] = JsonValue.Create(axis.AxisLineVisible);
        if (axis.AxisLine is { } axisLine &&
            (axisLine.StartArrow.Length > 0 || axisLine.EndArrow.Length > 0))
        {
            var arrows = new JsonObject();
            if (axisLine.StartArrow.Length > 0) arrows["start"] = StringNode(axisLine.StartArrow);
            if (axisLine.EndArrow.Length > 0) arrows["end"] = StringNode(axisLine.EndArrow);
            output["axisLineArrow"] = arrows;
        }
        if (axis.MajorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MajorGridlineStyle.Color?.Rgb))
            output["gridLine"] = ProjectChartLine(axis.MajorGridlineStyle);
        else if (axis.HasMajorGridlineVisible)
            output["gridLine"] = JsonValue.Create(axis.MajorGridlineVisible);
        else if (axis.HasShowMajorGridlines)
            output["gridLine"] = JsonValue.Create(axis.ShowMajorGridlines);
        if (axis.MinorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MinorGridlineStyle.Color?.Rgb))
            output["minorGridLine"] = ProjectChartLine(axis.MinorGridlineStyle);
        else if (axis.HasMinorGridlineVisible)
            output["minorGridLine"] = JsonValue.Create(axis.MinorGridlineVisible);
        else if (axis.HasShowMinorGridlines)
            output["minorGridLine"] = JsonValue.Create(axis.ShowMinorGridlines);
        if (axis.TextStyle is not null)
            output["textStyle"] = ProjectChartTextStyle(axis.TextStyle);
        if (axis.TitleTextStyle is not null)
            output["titleTextStyle"] = ProjectChartTextStyle(axis.TitleTextStyle);
        return output;
    }

    private static bool TryProjectRadarSpokeAxis(PresentationChart chart, out JsonObject output)
    {
        output = new JsonObject();
        var xAxis = chart.XAxis;
        var yAxis = chart.YAxis;
        if (xAxis is null || yAxis is null ||
            xAxis.Title.Length > 0 || xAxis.NumberFormatCode.Length > 0 || xAxis.HasTickLabelInterval ||
            xAxis.HasMinimum || xAxis.HasMaximum || xAxis.HasMajorUnit || xAxis.HasMinorUnit || xAxis.HasReverse && xAxis.Reverse ||
            xAxis.AxisLine is not null || xAxis.HasAxisLineVisible || xAxis.TextStyle is not null || xAxis.TitleTextStyle is not null ||
            xAxis.HasTickLabelsVisible || xAxis.HasTickLabelPosition ||
            xAxis.HasShowMinorGridlines || xAxis.HasMinorGridlineVisible || xAxis.MinorGridlineStyle is not null ||
            yAxis.Title.Length > 0 || yAxis.HasTickLabelInterval || yAxis.HasReverse && yAxis.Reverse ||
            yAxis.AxisLine is not null || yAxis.HasAxisLineVisible || yAxis.TitleTextStyle is not null ||
            yAxis.HasTickLabelPosition ||
            yAxis.HasShowMinorGridlines || yAxis.HasMinorGridlineVisible || yAxis.MinorGridlineStyle is not null)
            return false;

        if (xAxis.HasVisible != yAxis.HasVisible ||
            xAxis.HasVisible && xAxis.Visible != yAxis.Visible)
            return false;

        var hasEvidence = xAxis.HasVisible || yAxis.HasVisible ||
            xAxis.HasShowMajorGridlines || xAxis.HasMajorGridlineVisible || xAxis.MajorGridlineStyle is not null ||
            yAxis.HasShowMajorGridlines || yAxis.HasMajorGridlineVisible || yAxis.MajorGridlineStyle is not null ||
            yAxis.HasMinimum || yAxis.HasMaximum || yAxis.HasMajorUnit || yAxis.HasMinorUnit ||
            yAxis.HasTickLabelsVisible || yAxis.NumberFormatCode.Length > 0 || yAxis.TextStyle is not null;
        if (!hasEvidence) return false;

        var show = !xAxis.HasVisible || xAxis.Visible;
        if (xAxis.HasVisible) output["show"] = JsonValue.Create(show);
        if (yAxis.HasMinimum) output["min"] = JsonValue.Create(yAxis.Minimum);
        if (yAxis.HasMaximum) output["max"] = JsonValue.Create(yAxis.Maximum);
        if (yAxis.HasMajorUnit) output["majorUnit"] = JsonValue.Create(yAxis.MajorUnit);
        if (yAxis.HasMinorUnit) output["minorUnit"] = JsonValue.Create(yAxis.MinorUnit);

        if (yAxis.HasTickLabelsVisible && !yAxis.TickLabelsVisible)
            output["label"] = JsonValue.Create(false);
        else if (yAxis.NumberFormatCode.Length > 0 || yAxis.TextStyle is not null)
        {
            var label = yAxis.TextStyle is null ? new JsonObject() : ProjectChartTextStyle(yAxis.TextStyle);
            if (yAxis.NumberFormatCode.Length > 0) label["numberFormat"] = StringNode(yAxis.NumberFormatCode);
            output["label"] = label;
        }
        else if (yAxis.HasTickLabelsVisible)
            output["label"] = JsonValue.Create(yAxis.TickLabelsVisible);

        if (show)
        {
            output["axisLine"] = ProjectRadarGuideLine(xAxis) ?? JsonValue.Create(false);
            output["gridLine"] = ProjectRadarGuideLine(yAxis) ?? JsonValue.Create(false);
        }
        return true;
    }

    private static JsonNode? ProjectRadarGuideLine(SpreadsheetChartAxisArtifact axis)
    {
        if (axis.MajorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MajorGridlineStyle.Color?.Rgb))
            return ProjectChartLine(axis.MajorGridlineStyle);
        if (axis.HasMajorGridlineVisible) return JsonValue.Create(axis.MajorGridlineVisible);
        if (axis.HasShowMajorGridlines) return JsonValue.Create(axis.ShowMajorGridlines);
        return null;
    }

    private static JsonObject ProjectChartTextStyle(SpreadsheetChartTextStyleArtifact source)
    {
        var output = new JsonObject();
        if (source.HasFontSizePoints) output["fontSize"] = JsonValue.Create(source.FontSizePoints);
        if (source.FontFamily.Length > 0) output["fontFamily"] = StringNode(source.FontFamily);
        if (source.FontFamilyEastAsia.Length > 0) output["fontFamilyEastAsia"] = StringNode(source.FontFamilyEastAsia);
        if (source.FontFamilyComplexScript.Length > 0) output["fontFamilyComplexScript"] = StringNode(source.FontFamilyComplexScript);
        if (source.HasBold) output["bold"] = JsonValue.Create(source.Bold);
        if (source.HasItalic) output["italic"] = JsonValue.Create(source.Italic);
        if (source.Underline.Length > 0)
            output["underline"] = StringNode(source.Underline switch { "sng" => "single", "dbl" => "double", _ => source.Underline });
        if (source.Alignment.Length > 0)
            output["alignment"] = StringNode(source.Alignment switch { "l" => "left", "ctr" => "center", "r" => "right", "just" => "justify", _ => source.Alignment });
        if (source.Fill is not null)
            output["fill"] = ProjectChartSurfaceFill(source.Fill);
        if (source.ColorRgb.Length > 0)
            output["color"] = TextColor(source.ColorRgb, null,
                source.HasOpacityThousandthPercent, source.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject ProjectChartLine(SpreadsheetChartLineStyleArtifact line)
    {
        var output = new JsonObject
        {
            ["color"] = StringNode(Color(line.Color.Rgb)),
            ["width"] = JsonValue.Create(line.HasWidthPoints ? line.WidthPoints : 0.75),
        };
        if (ChartDash(line.DashStyle) is { } dash) output["dash"] = StringNode(dash);
        if (line.Cap is "flat" or "round" or "square") output["cap"] = StringNode(line.Cap);
        if (line.Join is "miter" or "round" or "bevel") output["join"] = StringNode(line.Join);
        if (line.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(line.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectChartSurfaceFill(SpreadsheetChartSurfaceFill fill)
    {
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            return new JsonObject { ["type"] = StringNode("none") };
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            return Gradient(fill.GradientFill);
        var output = new JsonObject
        {
            ["type"] = StringNode("solid"),
            ["color"] = StringNode(Color(fill.SolidRgb)),
        };
        if (fill.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(fill.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectTable(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var table = element.Table;
        if (table.ColumnWidthsEmu.Count == 0 || table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != table.ColumnWidthsEmu.Count))
            return ProjectOpaque(element, id, nativeRef, "table", "Preserved source table outside the bounded rectangular grid profile.");

        var output = ElementBase(id, element.Name, TableFrame(table), Accessibility(table.Accessibility), nativeRef);
        output["type"] = StringNode("table");
        var columns = new JsonArray();
        for (var index = 0; index < table.ColumnWidthsEmu.Count; index++)
            columns.Add(new JsonObject { ["id"] = StringNode($"column-{index + 1}"), ["width"] = JsonValue.Create(Points(table.ColumnWidthsEmu[index])) });
        output["columns"] = columns;
        output["rows"] = ProjectTableRows(table, context);
        var style = new JsonObject();
        if (table.HasFirstRow) style["headerRows"] = JsonValue.Create(table.FirstRow ? 1 : 0);
        if (table.HasBandedRows) style["bandedRows"] = JsonValue.Create(table.BandedRows);
        if (table.HasBandedColumns) style["bandedColumns"] = JsonValue.Create(table.BandedColumns);
        if (table.HasFirstColumn) style["firstColumnEmphasis"] = JsonValue.Create(table.FirstColumn);
        if (table.HasLastColumn) style["lastColumnEmphasis"] = JsonValue.Create(table.LastColumn);
        if (table.HasLastRow) style["lastRow"] = JsonValue.Create(table.LastRow);
        if (style.Count > 0) output["style"] = style;
        return output;
    }

    private static JsonArray ProjectTableRows(PresentationTable table, ProjectionContext context)
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
                var nativeCell = table.Rows[rowIndex].Cells[columnIndex];
                var cell = new JsonObject
                {
                    ["id"] = StringNode($"cell-{rowIndex + 1}-{columnIndex + 1}"),
                    // Keep the legacy string form for uniform/opaque cells.
                    // A bounded mixed-run cell carries its fixed-topology
                    // text body so each direct run style remains editable.
                    ["text"] = TextContent(nativeCell.TextBody, nativeCell.Text, context, includeBodyStyle: true),
                };
                if (mergeByOrigin.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    cell["rowSpan"] = JsonValue.Create(checked((int)(merge.EndRow - merge.StartRow + 1)));
                    cell["columnSpan"] = JsonValue.Create(checked((int)(merge.EndColumn - merge.StartColumn + 1)));
                }
                if (nativeCell.Fill is { } fill && ProjectTableCellFill(fill, context) is { } projectedFill)
                    cell["fill"] = projectedFill;
                if (nativeCell.Borders is { } borders)
                    cell["borders"] = ProjectTableCellBorders(borders);
                if (nativeCell.TextStyle is { } textStyle && ProjectTextStyle(textStyle) is { Count: > 0 } projectedTextStyle)
                    cell["textStyle"] = new JsonObject { ["defaultText"] = projectedTextStyle };
                cells.Add(cell);
            }
            rows.Add(new JsonObject
            {
                ["id"] = StringNode($"row-{rowIndex + 1}"),
                ["height"] = JsonValue.Create(Points(table.Rows[rowIndex].HeightEmu)),
                ["cells"] = cells,
            });
        }
        return rows;
    }

    private static JsonObject? ProjectTableCellFill(PresentationTableCellFill fill, ProjectionContext context)
    {
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.SolidRgb)
        {
            var output = new JsonObject
            {
                ["type"] = StringNode("solid"),
                ["color"] = StringNode(Color(fill.SolidRgb)),
            };
            if (fill.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(fill.OpacityThousandthPercent));
            return output;
        }
        return fill.KindCase switch
        {
            PresentationTableCellFill.KindOneofCase.NoFill => new JsonObject { ["type"] = StringNode("none") },
            PresentationTableCellFill.KindOneofCase.GradientFill => Gradient(fill.GradientFill),
            PresentationTableCellFill.KindOneofCase.ImagePaint => ProjectImagePaint(fill.ImagePaint, context),
            _ => null,
        };
    }

    private static JsonObject ProjectTableCellBorders(PresentationTableCellBorders borders)
    {
        var output = new JsonObject();
        if (borders.Left is { } left) output["left"] = ProjectChartLine(left);
        if (borders.Top is { } top) output["top"] = ProjectChartLine(top);
        if (borders.Right is { } right) output["right"] = ProjectChartLine(right);
        if (borders.Bottom is { } bottom) output["bottom"] = ProjectChartLine(bottom);
        return output;
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
        output["type"] = StringNode("connector");
        output["connectorType"] = StringNode(connector.ConnectorType is "elbow" or "curved" ? connector.ConnectorType : "straight");
        output["from"] = ConnectorEndpoint(connector.StartTargetId, connector.StartXEmu, connector.StartYEmu, pageId, context);
        output["to"] = ConnectorEndpoint(connector.EndTargetId, connector.EndXEmu, connector.EndYEmu, pageId, context);
        // A connector has one native line-alpha owner. Authored
        // compositing.opacity is therefore projected as the effective stroke
        // opacity rather than as a second, unrecoverable field.
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
        output["type"] = StringNode("group");
        output["childFrame"] = Frame(group.ChildLeftEmu, group.ChildTopEmu, group.ChildWidthEmu, group.ChildHeightEmu);
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
        if (children.Count > 0)
        {
            // DrawingML has no separate group-level accessibility order
            // owner in the bounded profile.  The local child shape-tree order
            // is therefore projected explicitly so a subsequent PPJ edit can
            // request a complete, auditable permutation without flattening
            // the group.
            var readingOrder = new JsonArray();
            foreach (var child in children)
            {
                if (child is JsonObject childObject && childObject["id"] is JsonValue childId)
                    readingOrder.Add(childId.GetValue<string>());
            }
            if (readingOrder.Count == children.Count)
                output["readingOrder"] = readingOrder;
        }
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
        var output = ElementBase(id, element.Name, ElementFrame(element), Accessibility(opaque?.Accessibility), nativeRef);
        output["type"] = StringNode("opaque");
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

    private static JsonObject ProjectSourceSmartArt(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context)
    {
        var diagram = element.Opaque.DiagramText;
        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = StringNode("smartArt");
        output["mode"] = StringNode("source-bound");
        if (context.SmartArtNativeSections(element.Opaque.PreservedPartPaths) is { } nativeSections)
            output["nativeSections"] = nativeSections;
        if (!string.IsNullOrWhiteSpace(diagram.LayoutDefinitionId))
            output["layoutDefinitionId"] = StringNode(diagram.LayoutDefinitionId);
        var nodes = new JsonArray();
        var nodeIds = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < diagram.Nodes.Count; index++)
        {
            var source = diagram.Nodes[index];
            var nodeId = context.UniqueId($"{id}-node-{NormalizeId(source.ModelId, $"node-{index + 1}")}");
            nodeIds.Add(source.ModelId, nodeId);
            var values = source.RunTexts.Count > 0 ? source.RunTexts.ToArray() : [source.Text];
            var hash = Sha256(Encoding.UTF8.GetBytes(source.ModelId + "\0" + string.Join("\0", values)));
            var capabilities = new[] { new CapabilitySpec("setSmartArtText", ["smartArt.text"]) };
            var node = new JsonObject
            {
                ["id"] = StringNode(nodeId),
                ["text"] = SmartArtText(values),
                ["nativeRef"] = NativeRef(
                    context,
                    $"element:{pageId}:{id}:smartArtNode:{index}",
                    hash,
                    capabilities),
            };
            node["kind"] = StringNode(source.PointType switch
            {
                "asst" => "assistant",
                "doc" => "document",
                _ => "node",
            });
            nodes.Add(node);
        }
        output["nodes"] = nodes;
        if (diagram.Connections.Count > 0)
        {
            var connections = new JsonArray();
            for (var index = 0; index < diagram.Connections.Count; index++)
            {
                var source = diagram.Connections[index];
                if (!nodeIds.TryGetValue(source.FromModelId, out var fromId) ||
                    !nodeIds.TryGetValue(source.ToModelId, out var toId))
                    throw new CodecException(
                        "ppj.smartArt.invalid_source_graph",
                        "A proven source-bound SmartArt parent edge no longer resolves to projected content nodes.",
                        $"elements.{id}.connections[{index}]");
                var connection = new JsonObject
                {
                    ["id"] = StringNode(context.UniqueId($"{id}-connection-{NormalizeId(source.ModelId, $"connection-{index + 1}")}")),
                    ["from"] = StringNode(fromId),
                    ["to"] = StringNode(toId),
                    ["role"] = StringNode("parent"),
                };
                if (source.Order > 0) connection["order"] = JsonValue.Create(source.Order);
                connections.Add(connection);
            }
            output["connections"] = connections;
        }
        return output;
    }

    private static JsonObject ProjectNativeSmartArt(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var diagram = element.Diagram;
        var output = ElementBase(id, element.Name, DiagramFrame(diagram), Accessibility(diagram.Accessibility), nativeRef);
        output["type"] = StringNode("smartArt");
        output["mode"] = StringNode("source-bound");
        if (!string.IsNullOrWhiteSpace(diagram.Layout)) output["layout"] = StringNode(diagram.Layout);
        if (!string.IsNullOrWhiteSpace(diagram.DefinitionAssetId) &&
            context.TryMaterializeAsset(diagram.DefinitionAssetId, out var definitionAssetId))
        {
            output.Remove("layout");
            output["definitionAsset"] = StringNode(definitionAssetId);
        }
        var nodes = new JsonArray();
        foreach (var node in diagram.Nodes)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(node.Id),
                ["text"] = TextContent(node.TextBody, PptxTextCodec.Flatten(node.TextBody), context),
            };
            if (!string.IsNullOrWhiteSpace(node.AssetId) && context.TryMaterializeAsset(node.AssetId, out var assetId))
                projected["asset"] = StringNode(assetId);
            var cachedShape = diagram.Drawing?.Children
                .SingleOrDefault(child => child.Name == node.Id && child.Shape is not null)?.Shape;
            if (cachedShape?.ImageFill is { } imagePaint &&
                context.TryMaterializeAsset(imagePaint.AssetId, out _))
                projected["image"] = ProjectSmartArtNodeImage(imagePaint);
            nodes.Add(projected);
        }
        output["nodes"] = nodes;
        if (diagram.Connections.Count > 0)
        {
            var connections = new JsonArray();
            foreach (var connection in diagram.Connections)
            {
                var projected = new JsonObject
                {
                    ["id"] = StringNode(connection.Id),
                    ["from"] = StringNode(connection.FromId),
                    ["to"] = StringNode(connection.ToId),
                    ["role"] = StringNode(connection.Role),
                };
                if (connection.Order > 0) projected["order"] = JsonValue.Create(connection.Order);
                connections.Add(projected);
            }
            output["connections"] = connections;
        }
        return output;
    }

    private static JsonObject ProjectSmartArtNodeImage(PresentationImagePaint paint)
    {
        var output = new JsonObject
        {
            ["fit"] = StringNode(paint.Mode == PresentationImagePaint.Types.Mode.Tile ? "tile" : "stretch"),
        };
        if (paint.Crop is { } crop)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(crop.BottomThousandthPercent)),
            };
        }
        if (paint.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(paint.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectSourceOle(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var workbook = element.Opaque.OleWorkbook;
        var package = element.Opaque.OleOfficePackage;
        var partPath = workbook?.PartPath ?? package?.PartPath;
        var contentType = workbook?.ContentType ?? package?.ContentType;
        if (string.IsNullOrWhiteSpace(partPath) || string.IsNullOrWhiteSpace(contentType) ||
            !context.TryMaterializeSourcePart(partPath, contentType, out var payloadAssetId))
            return ProjectOpaque(element, id, nativeRef, "oleObject", "Preserved source OLE object whose payload could not be materialized safely.");

        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = StringNode("ole");
        output["payloadAsset"] = payloadAssetId;
        return output;
    }

    private static JsonNode SmartArtText(IReadOnlyList<string> values)
    {
        if (values.Count == 1) return StringNode(values[0]);
        var runs = new JsonArray();
        for (var index = 0; index < values.Count; index++)
            runs.Add(new JsonObject { ["id"] = StringNode($"run-{index + 1}"), ["text"] = StringNode(values[index]) });
        return new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject { ["id"] = StringNode("paragraph-1"), ["runs"] = runs },
            },
        };
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
            ["id"] = StringNode(id),
            ["frame"] = frame,
            ["nativeRef"] = nativeRef,
        };
        if (!string.IsNullOrWhiteSpace(name)) output["name"] = StringNode(name);
        if (accessibility is not null) output["accessibility"] = accessibility;
        return output;
    }

    private static JsonObject? ProjectAction(PresentationRunHyperlink? source, ProjectionContext context)
    {
        if (source is null) return null;
        var output = new JsonObject();
        switch (source.TargetCase)
        {
            case PresentationRunHyperlink.TargetOneofCase.Uri:
                output["uri"] = StringNode(source.Uri);
                break;
            case PresentationRunHyperlink.TargetOneofCase.SlideId:
                if (!context.TryPageId(source.SlideId, out var pageId)) return null;
                output["slide"] = StringNode(pageId);
                break;
            case PresentationRunHyperlink.TargetOneofCase.CustomShowId:
                if (!context.TryCustomShowId(source.CustomShowId, out var customShowId)) return null;
                output["customShow"] = StringNode(customShowId);
                if (source.HasReturnToSlide) output["returnToSlide"] = JsonValue.Create(source.ReturnToSlide);
                break;
            case PresentationRunHyperlink.TargetOneofCase.Action:
                output["verb"] = StringNode(source.Action);
                break;
            default:
                return null;
        }
        if (source.HasTooltip) output["tooltip"] = StringNode(source.Tooltip);
        if (source.HasTargetFrame) output["targetFrame"] = StringNode(source.TargetFrame);
        if (source.HasHistory) output["history"] = JsonValue.Create(source.History);
        if (source.HasHighlightClick) output["highlightClick"] = JsonValue.Create(source.HighlightClick);
        return output;
    }

    private static JsonObject ShapeStyle(
        PresentationShape shape,
        ProjectionContext context,
        bool omitOwnerOpacity = false)
    {
        var style = new JsonObject();
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(shape.FillRgb)) };
            if (!omitOwnerOpacity && shape.HasFillOpacityThousandthPercent)
                fill["opacity"] = JsonValue.Create(Unit(shape.FillOpacityThousandthPercent));
            style["fill"] = fill;
        }
        else if (shape.GradientFill is not null)
        {
            style["fill"] = Gradient(shape.GradientFill, includeOpacity: !omitOwnerOpacity);
        }
        else if (shape.ImageFill is not null && ProjectImagePaint(shape.ImageFill, context, includeOpacity: !omitOwnerOpacity) is { } imageFill)
        {
            style["fill"] = imageFill;
        }
        else if (!string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
                 context.TryMaterializeAsset(shape.ImageFillAssetId, out var sourceImageAssetId))
        {
            style["fill"] = new JsonObject
            {
                ["type"] = StringNode("image"),
                ["asset"] = StringNode(sourceImageAssetId),
                ["fit"] = StringNode("stretch"),
            };
        }
        if ((!string.IsNullOrEmpty(shape.LineRgb) || !string.IsNullOrEmpty(shape.LineScheme)) && shape.LineStyle != "none")
            style["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin,
                !omitOwnerOpacity && shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                shape.LineScheme);
        if (shape.Shadow is not null && (!string.IsNullOrEmpty(shape.Shadow.ColorRgb) || !string.IsNullOrEmpty(shape.Shadow.ColorScheme)))
            style["shadow"] = Shadow(shape.Shadow, includeOpacity: !omitOwnerOpacity);
        if (shape.Glow is not null && (!string.IsNullOrEmpty(shape.Glow.ColorRgb) || !string.IsNullOrEmpty(shape.Glow.ColorScheme)))
            style["glow"] = Glow(shape.Glow);
        if (shape.InnerShadow is not null && (!string.IsNullOrEmpty(shape.InnerShadow.ColorRgb) || !string.IsNullOrEmpty(shape.InnerShadow.ColorScheme)))
            style["innerShadow"] = InnerShadow(shape.InnerShadow);
        if (shape.Reflection is not null)
            style["reflection"] = Reflection(shape.Reflection);
        if (shape.SoftEdge is not null)
            style["softEdge"] = SoftEdge(shape.SoftEdge);
        return style;
    }

    private static JsonNode TextContent(
        PresentationTextBody? body,
        string? fallback,
        ProjectionContext context,
        bool includeBodyStyle = false)
    {
        if (body is null || body.Paragraphs.Count == 0 ||
            body.Paragraphs.Any(paragraph => paragraph.Runs.Count == 0 ||
                paragraph.Runs.Any(run => run.ContentCase is not (PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.LineBreak or PresentationTextRun.ContentOneofCase.Field))))
            return StringNode(fallback ?? string.Empty);

        var paragraphs = new JsonArray();
        for (var paragraphIndex = 0; paragraphIndex < body.Paragraphs.Count; paragraphIndex++)
        {
            var source = body.Paragraphs[paragraphIndex];
            var paragraph = new JsonObject
            {
                ["id"] = StringNode($"paragraph-{paragraphIndex + 1}"),
            };
            var paragraphStyle = new JsonObject();
            if (source.HasLevel) paragraphStyle["level"] = JsonValue.Create(checked((int)source.Level));
            if (source.HasAlignment && ParagraphAlignment(source.Alignment) is { } alignment)
                paragraphStyle["alignment"] = StringNode(alignment);
            if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                paragraphStyle["indent"] = JsonValue.Create(Points(source.MarginLeftEmu));
            if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                paragraphStyle["hanging"] = JsonValue.Create(-Points(source.IndentEmu));
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                paragraphStyle["lineSpacing"] = JsonValue.Create(Math.Max(0.001, source.LineSpacingPoints));
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                paragraphStyle["lineSpacingMultiplier"] = JsonValue.Create(Math.Max(0.00001, source.LineSpacingMultiplier));
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                paragraphStyle["spaceBefore"] = JsonValue.Create(Math.Max(0, source.SpaceBeforePoints));
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                paragraphStyle["spaceBeforeMultiplier"] = JsonValue.Create(Math.Max(0, source.SpaceBeforeMultiplier));
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                paragraphStyle["spaceAfter"] = JsonValue.Create(Math.Max(0, source.SpaceAfterPoints));
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                paragraphStyle["spaceAfterMultiplier"] = JsonValue.Create(Math.Max(0, source.SpaceAfterMultiplier));
            if (source.TabStops.Count > 0)
            {
                var tabStops = new JsonArray();
                foreach (var tab in source.TabStops)
                    tabStops.Add(new JsonObject
                    {
                        ["position"] = JsonValue.Create(Points(tab.PositionEmu)),
                        ["alignment"] = StringNode(tab.Alignment),
                    });
                paragraphStyle["tabStops"] = tabStops;
            }
            if (source.HasNoTabStops && source.NoTabStops)
                paragraphStyle["noTabStops"] = true;
            if (source.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                ProjectTextStyle(source.DefaultRunProperties) is { Count: > 0 } defaultText)
                paragraphStyle["defaultText"] = defaultText;
            if (ProjectBullet(source, context) is { } bullet) paragraphStyle["bullet"] = bullet;
            if (paragraphStyle.Count > 0) paragraph["style"] = paragraphStyle;

            var runs = new JsonArray();
            for (var runIndex = 0; runIndex < source.Runs.Count; runIndex++)
            {
                var sourceRun = source.Runs[runIndex];
                var run = new JsonObject { ["id"] = StringNode($"run-{paragraphIndex + 1}-{runIndex + 1}") };
                if (sourceRun.ContentCase == PresentationTextRun.ContentOneofCase.Field)
                {
                    var field = new JsonObject
                    {
                        ["type"] = StringNode(sourceRun.Field.Type),
                        ["text"] = StringNode(sourceRun.Field.Text),
                    };
                    if (PptxTextCodec.IsAutomaticFieldType(sourceRun.Field.Type)) field["automatic"] = true;
                    if (!string.IsNullOrWhiteSpace(sourceRun.Field.Id)) field["id"] = StringNode(sourceRun.Field.Id);
                    run["field"] = field;
                }
                else if (sourceRun.ContentCase == PresentationTextRun.ContentOneofCase.LineBreak)
                    run["break"] = true;
                else run["text"] = StringNode(sourceRun.Text);
                var style = RunStyle(sourceRun);
                if (style.Count > 0) run["style"] = style;
                if (sourceRun.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.RunHyperlink &&
                    sourceRun.RunHyperlink.TargetCase == PresentationRunHyperlink.TargetOneofCase.Uri &&
                    !string.IsNullOrWhiteSpace(sourceRun.RunHyperlink.Uri))
                    run["hyperlink"] = new JsonObject { ["uri"] = StringNode(sourceRun.RunHyperlink.Uri) };
                runs.Add(run);
            }
            paragraph["runs"] = runs;
            paragraphs.Add(paragraph);
        }
        var output = new JsonObject { ["paragraphs"] = paragraphs };
        if (includeBodyStyle && TextBoxStyle(body) is { Count: > 0 } bodyStyle)
            output["style"] = bodyStyle;
        return output;
    }

    private static JsonObject RunStyle(PresentationTextRun run)
    {
        var style = new JsonObject();
        if (run.HasFontFamily) style["fontFamily"] = StringNode(run.FontFamily);
        if (run.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = StringNode(run.FontFamilyEastAsia);
        if (run.HasFontFamilyComplexScript) style["fontFamilyComplexScript"] = StringNode(run.FontFamilyComplexScript);
        if (run.HasFontSizePoints && run.FontSizePoints > 0) style["size"] = JsonValue.Create(run.FontSizePoints);
        if (run.HasBold) style["bold"] = JsonValue.Create(run.Bold);
        if (run.HasItalic) style["italic"] = JsonValue.Create(run.Italic);
        if (run.HasColorRgb && !string.IsNullOrEmpty(run.ColorRgb))
            style["color"] = TextColor(run.ColorRgb, null, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.HasColorScheme && !string.IsNullOrEmpty(run.ColorScheme))
            style["color"] = TextColor(null, run.ColorScheme, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.GradientFill is not null)
            style["gradient"] = TextGradient(run.GradientFill);
        if (run.Shadow is not null) style["shadow"] = Shadow(run.Shadow);
        if (run.Glow is not null) style["glow"] = Glow(run.Glow);
        if (run.InnerShadow is not null) style["innerShadow"] = InnerShadow(run.InnerShadow);
        if (run.Reflection is not null) style["reflection"] = Reflection(run.Reflection);
        if (run.SoftEdge is not null) style["softEdge"] = SoftEdge(run.SoftEdge);
        if (run.HighlightCase == PresentationTextRun.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(run.HighlightRgb))
            style["highlight"] = StringNode(Color(run.HighlightRgb));
        if (run.HasUnderline) style["underline"] = StringNode(run.Underline switch { "sng" => "single", "dbl" => "double", _ => run.Underline });
        if (run.HasStrike) style["strike"] = JsonValue.Create(run.Strike);
        if (run.HasFontKerningPoints) style["kerning"] = JsonValue.Create(run.FontKerningPoints);
        if (run.HasFontBaselinePercent) style["baseline"] = JsonValue.Create(run.FontBaselinePercent);
        if (run.HasFontSpacingPoints) style["letterSpacing"] = JsonValue.Create(run.FontSpacingPoints);
        if (run.HasFontCaps) style["capitalization"] = StringNode(run.FontCaps);
        if (run.HasLanguage) style["language"] = StringNode(run.Language);
        return style;
    }

    private static JsonObject ProjectTextStyle(PresentationTextStyle source)
    {
        var style = new JsonObject();
        if (source.HasFontFamily) style["fontFamily"] = StringNode(source.FontFamily);
        if (source.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = StringNode(source.FontFamilyEastAsia);
        if (source.HasFontFamilyComplexScript) style["fontFamilyComplexScript"] = StringNode(source.FontFamilyComplexScript);
        if (source.HasFontSizePoints && source.FontSizePoints > 0) style["size"] = JsonValue.Create(source.FontSizePoints);
        if (source.HasBold) style["bold"] = JsonValue.Create(source.Bold);
        if (source.HasItalic) style["italic"] = JsonValue.Create(source.Italic);
        if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorRgb && !string.IsNullOrEmpty(source.ColorRgb))
            style["color"] = TextColor(source.ColorRgb, null, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorScheme && !string.IsNullOrEmpty(source.ColorScheme))
            style["color"] = TextColor(null, source.ColorScheme, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.GradientFill is not null)
            style["gradient"] = TextGradient(source.GradientFill);
        if (source.Shadow is not null) style["shadow"] = Shadow(source.Shadow);
        if (source.Glow is not null) style["glow"] = Glow(source.Glow);
        if (source.InnerShadow is not null) style["innerShadow"] = InnerShadow(source.InnerShadow);
        if (source.Reflection is not null) style["reflection"] = Reflection(source.Reflection);
        if (source.SoftEdge is not null) style["softEdge"] = SoftEdge(source.SoftEdge);
        if (source.HighlightCase == PresentationTextStyle.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(source.HighlightRgb))
            style["highlight"] = StringNode(Color(source.HighlightRgb));
        if (source.HasUnderline) style["underline"] = StringNode(source.Underline switch { "sng" => "single", "dbl" => "double", _ => source.Underline });
        if (source.HasStrike) style["strike"] = JsonValue.Create(source.Strike);
        if (source.HasFontKerningPoints) style["kerning"] = JsonValue.Create(source.FontKerningPoints);
        if (source.HasFontBaselinePercent) style["baseline"] = JsonValue.Create(source.FontBaselinePercent);
        if (source.HasFontSpacingPoints) style["letterSpacing"] = JsonValue.Create(source.FontSpacingPoints);
        if (source.HasFontCaps) style["capitalization"] = StringNode(source.FontCaps);
        if (source.HasLanguage) style["language"] = StringNode(source.Language);
        return style;
    }

    private static JsonObject? ProjectBullet(PresentationTextParagraph paragraph, ProjectionContext context)
    {
        JsonObject? bullet = paragraph.BulletCase switch
        {
            PresentationTextParagraph.BulletOneofCase.NoBullet => new JsonObject { ["type"] = StringNode("none") },
            PresentationTextParagraph.BulletOneofCase.BulletCharacter => new JsonObject
            {
                ["type"] = StringNode("character"),
                ["character"] = StringNode(paragraph.BulletCharacter),
            },
            PresentationTextParagraph.BulletOneofCase.AutoNumber => new JsonObject
            {
                ["type"] = StringNode("number"),
                ["scheme"] = StringNode(paragraph.AutoNumber.Scheme),
            },
            PresentationTextParagraph.BulletOneofCase.PictureBullet => ProjectPictureBullet(paragraph.PictureBullet, context),
            _ => null,
        };
        if (bullet is null) return null;
        // Font, color, and size metadata on a noBullet paragraph are inert
        // native residue. Do not emit them into the closed `none` PPJ shape;
        // the source-bound leaves still preserve and edit those tokens.
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.NoBullet)
            return bullet;
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber && paragraph.AutoNumber.HasStartAt)
            bullet["startAt"] = JsonValue.Create(checked((int)paragraph.AutoNumber.StartAt));
        if (paragraph.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily)
            bullet["fontFamily"] = StringNode(paragraph.BulletFontFamily);
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
            bullet["size"] = JsonValue.Create(paragraph.BulletSizePoints);
        else if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent)
            bullet["sizePercent"] = JsonValue.Create(paragraph.BulletSizePercent);
        return bullet;
    }

    private static JsonObject? ProjectPictureBullet(PresentationPictureBullet picture, ProjectionContext context)
    {
        var bullet = new JsonObject { ["type"] = StringNode("picture") };
        if (picture.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
        {
            if (!context.TryMaterializeAsset(picture.AssetId, out var assetId)) return null;
            bullet["asset"] = StringNode(assetId);
        }
        else if (picture.SourceCase == PresentationPictureBullet.SourceOneofCase.Uri)
            bullet["uri"] = StringNode(picture.Uri);
        else
            return null;
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
                ["id"] = StringNode(item.Id),
                ["value"] = StringNode(item.Value),
                ["role"] = StringNode("Fallback only; imported native styling remains source-owned"),
            });
        return output;
    }

    private static JsonObject TextBoxStyle(PresentationTextBody? body)
    {
        var output = new JsonObject();
        var properties = body?.BodyProperties;
        if (properties is null) return output;
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor)
            output["verticalAlignment"] = StringNode(properties.VerticalAnchor == "center" ? "middle" : properties.VerticalAnchor);
        if (properties.HasAnchorCenter)
            output["anchorCenter"] = JsonValue.Create(properties.AnchorCenter);
        if (properties.HasForceAntiAlias)
            output["forceAntiAlias"] = JsonValue.Create(properties.ForceAntiAlias);
        if (properties.HasSpaceFirstLastParagraph)
            output["spaceFirstLastParagraph"] = JsonValue.Create(properties.SpaceFirstLastParagraph);
        if (properties.HasCompatibleLineSpacing)
            output["compatibleLineSpacing"] = JsonValue.Create(properties.CompatibleLineSpacing);
        if (properties.HasFromWordArt)
            output["fromWordArt"] = JsonValue.Create(properties.FromWordArt);
        if (properties.HasTextWarpPreset)
            output["textWarpPreset"] = StringNode(properties.TextWarpPreset);
        if (properties.TextWarpAdjustments.Count > 0)
        {
            output["textWarpAdjustments"] = new JsonArray(properties.TextWarpAdjustments.Select(adjustment =>
                new JsonObject
                {
                    ["name"] = StringNode(adjustment.Name),
                    ["value"] = JsonValue.Create(adjustment.Value),
                }).ToArray());
        }
        if (properties.HasFlatTextZ)
            output["flatTextZ"] = JsonValue.Create(properties.FlatTextZ);
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
            output["wrap"] = StringNode(properties.Wrap);
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode)
            output["autoFit"] = StringNode(properties.AutoFitMode switch { "shrinkText" => "shrink-text", "resizeShape" => "resize-shape", _ => "none" });
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode &&
            properties.AutoFitMode == "shrinkText" && properties.NormalAutoFit is { } normalAutoFit &&
            (normalAutoFit.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000 ||
             normalAutoFit.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000))
        {
            var normal = new JsonObject();
            if (normalAutoFit.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000)
                normal["fontScale"] = JsonValue.Create(normalAutoFit.FontScale1000 / 1_000d);
            if (normalAutoFit.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000)
                normal["lineSpacingReduction"] = JsonValue.Create(normalAutoFit.LineSpacingReduction1000 / 1_000d);
            output["normalAutoFit"] = normal;
        }
        var margins = new JsonObject();
        if (properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) margins["left"] = JsonValue.Create(Points(properties.LeftInsetEmu));
        if (properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) margins["top"] = JsonValue.Create(Points(properties.TopInsetEmu));
        if (properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) margins["right"] = JsonValue.Create(Points(properties.RightInsetEmu));
        if (properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) margins["bottom"] = JsonValue.Create(Points(properties.BottomInsetEmu));
        if (margins.Count > 0) output["margins"] = margins;
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
            output["columns"] = JsonValue.Create(checked((int)properties.Columns));
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu)
            output["columnGap"] = JsonValue.Create(Points(properties.ColumnSpacingEmu));
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns)
            output["columnDirection"] = StringNode(properties.RightToLeftColumns ? "right-to-left" : "left-to-right");
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode)
            output["verticalText"] = StringNode(properties.VerticalTextMode);
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000)
            output["rotation"] = JsonValue.Create(properties.RotationAngle60000 / 60_000d);
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode)
            output["verticalOverflow"] = StringNode(properties.VerticalOverflowMode);
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode)
            output["horizontalOverflow"] = StringNode(properties.HorizontalOverflowMode);
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright)
            output["upright"] = JsonValue.Create(properties.Upright);
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
            var fill = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(shape.FillRgb)) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = JsonValue.Create(Unit(shape.FillOpacityThousandthPercent));
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
            return new JsonObject { ["type"] = StringNode("image"), ["asset"] = StringNode(assetId), ["fit"] = StringNode("stretch") };
        if (background.GradientFill is not null)
            return Gradient(background.GradientFill);
        if (!string.IsNullOrEmpty(background.ColorRgb))
        {
            var output = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(background.ColorRgb)) };
            if (background.HasOpacityThousandthPercent)
                output["opacity"] = JsonValue.Create(Unit(background.OpacityThousandthPercent));
            return output;
        }
        return null;
    }

    private static JsonObject? ProjectImagePaint(
        PresentationImagePaint paint,
        ProjectionContext context,
        bool includeOpacity = true)
    {
        if (!context.TryMaterializeAsset(paint.AssetId, out var assetId)) return null;
        var output = new JsonObject
        {
            ["type"] = StringNode("image"),
            ["asset"] = StringNode(assetId),
            ["fit"] = StringNode(paint.Mode == PresentationImagePaint.Types.Mode.Tile ? "tile" : "stretch"),
        };
        if (paint.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(paint.Crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(paint.Crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(paint.Crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(paint.Crop.BottomThousandthPercent)),
            };
        }
        if (includeOpacity && paint.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(paint.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject Gradient(PresentationGradientFill source, bool includeOpacity = true)
    {
        var stops = new JsonArray();
        foreach (var stop in source.Stops)
        {
            var item = new JsonObject
            {
                ["offset"] = JsonValue.Create(Unit(stop.PositionThousandthPercent)),
                ["color"] = StringNode(Color(stop.ColorRgb)),
            };
            if (includeOpacity && stop.HasOpacityThousandthPercent)
                item["opacity"] = JsonValue.Create(Unit(stop.OpacityThousandthPercent));
            stops.Add(item);
        }
        var output = new JsonObject
        {
            ["type"] = StringNode("gradient"),
            ["kind"] = StringNode(source.Kind == PresentationGradientFill.Types.Kind.Radial ? "radial" : "linear"),
            ["stops"] = stops,
        };
        if (source.Kind == PresentationGradientFill.Types.Kind.Linear && source.HasAngle60000)
            output["angle"] = JsonValue.Create(source.Angle60000 / 60_000d);
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
                ["id"] = StringNode(context.UniqueId($"animation-{animation.Id}")),
                ["target"] = StringNode(targetId),
                ["phase"] = StringNode(animation.Phase),
                ["effect"] = StringNode(animation.Effect),
                ["start"] = StringNode(animation.Start),
                // PresentationML exposes the effective start condition, not
                // the PPJ sugar distinction.  Emit the normalized trigger as
                // well so imported click/previous timing remains a typed,
                // inspectable field instead of disappearing during projection.
                ["trigger"] = StringNode(animation.Start),
                ["durationMs"] = JsonValue.Create(animation.HasDurationMs ? checked((int)animation.DurationMs) : 500),
            };
            if (!string.IsNullOrEmpty(animation.Direction)) item["direction"] = StringNode(animation.Direction);
            if (animation.HasDelayMs) item["delayMs"] = JsonValue.Create(checked((int)animation.DelayMs));
            if (!string.IsNullOrEmpty(animation.TextBuild)) item["textBuild"] = StringNode(animation.TextBuild);
            if (!string.IsNullOrEmpty(animation.ChartBuild)) item["chartBuild"] = StringNode(animation.ChartBuild);
            if (animation.HasStaggerMs) item["staggerMs"] = JsonValue.Create(checked((int)animation.StaggerMs));
            if (animation.HasAnimateChartBackground) item["animateChartBackground"] = JsonValue.Create(animation.AnimateChartBackground);
            if (animation.HasRepeatCount) item["repeat"] = JsonValue.Create(checked((int)animation.RepeatCount));
            if (animation.HasAutoReverse) item["autoReverse"] = JsonValue.Create(animation.AutoReverse);
            if (!string.IsNullOrEmpty(animation.Easing) && animation.Easing != "linear") item["easing"] = StringNode(animation.Easing);
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
                pairs.Add(new JsonObject { ["key"] = StringNode(context.UniqueId($"morph-{pair.Key}")), ["from"] = StringNode(fromId), ["to"] = StringNode(toId) });
            }
            if (pairs.Count == 0) return null;
            return new JsonObject
            {
                ["type"] = StringNode("morph"),
                ["durationMs"] = JsonValue.Create(slide.Morph.HasDurationMs ? checked((int)slide.Morph.DurationMs) : 800),
                ["fromPage"] = StringNode(context.PageId(fromPage.Id)),
                ["morphPairs"] = pairs,
            };
        }
        var transition = slide.Transition;
        if (transition is null || !PpjTransitionLowering.IsBaseEffect(transition.Effect)) return null;
        var output = new JsonObject
        {
            ["type"] = StringNode(transition.Effect),
            ["speed"] = StringNode(transition.Speed),
            ["advanceOnClick"] = JsonValue.Create(transition.AdvanceOnClick),
        };
        if (transition.HasDurationMs) output["durationMs"] = JsonValue.Create(checked((int)transition.DurationMs));
        if (!string.IsNullOrEmpty(transition.Direction)) output["direction"] = StringNode(transition.Direction);
        if (!string.IsNullOrEmpty(transition.Orientation)) output["orientation"] = StringNode(transition.Orientation);
        if (transition.HasThroughBlack) output["throughBlack"] = JsonValue.Create(transition.ThroughBlack);
        if (transition.HasSpokes) output["spokes"] = JsonValue.Create(checked((int)transition.Spokes));
        if (transition.HasAdvanceAfterMs) output["advanceAfterMs"] = JsonValue.Create(checked((int)transition.AdvanceAfterMs));
        return output;
    }

    private static JsonArray ProjectSections(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        for (var sectionIndex = 0; sectionIndex < presentation.Sections.Count; sectionIndex++)
        {
            var section = presentation.Sections[sectionIndex];
            var pages = new JsonArray();
            foreach (var id in section.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            var item = new JsonObject
            {
                ["id"] = StringNode(context.UniqueId($"section-{section.Id}")),
                ["name"] = StringNode(string.IsNullOrWhiteSpace(section.Name) ? "Section" : section.Name),
                ["pages"] = pages,
            };
            if (section.Source is { } source)
            {
                var capabilities = source.Editable
                    ? new[] { new CapabilitySpec("setName", ["name"]), new CapabilitySpec("setPages", ["pages"]) }
                    : [];
                item["nativeRef"] = NativeRef(
                    context,
                    $"section:{sectionIndex}",
                    HashOrFallback(source.SectionXmlSha256, section),
                    capabilities);
            }
            output.Add(item);
        }
        return output;
    }

    private static JsonArray ProjectCustomShows(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        for (var showIndex = 0; showIndex < presentation.CustomShows.Count; showIndex++)
        {
            var show = presentation.CustomShows[showIndex];
            var pages = new JsonArray();
            foreach (var id in show.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            var item = new JsonObject
            {
                ["id"] = StringNode(context.CustomShowId(show.Id)),
                ["name"] = StringNode(string.IsNullOrWhiteSpace(show.Name) ? "Custom show" : show.Name),
                ["pages"] = pages,
            };
            if (show.Source is { } source)
            {
                var capabilities = source.Editable
                    ? new[] { new CapabilitySpec("setName", ["name"]), new CapabilitySpec("setPages", ["pages"]) }
                    : [];
                item["nativeRef"] = NativeRef(
                    context,
                    $"customShow:{showIndex}",
                    HashOrFallback(source.ShowXmlSha256, show),
                    capabilities);
            }
            output.Add(item);
        }
        return output;
    }

    private static JsonArray ProjectComments(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.PageId(slide.Id);
            for (var commentIndex = 0; commentIndex < slide.LegacyComments.Count; commentIndex++)
            {
                var comment = slide.LegacyComments[commentIndex];
                if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Text)) continue;
                var capabilities = slide.Source?.LegacyCommentsEditable == true
                    ? new[] { new CapabilitySpec("replaceText", ["text"]) }
                    : [];
                var commentHash = HashOrFallback(null, comment);
                output.Add(new JsonObject
                {
                    ["id"] = StringNode(context.UniqueId($"comment-{pageId}-{comment.Id}")),
                    ["page"] = StringNode(pageId),
                    ["author"] = StringNode(string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author),
                    ["text"] = StringNode(comment.Text),
                    ["createdAt"] = StringNode(createdAt.ToUniversalTime().ToString("O")),
                    ["resolved"] = JsonValue.Create(false),
                    ["position"] = new JsonObject { ["x"] = JsonValue.Create(Points(comment.PositionXEmu)), ["y"] = JsonValue.Create(Points(comment.PositionYEmu)) },
                    ["nativeRef"] = NativeRef(context, $"comment:{pageId}:{commentIndex}", commentHash, capabilities),
                });
            }

            for (var threadIndex = 0; threadIndex < slide.ModernComments.Count; threadIndex++)
            {
                var thread = slide.ModernComments[threadIndex];
                if (thread.Root is null || thread.Anchor is null || thread.Anchor.Monikers.Count != 1)
                    continue;

                var sourceTargetId = thread.TargetId;
                if (sourceTargetId.EndsWith("/text", StringComparison.Ordinal))
                    sourceTargetId = sourceTargetId[..^5];
                if (!context.TryElementId(pageId, sourceTargetId, out var targetId))
                    continue;

                var rootId = context.UniqueId($"comment-{pageId}-modern-{threadIndex}-{thread.Root.Id}");
                ProjectModernComment(
                    output,
                    context,
                    pageId,
                    targetId,
                    thread,
                    thread.Root,
                    rootId,
                    parentId: null,
                    threadIndex,
                    replyIndex: null,
                    includeAnchor: true,
                    includePosition: true);
                for (var replyIndex = 0; replyIndex < thread.Replies.Count; replyIndex++)
                {
                    var reply = thread.Replies[replyIndex];
                    ProjectModernComment(
                        output,
                        context,
                        pageId,
                        targetId,
                        thread,
                        reply,
                        context.UniqueId($"comment-{pageId}-modern-{threadIndex}-reply-{replyIndex}-{reply.Id}"),
                        rootId,
                        threadIndex,
                        replyIndex,
                        includeAnchor: false,
                        includePosition: false);
                }
            }
        }
        return output;
    }

    private static void ProjectModernComment(
        JsonArray output,
        ProjectionContext context,
        string pageId,
        string targetId,
        PresentationModernCommentThread thread,
        PresentationModernComment comment,
        string commentId,
        string? parentId,
        int threadIndex,
        int? replyIndex,
        bool includeAnchor,
        bool includePosition)
    {
        if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Status))
            return;

        var capabilities = thread.Source?.Editable == true
            ? new[]
            {
                new CapabilitySpec("replaceText", ["text"]),
                new CapabilitySpec("setCommentStatus", ["status", "resolved"]),
            }
            : [];
        var item = new JsonObject
        {
            ["id"] = StringNode(commentId),
            ["page"] = StringNode(pageId),
            ["kind"] = StringNode("modern"),
            ["author"] = StringNode(string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author),
            ["text"] = StringNode(comment.Text),
            ["createdAt"] = StringNode(createdAt.ToUniversalTime().ToString("O")),
            ["resolved"] = JsonValue.Create(!comment.Status.Equals("active", StringComparison.Ordinal)),
            ["status"] = StringNode(comment.Status),
            ["nativeRef"] = NativeRef(
                context,
                replyIndex is { } index
                    ? $"comment:{pageId}:modern:{threadIndex}:reply:{index}"
                    : $"comment:{pageId}:modern:{threadIndex}:root",
                HashOrFallback(null, comment),
                capabilities),
        };
        if (parentId is not null) item["parent"] = StringNode(parentId);
        else item["target"] = StringNode(targetId);
        if (includePosition)
        {
            item["position"] = new JsonObject
            {
                ["x"] = JsonValue.Create(Points(thread.PositionXEmu)),
                ["y"] = JsonValue.Create(Points(thread.PositionYEmu)),
            };
        }
        if (includeAnchor)
        {
            var anchor = new JsonObject
            {
                ["kind"] = StringNode(thread.Anchor.Kind == PresentationModernCommentAnchor.Types.Kind.TextRange ? "textRange" : "element"),
                ["moniker"] = StringNode(thread.Anchor.Monikers[0].Type),
            };
            if (thread.Anchor.HasTextStart) anchor["textStart"] = JsonValue.Create(thread.Anchor.TextStart);
            if (thread.Anchor.HasTextLength) anchor["textLength"] = JsonValue.Create(thread.Anchor.TextLength);
            if (thread.Anchor.HasContextLength) anchor["contextLength"] = JsonValue.Create(thread.Anchor.ContextLength);
            if (thread.Anchor.HasContextHash) anchor["contextHash"] = JsonValue.Create(thread.Anchor.ContextHash);
            item["anchor"] = anchor;
        }
        output.Add(item);
    }

    private static IReadOnlyList<CapabilitySpec> Capabilities(
        PresentationElement element,
        bool hasEffectivePlaceholderFrame = false)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        // Accessibility is a separate non-visual cNvPr leaf.  Keep its
        // capability independent from the visual writer so a source-bound
        // edit can change title/description/decorative state without
        // widening the element's paint, geometry, or text profile.
        if (source.AccessibilityEditable)
            output.Add(new("setAccessibility", ["accessibility"]));
        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                if (source.TextEditable && TextTopologyRepresentable(element.Shape.TextBody))
                {
                    output.Add(new("replaceText", ["text"]));
                    // Slide placeholders use the text-only source-bound
                    // writer; it intentionally rejects paragraph/body-style
                    // changes because those may be inherited from the
                    // layout/master. Keep those capabilities on ordinary
                    // shape/text owners only.
                    if (element.Shape.Placeholder is null)
                    {
                        output.Add(new("setTextParagraphStyle", [
                            "text.paragraphs[].style.alignment",
                            "text.paragraphs[].style.tabStops",
                        ]));
                        if (PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(element.Shape.TextBody?.BodyProperties))
                        {
                            output.Add(new("setTextBodyStyle", [
                                element.Shape.Geometry is not ("textbox" or "none" or "")
                                    ? "textStyle"
                                    : "text.style",
                            ]));
                        }
                    }
                }
                if (element.Shape.Placeholder is not null &&
                    source.DirectFramePresenceEditable &&
                    (element.Shape.DirectFrame is not null || hasEffectivePlaceholderFrame))
                    // A complete direct transform is owner-local.  When the
                    // slide only inherits its geometry, the same fields are
                    // safe to materialize on the slide after the effective
                    // frame has been uniquely resolved from layout/master.
                    output.Add(new("setFrame", EditableFrameFields));
                if (source.Editable)
                {
                    // A recognized shape-tree click action is owned by the
                    // same cNvPr closure as the shape.  Advertise a narrow
                    // replacement capability; PptxHyperlinkCodec still
                    // rejects unknown sound/macro/extension children.
                    output.Add(new("setAction", ["action"]));
                    output.Add(new("setHoverAction", ["hoverAction"]));
                    if (element.Shape.Geometry == "line" || PpjLinePathCodec.IsLineLike(element.Shape))
                    {
                        output.Add(new("setLinePath", ["line.path"]));
                        output.Add(new("setStroke", ["stroke"]));
                        output.Add(new("setShapeEffects", ["shape.shadow", "shape.glow", "shape.innerShadow", "shape.reflection", "shape.softEdge"]));
                        output.Add(new("setFrame", EditableFrameFields));
                        break;
                    }
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or "" or "custom") &&
                        TryGetCompoundShapeOpacity(element.Shape, out _))
                        output.Add(new("setOpacity", ["compositing.opacity"]));
                    // Source-bound image-filled custom geometry is projected
                    // for discovery and frame/stroke edits, but its native
                    // fill graph is not represented by a lossless PPJ
                    // replacement operation. Do not advertise setFill for
                    // that bounded shape profile.
                    if (string.IsNullOrWhiteSpace(element.Shape.ImageFillAssetId) ||
                        element.Shape.Geometry != "custom")
                        output.Add(new("setFill", ["fill"]));
                    output.Add(new("setStroke", ["stroke"]));
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or ""))
                        output.Add(new("setShapeEffects", ["shape.shadow", "shape.glow", "shape.innerShadow", "shape.reflection", "shape.softEdge"]));
                    output.Add(new("setFrame", element.Shape.Placeholder is null ? EditableFrameFields : PositionFrameFields));
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or "custom") &&
                        PptxPresetGeometryAdjustmentCodec.TryExpectedCount(element.Shape.Geometry, out var adjustmentCount) &&
                        adjustmentCount > 0)
                        output.Add(new("setGeometry", ["geometry.adjustments"]));
                    else if (element.Shape.Placeholder is null &&
                             element.Shape.Geometry == "custom" &&
                             CanProjectCustomGeometry(element.Shape))
                        output.Add(new("setGeometry", ["geometry.paths"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Image when source.Editable:
                output.Add(new("replaceImage", ["image.asset"]));
                if (!string.IsNullOrEmpty(element.Image.SvgAssetId))
                    output.Add(new("replaceSvg", ["image.svgAsset"]));
                output.Add(new("setImageCrop", ["image.crop"]));
                output.Add(new("setImageFit", ["image.fit"]));
                output.Add(new("setFrame", EditableFrameFields));
                output.Add(new("setOpacity", ["opacity"]));
                output.Add(new("setImageEffects", ["image.border", "image.shadow", "image.glow", "image.innerShadow", "image.reflection", "image.softEdge"]));
                if (element.Image.CustomMaskPaths.Count == 0 ||
                    CanProjectCustomGeometry(ImageMaskShape(element.Image)))
                {
                    // Keep one capability per picture mask owner.  The
                    // compiler still gates each transition by topology and
                    // profile, while the field set covers preset identity,
                    // complete adjustment lists, and literal path changes.
                    output.Add(new("setImageMask", ["image.mask.preset", "image.mask.adjustments", "image.mask.paths"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Chart when source.Editable:
                output.Add(new("setChartTitle", ["chart.title"]));
                output.Add(new("setChartData", ["chart.data"]));
                output.Add(new("setChartTextStyle", ["chart.textStyle"]));
                output.Add(new("setChartFill", ["chart.fill", "chart.legendFill"]));
                // Direct series stroke and marker graphs are parsed by the
                // existing ChartSpace codecs.  Keep this capability separate
                // from fill/data so a source-bound edit can change only the
                // series style leaves and still fail closed by chart family
                // in the compiler (scatter uses marker.line, not series.line).
                output.Add(new("setChartSeriesStyle", [
                    "chart.data.series[].stroke",
                    "chart.data.series[].marker",
                ]));
                // Trendlines and error bars are direct c:series children with
                // bounded readers/writers of their own.  Keep them separate
                // from paint so a source-bound edit can replace only an
                // existing analytic child while topology changes still fail
                // closed in the shared ChartSpace patcher.
                output.Add(new("setChartSeriesAnalytics", [
                    "chart.data.series[].trendlines",
                    "chart.data.series[].errorBars",
                ]));
                if (element.Chart.Frame is not null)
                    output.Add(new("setChartFrame", ["chart.frame"]));
                output.Add(new("setChartLabels", ["chart.labels"]));
                if (element.Chart.XAxis is not null || element.Chart.YAxis is not null ||
                    element.Chart.SecondaryXAxis is not null || element.Chart.SecondaryYAxis is not null)
                    output.Add(new("setChartAxis", ["chart.axis"]));
                // Common ChartSpace plot scalars (legend placement, grouping,
                // gap width, axis visibility, line smoothness/colour
                // variation, and bounded circular/bubble geometry) are safe
                // only for the parser-owned chart families below.  The
                // capability is deliberately broad at the operation level;
                // the source-bound compiler still checks each field and the
                // chart family before writing the existing ChartPart.
                if (element.Chart.Type is SpreadsheetChartType.Bar or
                    SpreadsheetChartType.Line or
                    SpreadsheetChartType.Area or
                    SpreadsheetChartType.Pie or
                    SpreadsheetChartType.Doughnut or
                    SpreadsheetChartType.Bubble or
                    SpreadsheetChartType.Radar or
                    SpreadsheetChartType.Combo)
                    output.Add(new("setChartPlot", ["chart.plot"]));
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Table when source.Editable:
                output.Add(new("replaceText", ["text"]));
                output.Add(new("setTableStyle", ["table.style"]));
                // The native rectangular table profile keeps grid identity
                // and merge topology fixed, but its existing column widths
                // and row heights are safe scalar leaves. Expose them as a
                // separate capability so a geometry edit cannot be confused
                // with text reflow or a topology change.
                output.Add(new("setTableGeometry", ["table.geometry"]));
                if (element.Table.CellStyleEditable || element.Table.CellTextStyleEditable)
                {
                    var fields = new List<string>();
                    if (element.Table.CellStyleEditable)
                    {
                        fields.Add("table.cell.fill");
                        fields.Add("table.cell.borders");
                    }
                    if (element.Table.CellTextStyleEditable)
                        fields.Add("table.cell.textStyle");
                    output.Add(new("setTableCellStyle", fields));
                }
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Connector when source.Editable:
                output.Add(new("setStroke", ["stroke"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Group when source.Editable:
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Diagram:
                output.Add(new("setSmartArtText", ["smartArt.text"]));
                output.Add(new("setSmartArtGraph", ["smartArt.connections"]));
                if (HasSmartArtPicturePaintProfile(element.Diagram))
                {
                    output.Add(new("setSmartArtImage", ["smartArt.nodes[].asset"]));
                    output.Add(new("setSmartArtImagePaint", ["smartArt.nodes[].image"]));
                }
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                if (element.Diagram.DrawingCacheVerified)
                    output.Add(new("detachSmartArt", ["smartArt.detachToShapes"]));
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (element.Opaque.DiagramText is { Nodes.Count: > 0 })
                    output.Add(new("setSmartArtText", ["smartArt.text"]));
                if (element.Opaque.OleWorkbook is not null || element.Opaque.OleOfficePackage is not null)
                    output.Add(new("setOlePayload", ["ole.payload"]));
                if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
        }
        if (source.VisibilityEditable) output.Add(new("setHidden", ["hidden"]));
        if (source.LockingEditable) output.Add(new("setLocked", ["locked"]));
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static bool HasSmartArtPicturePaintProfile(PresentationDiagram diagram)
    {
        if (!diagram.DrawingCacheVerified || diagram.Nodes.Count == 0 ||
            diagram.Nodes.Any(node => string.IsNullOrWhiteSpace(node.AssetId)))
            return false;
        return diagram.Nodes.All(node => diagram.Drawing?.Children
            .SingleOrDefault(child => child.Name == node.Id && child.Shape is not null)
            ?.Shape?.ImageFill is { AssetId.Length: > 0 });
    }

    private static bool TextTopologyRepresentable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase is PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.LineBreak or PresentationTextRun.ContentOneofCase.Field));

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
        if (source.AccessibilityEditable) output.Add(new("setAccessibility", ["accessibility"]));
        // OLE and source-owned chart payloads stay opaque until the
        // corresponding PPJ typed state is projected. Proven diagram text is
        // handled by ProjectSourceSmartArt before this fallback is selected.
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
                ["id"] = StringNode($"cap-{capability.Operation}-{Sha256(Encoding.UTF8.GetBytes(scope + capability.Operation))[..10]}"),
                ["operation"] = StringNode(capability.Operation),
                ["expectedHash"] = StringNode(objectHash),
                ["fields"] = fields,
            });
        }
        var output = new JsonObject
        {
            ["handle"] = StringNode($"nr-{Sha256(Encoding.UTF8.GetBytes(context.SourceSha256 + "\0" + scope))}"),
            ["sourceSha256"] = StringNode(context.SourceSha256),
            ["revision"] = StringNode(context.Revision),
            ["objectHash"] = StringNode(objectHash),
            ["capabilitySetSha256"] = StringNode(Sha256(CanonicalBytes(capabilityArray))),
            ["capabilities"] = capabilityArray,
        };
        if (leaves is { Count: > 0 }) output["leaves"] = leaves;
        return output;
    }

    private static JsonObject ShapeFrame(PresentationShape shape)
    {
        var frame = Frame(shape.LeftEmu, shape.TopEmu, shape.WidthEmu, shape.HeightEmu);
        if (shape.Transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(shape.Transform.RotationAngle60000 / 60_000d);
        if (shape.Transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(shape.Transform.FlipHorizontal);
        if (shape.Transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(shape.Transform.FlipVertical);
        return frame;
    }

    private static JsonObject ShapeFrame(PresentationPlaceholderFrame source)
    {
        var frame = Frame(source.LeftEmu, source.TopEmu, source.WidthEmu, source.HeightEmu);
        if (source.HasRotationAngle60000)
            frame["rotation"] = JsonValue.Create(source.RotationAngle60000 / 60_000d);
        if (source.HasFlipHorizontal) frame["flipH"] = JsonValue.Create(source.FlipHorizontal);
        if (source.HasFlipVertical) frame["flipV"] = JsonValue.Create(source.FlipVertical);
        return frame;
    }

    private static readonly string[] EditableFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height", "frame.rotation", "frame.flipH", "frame.flipV"];

    private static readonly string[] PositionFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height"];

    private static JsonObject ImageFrame(PresentationImage image)
    {
        var frame = Frame(image.LeftEmu, image.TopEmu, image.WidthEmu, image.HeightEmu);
        if (image.Transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(image.Transform.RotationAngle60000 / 60_000d);
        if (image.Transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(image.Transform.FlipHorizontal);
        if (image.Transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(image.Transform.FlipVertical);
        return frame;
    }

    private static JsonObject TableFrame(PresentationTable table) =>
        Frame(table.LeftEmu, table.TopEmu, table.WidthEmu, table.HeightEmu, table.FrameTransform);

    private static JsonObject ChartFrame(PresentationChart chart) =>
        Frame(chart.LeftEmu, chart.TopEmu, chart.WidthEmu, chart.HeightEmu, chart.FrameTransform);

    private static JsonObject GroupFrame(PresentationGroup group) =>
        Frame(group.LeftEmu, group.TopEmu, group.WidthEmu, group.HeightEmu, group.FrameTransform);

    private static JsonObject DiagramFrame(PresentationDiagram diagram) =>
        Frame(diagram.LeftEmu, diagram.TopEmu, diagram.WidthEmu, diagram.HeightEmu);

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
        PresentationElement.ContentOneofCase.Diagram => DiagramFrame(element.Diagram),
        PresentationElement.ContentOneofCase.Group => GroupFrame(element.Group),
        PresentationElement.ContentOneofCase.Opaque => Frame(element.Opaque.LeftEmu, element.Opaque.TopEmu, element.Opaque.WidthEmu, element.Opaque.HeightEmu),
        _ => Frame(0, 0, 1, 1),
    };

    private static JsonObject Frame(long left, long top, long width, long height) => new()
    {
        ["x"] = JsonValue.Create(Points(left)),
        ["y"] = JsonValue.Create(Points(top)),
        ["width"] = JsonValue.Create(Math.Max(0.001, Points(width))),
        ["height"] = JsonValue.Create(Math.Max(0.001, Points(height))),
    };

    private static JsonObject Frame(
        long left,
        long top,
        long width,
        long height,
        PresentationFrameTransform? transform)
    {
        var frame = Frame(left, top, width, height);
        if (transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(transform.RotationAngle60000 / 60_000d);
        if (transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(transform.FlipHorizontal);
        if (transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(transform.FlipVertical);
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
            return new JsonObject { ["element"] = StringNode(projected), ["anchor"] = StringNode("auto") };
        return new JsonObject { ["x"] = JsonValue.Create(Points(x)), ["y"] = JsonValue.Create(Points(y)) };
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
                ? new JsonObject { ["token"] = StringNode(scheme) }
                : StringNode(Color(string.IsNullOrEmpty(rgb) ? "000000" : rgb)),
                ["width"] = JsonValue.Create(Math.Max(0, Points(widthEmu))),
        };
        var dash = Dash(style);
        if (dash is not null) output["dash"] = StringNode(dash);
        if (cap is "flat" or "round" or "square") output["cap"] = StringNode(cap);
        if (join is "miter" or "round" or "bevel") output["join"] = StringNode(join);
        if (opacity is not null) output["opacity"] = JsonValue.Create(opacity.Value);
        return output;
    }

    private static JsonObject Shadow(PresentationShadow shadow, bool includeOpacity = true)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(shadow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(shadow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(shadow.ColorRgb) ? "000000" : shadow.ColorRgb)),
            ["blur"] = JsonValue.Create(Math.Max(0, Points(shadow.HasBlurRadiusEmu ? shadow.BlurRadiusEmu : 0))),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(shadow.HasDistanceEmu ? shadow.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((shadow.HasDirectionAngle60000 ? shadow.DirectionAngle60000 : 0) / 60_000d),
        };
        if (includeOpacity)
            output["opacity"] = JsonValue.Create(shadow.HasOpacityThousandthPercent ? Unit(shadow.OpacityThousandthPercent) : 1);
        if (shadow.HasAlignment) output["alignment"] = StringNode(shadow.Alignment);
        if (shadow.HasRotateWithShape) output["rotateWithShape"] = JsonValue.Create(shadow.RotateWithShape);
        return output;
    }

    private static JsonObject Glow(PresentationGlow glow)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(glow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(glow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(glow.ColorRgb) ? "000000" : glow.ColorRgb)),
            ["radius"] = JsonValue.Create(Math.Max(0, Points(glow.HasRadiusEmu ? glow.RadiusEmu : 0))),
        };
        if (glow.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(glow.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject InnerShadow(PresentationInnerShadow shadow)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(shadow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(shadow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(shadow.ColorRgb) ? "000000" : shadow.ColorRgb)),
            ["blur"] = JsonValue.Create(Math.Max(0, Points(shadow.HasBlurRadiusEmu ? shadow.BlurRadiusEmu : 0))),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(shadow.HasDistanceEmu ? shadow.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((shadow.HasDirectionAngle60000 ? shadow.DirectionAngle60000 : 0) / 60_000d),
        };
        if (shadow.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(shadow.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject Reflection(PresentationReflection reflection)
    {
        var output = new JsonObject
        {
            ["blur"] = JsonValue.Create(Math.Max(0, Points(reflection.HasBlurRadiusEmu ? reflection.BlurRadiusEmu : 0))),
            ["startOpacity"] = JsonValue.Create(Unit(reflection.HasStartOpacityThousandthPercent ? reflection.StartOpacityThousandthPercent : 100_000)),
            ["endOpacity"] = JsonValue.Create(Unit(reflection.HasEndOpacityThousandthPercent ? reflection.EndOpacityThousandthPercent : 0)),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(reflection.HasDistanceEmu ? reflection.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((reflection.HasDirectionAngle60000 ? reflection.DirectionAngle60000 : 0) / 60_000d),
        };
        return output;
    }

    private static JsonObject SoftEdge(PresentationSoftEdge softEdge) => new()
    {
        ["radius"] = JsonValue.Create(Math.Max(0, Points(softEdge.HasRadiusEmu ? softEdge.RadiusEmu : 0))),
    };

    private static JsonObject? Accessibility(PresentationNonVisualAccessibility? value)
    {
        if (value is null) return null;
        var output = new JsonObject { ["decorative"] = JsonValue.Create(value.HasDecorative && value.Decorative) };
        if (!string.IsNullOrEmpty(value.Title)) output["title"] = StringNode(value.Title);
        if (!string.IsNullOrEmpty(value.Description)) output["description"] = StringNode(value.Description);
        return output;
    }

    private static JsonObject? ImageAccessibility(PresentationImage image)
    {
        if (!image.HasAccessibilityDecorative && string.IsNullOrEmpty(image.AccessibilityTitle) && string.IsNullOrEmpty(image.AltText)) return null;
        var output = new JsonObject { ["decorative"] = JsonValue.Create(image.HasAccessibilityDecorative && image.AccessibilityDecorative) };
        if (!string.IsNullOrEmpty(image.AccessibilityTitle)) output["title"] = StringNode(image.AccessibilityTitle);
        if (!string.IsNullOrEmpty(image.AltText)) output["description"] = StringNode(image.AltText);
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
            if (!hasOpacity) return StringNode(value);
            var alpha = Math.Clamp((int)Math.Round(Unit(opacity) * 255), 0, 255);
            return StringNode($"{value}{alpha:X2}");
        }
        var output = new JsonObject { ["token"] = StringNode(scheme ?? string.Empty) };
        if (hasOpacity) output["alpha"] = JsonValue.Create(Unit(opacity));
        return output;
    }
    private static double Crop(int value) => Math.Clamp(value / 100_000d, -1, 1);
    private static double Unit(uint value) => Math.Clamp(value / 100_000d, 0, 1);
    private static double Points(long emu) => Math.Round(emu / EmuPerPoint, 6, MidpointRounding.AwayFromZero);

    private static string HashOrFallback(string? hash, IMessage fallback) =>
        IsCanonicalSha256(hash) ? hash! : Sha256(fallback.ToByteArray());

    private static string HashOrFallback(string? hash, ByteString fallback) =>
        IsCanonicalSha256(hash) ? hash! : Sha256(fallback.Span);

    private static bool IsCanonicalSha256(string? hash) =>
        hash is { Length: 64 } && hash.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

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
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            node.WriteTo(writer);
        using var document = JsonDocument.Parse(buffer.WrittenMemory);
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

    private static string Sha256(byte[] bytes) => Sha256(bytes.AsSpan());
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    [GeneratedRegex("[^A-Za-z0-9._:-]+")]
    private static partial Regex InvalidIdCharacters();

    private sealed record CapabilitySpec(string Operation, IReadOnlyList<string> Fields);

    private sealed class ProjectionContext
    {
        private readonly IReadOnlyDictionary<string, Asset> sourceAssets;
        private readonly IReadOnlyDictionary<string, OpaqueOpcPart> sourceParts;
        private readonly PptxPackageSource sourcePackage;
        private readonly Dictionary<string, string> pageIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> masterIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> layoutIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> customShowIds = new(StringComparer.Ordinal);
        private readonly Dictionary<(string Page, string Element), string> elementIds = new();
        private readonly HashSet<string> usedIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdBySourceId = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdByHash = new(StringComparer.Ordinal);
        private readonly JsonArray programAssets = new();
        private readonly List<Asset> resultAssets = [];
        private readonly List<Asset> nativeSourceAssets = [];
        private readonly JsonArray nodes = new();
        private readonly Dictionary<string, PpjNativeLeafBinding> nativeLeafBindings = new(StringComparer.Ordinal);

        internal ProjectionContext(
            string sourceSha256,
            string revision,
            string assetRoot,
            IEnumerable<Asset> assets,
            OpaqueOpcGraph? opaque,
            PptxPackageSource sourcePackage)
        {
            SourceSha256 = sourceSha256;
            Revision = revision;
            AssetRoot = assetRoot;
            sourceAssets = assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
            sourceParts = (opaque?.Parts ?? []).ToDictionary(part => part.Path, StringComparer.OrdinalIgnoreCase);
            this.sourcePackage = sourcePackage;
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
        internal IReadOnlyList<Asset> NativeSourceAssets => nativeSourceAssets;

        internal void ReleaseProjectionJson()
        {
            programAssets.Clear();
            nodes.Clear();
        }

        internal string RegisterPage(string sourceId, string? stableSourceId)
        {
            var id = UniqueId($"page-{NormalizeId(stableSourceId, NormalizeId(sourceId, "slide"))}");
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

        internal string RegisterCustomShow(string sourceId)
        {
            var id = UniqueId($"show-{NormalizeId(sourceId, "show")}");
            customShowIds[sourceId] = id;
            return id;
        }

        internal string RegisterElement(string pageId, string sourceId)
        {
            var id = UniqueId($"{pageId}-{NormalizeId(PageLocalElementPath(sourceId), "element")}");
            elementIds[(pageId, sourceId)] = id;
            return id;
        }

        private static string PageLocalElementPath(string sourceId)
        {
            const string marker = "/element/";
            var index = sourceId.IndexOf(marker, StringComparison.Ordinal);
            return index < 0 ? sourceId : sourceId[(index + 1)..];
        }

        internal string PageId(string sourceId) => pageIds[sourceId];
        internal bool TryPageId(string sourceId, out string id) => pageIds.TryGetValue(sourceId, out id!);
        internal string MasterId(string sourceId) => masterIds[sourceId];
        internal bool TryLayoutId(string sourceId, out string id) => layoutIds.TryGetValue(sourceId, out id!);
        internal string LayoutId(string sourceId) => layoutIds[sourceId];
        internal string CustomShowId(string sourceId) => customShowIds[sourceId];
        internal bool TryCustomShowId(string sourceId, out string id) => customShowIds.TryGetValue(sourceId, out id!);
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
            if (!sourceAssets.TryGetValue(sourceId, out var source)) return false;
            var hash = HashOrFallback(source.Sha256, source.Data);
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
                ["id"] = StringNode(programAssetId),
                ["uri"] = StringNode($"{AssetRoot}/{fileName}"),
                ["mimeType"] = StringNode(source.ContentType),
                ["sha256"] = StringNode(hash),
                ["rights"] = new JsonObject { ["status"] = StringNode("user-provided") },
                ["accessibility"] = new JsonObject { ["decorative"] = JsonValue.Create(false), ["description"] = StringNode("Imported source media.") },
            });
            var materialized = source.Clone();
            materialized.Id = programAssetId;
            materialized.FileName = fileName;
            materialized.Sha256 = hash;
            resultAssets.Add(materialized);
            return true;
        }

        internal bool TryMaterializeSourcePart(string partPath, string contentType, out string programAssetId)
        {
            var sourceId = $"source-part:{partPath}";
            if (assetIdBySourceId.TryGetValue(sourceId, out programAssetId!)) return true;
            if (!sourceParts.TryGetValue(partPath, out var metadata) ||
                !metadata.ContentType.Equals(contentType, StringComparison.OrdinalIgnoreCase))
            {
                programAssetId = string.Empty;
                return false;
            }

            byte[] data;
            using (var stream = sourcePackage.OpenRead())
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
            {
                var entry = archive.Entries.SingleOrDefault(candidate => candidate.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));
                if (entry is null || entry.Length is <= 0 or > 16 * 1024 * 1024)
                {
                    programAssetId = string.Empty;
                    return false;
                }
                using var entryStream = entry.Open();
                using var memory = new MemoryStream();
                entryStream.CopyTo(memory);
                data = memory.ToArray();
            }
            var hash = Sha256(data);
            if (!hash.Equals(metadata.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                programAssetId = string.Empty;
                return false;
            }

            if (!assetIdByHash.TryGetValue(hash, out programAssetId!))
            {
                programAssetId = UniqueId($"asset-{hash[..20]}");
                assetIdByHash[hash] = programAssetId;
                var extension = Extension(contentType, Path.GetFileName(partPath));
                var fileName = $"{hash}.{extension}";
                programAssets.Add(new JsonObject
                {
                    ["id"] = StringNode(programAssetId),
                    ["uri"] = StringNode($"{AssetRoot}/{fileName}"),
                    ["mimeType"] = StringNode(contentType),
                    ["sha256"] = StringNode(hash),
                    ["rights"] = new JsonObject { ["status"] = StringNode("user-provided") },
                    ["accessibility"] = new JsonObject { ["decorative"] = JsonValue.Create(false), ["description"] = StringNode("Imported embedded Office package.") },
                });
                resultAssets.Add(new Asset
                {
                    Id = programAssetId,
                    FileName = fileName,
                    ContentType = contentType,
                    Data = ByteString.CopyFrom(data),
                    Sha256 = hash,
                });
            }
            assetIdBySourceId[sourceId] = programAssetId;
            var nativeId = PptxAssetCatalog.NativeAssetIdFor(contentType, hash);
            if (!nativeSourceAssets.Any(asset => asset.Id.Equals(nativeId, StringComparison.Ordinal)))
                nativeSourceAssets.Add(new Asset
                {
                    Id = nativeId,
                    FileName = Path.GetFileName(partPath),
                    ContentType = contentType,
                    Data = ByteString.CopyFrom(data),
                    Sha256 = hash,
                });
            return true;
        }

        internal JsonObject? SmartArtNativeSections(IEnumerable<string> partPaths)
        {
            var sections = new SortedDictionary<string, string>(StringComparer.Ordinal);
            foreach (var partPath in partPaths)
            {
                if (!sourceParts.TryGetValue(partPath, out var part) || string.IsNullOrWhiteSpace(part.Sha256)) continue;
                var name = part.ContentType switch
                {
                    var value when value.Contains("diagramData", StringComparison.OrdinalIgnoreCase) => "dataSha256",
                    var value when value.Contains("diagramLayout", StringComparison.OrdinalIgnoreCase) => "layoutSha256",
                    var value when value.Contains("diagramStyle", StringComparison.OrdinalIgnoreCase) => "styleSha256",
                    var value when value.Contains("diagramColors", StringComparison.OrdinalIgnoreCase) => "colorsSha256",
                    var value when value.Contains("diagramDrawing", StringComparison.OrdinalIgnoreCase) => "drawingSha256",
                    _ => string.Empty,
                };
                if (name.Length > 0) sections.TryAdd(name, part.Sha256.ToLowerInvariant());
            }
            if (sections.Count == 0) return null;
            var output = new JsonObject();
            foreach (var (name, hash) in sections) output[name] = StringNode(hash);
            output["closureSha256"] = StringNode(Sha256(Encoding.UTF8.GetBytes(string.Join(
                "\n",
                sections.Select(pair => $"{pair.Key}:{pair.Value}")))));
            return output;
        }

        internal void RecordNode(string pageId, string id, string type, JsonObject nativeRef)
        {
            VisibleObjectCount++;
            nodes.Add(new JsonObject
            {
                ["id"] = StringNode(id),
                ["page"] = StringNode(pageId),
                ["type"] = StringNode(type),
                ["handle"] = StringNode(nativeRef["handle"]!.GetValue<string>()),
                ["objectHash"] = StringNode(nativeRef["objectHash"]!.GetValue<string>()),
                ["capabilitySetSha256"] = StringNode(nativeRef["capabilitySetSha256"]!.GetValue<string>()),
            });
        }

        internal JsonObject BuildNodeMap() => new()
        {
            ["schema"] = StringNode("office-kit/ppj-node-map/v1"),
            ["sourceSha256"] = StringNode(SourceSha256),
            ["revision"] = StringNode(Revision),
            ["nodes"] = nodes,
        };

        private static string Extension(string contentType, string fileName)
        {
            var fromType = contentType.ToLowerInvariant() switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/svg+xml" => "svg",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => "xlsx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => "docx",
                _ => string.Empty,
            };
            if (fromType.Length > 0) return fromType;
            var extension = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            return Regex.IsMatch(extension, "^[a-z0-9]{1,8}$") ? extension : "bin";
        }
    }
}
