using System.Buffers;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjCompileResult(
    byte[] File,
    PresentationProgramResult Program,
    IReadOnlyList<Diagnostic> Diagnostics,
    bool ReuseSourceFile = false);

/// <summary>
/// Lowers one validated, source-free PPJ program directly into the native C#
/// Presentation compiler IR. JavaScript never materializes a Presentation
/// object model on this path; the protobuf model is an internal writer IR.
/// </summary>
internal static partial class PpjAuthoredPresentationCompiler
{
    private const double EmuPerPoint = 12_700d;
    private const double CustomPathUnitsPerPoint = 1_000d;

    internal sealed record TextPrecedenceContext(
        JsonElement? LayoutStyle,
        JsonElement? MasterStyle,
        JsonElement? ElementStyle,
        JsonElement? ParagraphStyle);

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
        if (request.IncludePreviewScene && program.Components.Count != 0 && validation.Expansion!.PreviewOrigins is null)
            throw new CodecException("ppj.preview.originsRequired",
                "Component preview requires provenance captured during the original validation/expansion pass.", "$.components");
        // Ordinary builds keep the original plan and its per-slide lifetime.
        var preview = request.IncludePreviewScene
            ? new PpjPreviewSourceFreeBuildPlan(plan, validation.Expansion.Nodes, limits, validation.Expansion.PreviewOrigins)
            : null;
        var originalProgramJson = request.ProgramJson.ToByteArray();
        PptxExportResult exported;
        using (PpjBuildProfiler.Measure("writer"))
        {
            exported = PptxCodec.ExportSourceFree(
                (IPptxSourceFreeBuildPlan?)preview ?? plan,
                assets,
                limits,
                parts => PpjEmbeddedProgramCodec.AddToSourceFreePackage(
                    parts,
                    originalProgramJson,
                    validation,
                    plan.NativeBindings,
                    assets));
        }
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
        if (preview is not null)
            receipt.PreviewScene = preview.Complete(validation.ProgramSha256, fileSha256, assets);
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

    // Source-bound table cells use the same concrete text grammar as
    // source-free PPJ, but deliberately call this helper without style
    // overlays. The caller has already proved that the imported body is a
    // fixed-topology, direct plain-text profile; this keeps per-run styles
    // editable without making source-bound compilation infer unsupported
    // paragraph/list/layout semantics.
    internal static PresentationTextBody BuildSourceBoundTextBody(JsonElement text, JsonElement programRoot)
    {
        var catalog = new Catalog(programRoot);
        var body = BuildTextBody(text, null, null, null, catalog);
        // BuildTextBody allocates a body-properties shell while resolving
        // authored styles. Source projection omits an empty shell, so clear
        // it here to keep source-bound semantic hashes stable.
        if (!PptxBodyPropertiesCodec.HasModeledProperties(body.BodyProperties)) body.BodyProperties = null;
        return body;
    }

    internal static PresentationTextBodyProperties BuildSourceBoundTextBodyStyle(
        JsonElement style,
        string path)
    {
        if (style.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "source-bound text body style must be an object");
        foreach (var property in style.EnumerateObject())
        {
            if (property.Name is not ("verticalAlignment" or "anchorCenter" or "forceAntiAlias" or "spaceFirstLastParagraph" or "compatibleLineSpacing" or "fromWordArt" or "textWarpPreset" or "textWarpAdjustments" or "flatTextZ" or "wrap" or "autoFit" or "normalAutoFit" or "margins" or "columns" or "columnGap" or "columnDirection" or "verticalText" or "rotation" or "verticalOverflow" or "horizontalOverflow" or "upright"))
                throw Unsupported(path + "." + property.Name, "source-bound text body style field is outside the direct bodyPr profile");
        }
        var target = new PresentationTextBodyProperties();
        ApplyBodyProperties(target, null, null, style);
        if (!PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(target))
            throw Unsupported(path, "source-bound text body style is outside the bounded direct bodyPr profile");
        return target;
    }

    internal static void MergeSourceBoundTextBodyStyle(
        PresentationTextBody target,
        JsonElement style,
        string path,
        JsonElement? previousStyle)
    {
        if (target is null || target.Paragraphs.Count == 0)
            throw Unsupported(path, "source-bound text body style requires an existing text body");
        var requested = BuildSourceBoundTextBodyStyle(style, path);
        if (!PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(target.BodyProperties))
            throw Unsupported(path, "the imported text body has unmodeled bodyPr properties");
        var current = target.BodyProperties ?? new PresentationTextBodyProperties();
        if (style.TryGetProperty("verticalAlignment", out _))
        {
            current.ClearAnchor();
            if (requested.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor)
                current.VerticalAnchor = requested.VerticalAnchor;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousAlignment && previousAlignment.TryGetProperty("verticalAlignment", out _))
            current.NoVerticalAnchor = true;
        if (style.TryGetProperty("anchorCenter", out _))
        {
            current.ClearNoAnchorCenter();
            current.AnchorCenter = requested.AnchorCenter;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousCenter && previousCenter.TryGetProperty("anchorCenter", out _))
        {
            current.ClearAnchorCenter();
            current.NoAnchorCenter = true;
        }
        if (style.TryGetProperty("forceAntiAlias", out _))
        {
            current.ClearNoForceAntiAlias();
            current.ForceAntiAlias = requested.ForceAntiAlias;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousAntiAlias && previousAntiAlias.TryGetProperty("forceAntiAlias", out _))
        {
            current.ClearForceAntiAlias();
            current.NoForceAntiAlias = true;
        }
        if (style.TryGetProperty("spaceFirstLastParagraph", out _))
        {
            current.ClearNoSpaceFirstLastParagraph();
            current.SpaceFirstLastParagraph = requested.SpaceFirstLastParagraph;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousParagraphSpacing && previousParagraphSpacing.TryGetProperty("spaceFirstLastParagraph", out _))
        {
            current.ClearSpaceFirstLastParagraph();
            current.NoSpaceFirstLastParagraph = true;
        }
        if (style.TryGetProperty("compatibleLineSpacing", out _))
        {
            current.ClearNoCompatibleLineSpacing();
            current.CompatibleLineSpacing = requested.CompatibleLineSpacing;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousCompatibleSpacing && previousCompatibleSpacing.TryGetProperty("compatibleLineSpacing", out _))
        {
            current.ClearCompatibleLineSpacing();
            current.NoCompatibleLineSpacing = true;
        }
        if (style.TryGetProperty("fromWordArt", out _))
        {
            current.ClearNoFromWordArt();
            current.FromWordArt = requested.FromWordArt;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousWordArt && previousWordArt.TryGetProperty("fromWordArt", out _))
        {
            current.ClearFromWordArt();
            current.NoFromWordArt = true;
        }
        if (style.TryGetProperty("textWarpPreset", out _))
        {
            current.ClearNoTextWarpPreset();
            current.TextWarpPreset = requested.TextWarpPreset;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousWarp && previousWarp.TryGetProperty("textWarpPreset", out _))
        {
            current.ClearTextWarpPreset();
            current.TextWarpAdjustments.Clear();
            current.NoTextWarpPreset = true;
        }
        if (style.TryGetProperty("textWarpAdjustments", out _))
        {
            current.TextWarpAdjustments.Clear();
            current.TextWarpAdjustments.Add(requested.TextWarpAdjustments);
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousGuides && previousGuides.TryGetProperty("textWarpAdjustments", out _))
            current.TextWarpAdjustments.Clear();
        if (style.TryGetProperty("flatTextZ", out _))
        {
            current.ClearNoFlatTextZ();
            current.FlatTextZ = requested.FlatTextZ;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousFlatText && previousFlatText.TryGetProperty("flatTextZ", out _))
        {
            current.ClearFlatTextZ();
            current.NoFlatTextZ = true;
        }
        if (style.TryGetProperty("wrap", out _))
        {
            current.ClearWrapping();
            if (requested.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
                current.Wrap = requested.Wrap;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousWrap && previousWrap.TryGetProperty("wrap", out _))
            current.NoWrap = true;
        if (style.TryGetProperty("autoFit", out _))
        {
            current.ClearAutoFit();
            if (requested.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode)
                current.AutoFitMode = requested.AutoFitMode;
            if (current.AutoFitMode != "shrinkText") current.NormalAutoFit = null;
        }
        if (style.TryGetProperty("normalAutoFit", out var normalAutoFit))
        {
            if (current.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode || current.AutoFitMode != "shrinkText" ||
                requested.NormalAutoFit is not { } requestedNormal)
                throw Unsupported(path + ".normalAutoFit", "normalAutoFit percentages require autoFit shrink-text");
            var mergedNormal = current.NormalAutoFit?.Clone() ?? new PresentationNormalAutoFit();
            if (normalAutoFit.TryGetProperty("fontScale", out _))
            {
                mergedNormal.ClearFontScale();
                if (requestedNormal.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000)
                    mergedNormal.FontScale1000 = requestedNormal.FontScale1000;
            }
            if (normalAutoFit.TryGetProperty("lineSpacingReduction", out _))
            {
                mergedNormal.ClearLineSpacingReduction();
                if (requestedNormal.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000)
                    mergedNormal.LineSpacingReduction1000 = requestedNormal.LineSpacingReduction1000;
            }
            current.NormalAutoFit = mergedNormal;
        }
        if (style.TryGetProperty("margins", out var margins))
        {
            if (margins.TryGetProperty("left", out _))
            {
                current.ClearLeftInset();
                if (requested.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu)
                    current.LeftInsetEmu = requested.LeftInsetEmu;
            }
            if (margins.TryGetProperty("top", out _))
            {
                current.ClearTopInset();
                if (requested.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu)
                    current.TopInsetEmu = requested.TopInsetEmu;
            }
            if (margins.TryGetProperty("right", out _))
            {
                current.ClearRightInset();
                if (requested.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu)
                    current.RightInsetEmu = requested.RightInsetEmu;
            }
            if (margins.TryGetProperty("bottom", out _))
            {
                current.ClearBottomInset();
                if (requested.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu)
                    current.BottomInsetEmu = requested.BottomInsetEmu;
            }
        }
        if (style.TryGetProperty("columns", out _))
        {
            current.ClearColumnCount();
            if (requested.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
                current.Columns = requested.Columns;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousColumns && previousColumns.TryGetProperty("columns", out _))
            current.NoColumns = true;
        if (style.TryGetProperty("columnGap", out _))
        {
            current.ClearColumnSpacing();
            if (requested.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu)
                current.ColumnSpacingEmu = requested.ColumnSpacingEmu;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousGap && previousGap.TryGetProperty("columnGap", out _))
            current.NoColumnSpacing = true;
        if (style.TryGetProperty("columnDirection", out _))
        {
            current.ClearColumnDirection();
            if (requested.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns)
                current.RightToLeftColumns = requested.RightToLeftColumns;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousDirection && previousDirection.TryGetProperty("columnDirection", out _))
            current.NoColumnDirection = true;
        if (style.TryGetProperty("verticalText", out _))
        {
            current.ClearVerticalText();
            if (requested.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode)
                current.VerticalTextMode = requested.VerticalTextMode;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousVertical && previousVertical.TryGetProperty("verticalText", out _))
            current.NoVerticalTextMode = true;
        if (style.TryGetProperty("rotation", out _))
        {
            current.ClearRotation();
            if (requested.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000)
                current.RotationAngle60000 = requested.RotationAngle60000;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousRotation && previousRotation.TryGetProperty("rotation", out _))
            current.NoRotation = true;
        if (style.TryGetProperty("verticalOverflow", out _))
        {
            current.ClearVerticalOverflow();
            if (requested.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode)
                current.VerticalOverflowMode = requested.VerticalOverflowMode;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousVerticalOverflow && previousVerticalOverflow.TryGetProperty("verticalOverflow", out _))
            current.NoVerticalOverflowMode = true;
        if (style.TryGetProperty("horizontalOverflow", out _))
        {
            current.ClearHorizontalOverflow();
            if (requested.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode)
                current.HorizontalOverflowMode = requested.HorizontalOverflowMode;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previousHorizontalOverflow && previousHorizontalOverflow.TryGetProperty("horizontalOverflow", out _))
            current.NoHorizontalOverflowMode = true;
        if (style.TryGetProperty("upright", out _))
        {
            current.ClearUprightText();
            if (requested.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright)
                current.Upright = requested.Upright;
        }
        else if (previousStyle is { ValueKind: JsonValueKind.Object } previous && previous.TryGetProperty("upright", out _))
            current.NoUpright = true;
        ApplySourceBoundMarginRemoval(current, style, previousStyle);
        ApplySourceBoundAutoFitRemoval(current, style, previousStyle);
        target.BodyProperties = PptxBodyPropertiesCodec.HasModeledProperties(current) ? current : null;
    }

    internal static void ApplySourceBoundAutoFitRemoval(
        PresentationTextBodyProperties target, JsonElement? style, JsonElement? previousStyle)
    {
        if (previousStyle is not { ValueKind: JsonValueKind.Object } previous) return;
        var next = style ?? default;
        if (previous.TryGetProperty("autoFit", out _) &&
            (next.ValueKind != JsonValueKind.Object || !next.TryGetProperty("autoFit", out _)))
        {
            target.NoAutoFitMode = true;
            target.NormalAutoFit = null;
            return;
        }
        if (target.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode ||
            target.AutoFitMode != "shrinkText" || !previous.TryGetProperty("normalAutoFit", out var oldNormal)) return;
        var nextNormal = next.ValueKind == JsonValueKind.Object && next.TryGetProperty("normalAutoFit", out var normal)
            ? normal : default;
        bool Removed(string field) => oldNormal.TryGetProperty(field, out _) &&
            (nextNormal.ValueKind != JsonValueKind.Object || !nextNormal.TryGetProperty(field, out _));
        if (Removed("fontScale"))
        {
            target.NormalAutoFit ??= new PresentationNormalAutoFit();
            target.NormalAutoFit.NoFontScale = true;
        }
        if (Removed("lineSpacingReduction"))
        {
            target.NormalAutoFit ??= new PresentationNormalAutoFit();
            target.NormalAutoFit.NoLineSpacingReduction = true;
        }
    }

    internal static void ApplySourceBoundMarginRemoval(
        PresentationTextBodyProperties target, JsonElement? style, JsonElement? previousStyle)
    {
        if (previousStyle is not { ValueKind: JsonValueKind.Object } previous ||
            !previous.TryGetProperty("margins", out var previousMargins)) return;
        var nextMargins = style is { ValueKind: JsonValueKind.Object } next && next.TryGetProperty("margins", out var margins)
            ? margins : default;
        bool Removed(string edge) => previousMargins.TryGetProperty(edge, out _) &&
            (nextMargins.ValueKind != JsonValueKind.Object || !nextMargins.TryGetProperty(edge, out _));
        if (Removed("left")) target.NoLeftInset = true;
        if (Removed("top")) target.NoTopInset = true;
        if (Removed("right")) target.NoRightInset = true;
        if (Removed("bottom")) target.NoBottomInset = true;
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
            AddComments(Presentation, program, _expandedByPage);
        }

        public PresentationArtifact Presentation { get; }
        internal IReadOnlyList<PptxNativeBinding> NativeBindings => _nativeBindings;

        public bool RequiresPreviousSlide(int pageIndex) =>
            pageIndex >= 0 && pageIndex < _program.Pages.Count &&
            _program.Pages[pageIndex].Transition?.Type == "morph";

        public PresentationSlide MaterializeSlide(int pageIndex, PresentationSlide? previousSlide)
        {
            var page = _program.Pages[pageIndex];
            var expanded = _expandedByPage[page.Id];
            var textPrecedence = _catalog.TextPrecedenceForLayout(page.LayoutId);
            var slide = Presentation.Slides[pageIndex].Clone();
            if (page.Raw.TryGetProperty("background", out var background))
                slide.Background = BuildBackground(background, _catalog, _program.Design.Width, _program.Design.Height);
            var endpoints = new PpjConnectorEndpointResolver(expanded.Elements);
            for (var elementIndex = 0; elementIndex < expanded.Elements.Count; elementIndex++)
                slide.Elements.Add(BuildElement(
                    expanded.Elements[elementIndex],
                    expanded.ElementJson[elementIndex],
                    _catalog,
                    textPrecedence,
                    endpoints));
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

    internal static PresentationTextParagraph BuildMasterTextLevel(JsonElement source, Catalog catalog)
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

    private static PresentationElement BuildElement(
        PpjElementModel element,
        JsonElement raw,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null,
        PpjConnectorEndpointResolver? endpoints = null)
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
                output.Shape = BuildTextShape(element, raw, catalog, "textbox", textPrecedence);
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
                output.Connector = BuildConnector(connector, raw, catalog, endpoints);
                break;
            case PpjGroupElementModel group:
                output.Group = BuildGroup(group, raw, catalog, textPrecedence, endpoints);
                break;
            case PpjPlaceholderElementModel placeholder:
                output.Shape = BuildPlaceholder(placeholder, raw, catalog, textPrecedence);
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
        if (raw.TryGetProperty("action", out var action))
        {
            if (output.ContentCase != PresentationElement.ContentOneofCase.Shape)
                throw Unsupported(element.Id, "shape-level action is currently authored only for shape/text elements");
            output.Shape.Action = BuildAction(action, element.Id);
        }
        if (raw.TryGetProperty("hoverAction", out var hoverAction))
        {
            if (output.ContentCase != PresentationElement.ContentOneofCase.Shape)
                throw Unsupported(element.Id, "shape-level hoverAction is currently authored only for shape/text elements");
            output.Shape.HoverAction = BuildAction(hoverAction, element.Id);
        }
        ApplyAuthoredCompositing(output, raw, element.Id, catalog);
        return output;
    }

    internal static PresentationRunHyperlink BuildAction(JsonElement source, string elementId)
    {
        var action = new PresentationRunHyperlink();
        if (source.TryGetProperty("uri", out var uri)) action.Uri = uri.GetString()!;
        else if (source.TryGetProperty("slide", out var slide)) action.SlideId = slide.GetString()!;
        else if (source.TryGetProperty("customShow", out var customShow)) action.CustomShowId = customShow.GetString()!;
        else if (source.TryGetProperty("verb", out var verb)) action.Action = verb.GetString()!;
        else throw Unsupported(elementId, "action requires uri, slide, customShow, or verb");
        if (source.TryGetProperty("tooltip", out var tooltip)) action.Tooltip = tooltip.GetString()!;
        if (source.TryGetProperty("targetFrame", out var targetFrame)) action.TargetFrame = targetFrame.GetString()!;
        if (source.TryGetProperty("history", out var history)) action.History = history.GetBoolean();
        if (source.TryGetProperty("highlightClick", out var highlightClick)) action.HighlightClick = highlightClick.GetBoolean();
        if (source.TryGetProperty("returnToSlide", out var returnToSlide)) action.ReturnToSlide = returnToSlide.GetBoolean();
        try
        {
            PptxHyperlinkCodec.Validate(action);
        }
        catch (CodecException error)
        {
            throw new CodecException(error.Code, $"PPJ {elementId} action: {error.Message}");
        }
        return action;
    }

    private static void ApplyAuthoredCompositing(
        PresentationElement output,
        JsonElement raw,
        string elementId,
        Catalog catalog)
    {
        if (!raw.TryGetProperty("compositing", out var compositing)) return;
        if (compositing.TryGetProperty("blendMode", out var blendMode) && blendMode.GetString() is not (null or "normal"))
            throw Unsupported(elementId, "non-normal compositing blend modes are not representable by native DrawingML");
        if (compositing.TryGetProperty("isolation", out var isolation) && isolation.GetBoolean())
            throw Unsupported(elementId, "compositing isolation has no bounded native owner");
        if (compositing.TryGetProperty("clipStack", out var clipStack) && clipStack.GetArrayLength() > 0 &&
            output.ContentCase != PresentationElement.ContentOneofCase.Image)
            throw Unsupported(elementId, "compositing clips require an image mask owner");
        if (!compositing.TryGetProperty("opacity", out var opacity)) return;
        var value = catalog.NumberToken(opacity, "opacity", $"{elementId} compositing.opacity");
        switch (output.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                ApplyCompoundShapeOpacity(output.Shape, value, elementId);
                break;
            case PresentationElement.ContentOneofCase.Image:
                output.Image.OpacityThousandthPercent = Opacity(value);
                break;
            case PresentationElement.ContentOneofCase.Connector:
                if (value < 1)
                    output.Connector.LineOpacityThousandthPercent = MultiplyOpacity(
                        output.Connector.HasLineOpacityThousandthPercent,
                        output.Connector.LineOpacityThousandthPercent,
                        value);
                break;
            default:
                throw Unsupported(elementId, $"compositing.opacity is not supported for {output.ContentCase}");
        }
    }

    internal static PresentationTextBody BuildChartTitleBody(
        PpjProgramModel program,
        PpjChartElementModel chart) => ChartCompiler.BuildTitleBody(program, chart);

    private static PresentationShape BuildTextShape(
        PpjElementModel element,
        JsonElement raw,
        Catalog catalog,
        string geometry,
        TextPrecedenceContext? textPrecedence = null)
    {
        var shape = ShapeFrame(element.Frame, geometry);
        var namedStyle = catalog.TextStyle(OptionalString(raw, "styleRef"));
        var inlineStyle = Property(raw, "style");
        var elementTextStyle = Property(raw, "textStyle");
        shape.TextBody = raw.TryGetProperty("text", out var text)
            ? BuildTextBody(text, namedStyle, inlineStyle, catalog, elementTextStyle, textPrecedence)
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
        if (element.Type == "line") return BuildLine(element, raw, catalog);
        var geometry = element.GeometryKind == "custom" ? "custom" : element.GeometryPreset ?? "rect";
        var shape = ShapeFrame(element.Frame, geometry);
        var namedStyle = catalog.ShapeStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id, raw);
        if (raw.TryGetProperty("text", out var text))
        {
            shape.TextBody = BuildTextBody(text, null, Property(raw, "textStyle"), catalog);
            shape.Text = Flatten(shape.TextBody);
        }
        var opacity = catalog.PropertyByPrecedence("shape.opacity", raw, inlineStyle, namedStyle);
        if (opacity is null && raw.TryGetProperty("opacity", out var directOpacity)) opacity = directOpacity;
        if (opacity is { } opacityValue)
            ApplyCompoundShapeOpacity(shape, catalog.NumberToken(opacityValue, "opacity", $"{element.Id} opacity"), element.Id);
        if (geometry == "custom") ApplyCustomGeometry(shape, raw.GetProperty("geometry"), element.Id, allowShapeGraph: true);
        else shape.PresetAdjustments.Add(element.GeometryAdjustments);
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static PresentationShape BuildLine(PpjShapeElementModel element, JsonElement raw, Catalog catalog)
    {
        var shape = ShapeFrame(element.Frame, "custom");
        var namedStyle = catalog.ShapeStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id, raw);
        var stroke = catalog.PropertyByPrecedence("shape.stroke", raw, inlineStyle, namedStyle);
        if (stroke is null && raw.TryGetProperty("stroke", out var directStroke)) stroke = directStroke;
        if (stroke is { } strokeValue) ApplyLine(shape, strokeValue, catalog);
        else if (shape.LineStyle.Length == 0) shape.LineStyle = "none";
        var curve = OptionalString(raw, "curve") ?? "round";
        if (!HasStrokeJoin(stroke))
        {
            if (curve == "round") shape.LineJoin = "round";
            else if (curve == "sharp") shape.LineJoin = "miter";
        }
        shape.StartArrow = Arrow(OptionalString(raw, "startArrow"));
        shape.EndArrow = Arrow(OptionalString(raw, "endArrow"));
        if (raw.TryGetProperty("shadow", out var shadow)) shape.Shadow = BuildShadow(shadow, catalog);
        if (raw.TryGetProperty("glow", out var glow)) shape.Glow = BuildGlow(glow, catalog);
        if (raw.TryGetProperty("reflection", out var reflection)) shape.Reflection = BuildReflection(reflection, catalog);
        var opacity = catalog.PropertyByPrecedence("shape.opacity", raw, inlineStyle, namedStyle);
        if (opacity is null && raw.TryGetProperty("opacity", out var directOpacity)) opacity = directOpacity;
        if (opacity is { } opacityValue)
            ApplyCompoundShapeOpacity(shape, catalog.NumberToken(opacityValue, "opacity", $"{element.Id} opacity"), element.Id);
        var path = raw.TryGetProperty("path", out var explicitPath)
            ? explicitPath
            : PpjLinePathCodec.KimiPath(raw, element.Frame.Width, element.Frame.Height, element.Id);
        PpjLinePathCodec.Apply(shape, path, element.Id);
        ApplyTransform(shape, element.Frame);
        ApplyAccessibility(shape, element.Accessibility);
        return shape;
    }

    private static bool HasStrokeJoin(JsonElement? stroke) =>
        stroke is { ValueKind: JsonValueKind.Object } value && value.TryGetProperty("join", out _);

    private static PresentationShape BuildIcon(PpjIconElementModel element, JsonElement raw, Catalog catalog)
    {
        var definition = PpjIconCatalog.Resolve(element.IconName);
        var shape = ShapeFrame(element.Frame, "custom");
        shape.CatalogIconName = element.IconName;
        var namedStyle = catalog.ShapeStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        if (catalog.PropertyByPrecedence("shape.fill", raw, inlineStyle, namedStyle) is null) shape.FillRgb = "000000";
        ApplyShapeStyle(shape, namedStyle, inlineStyle, catalog, element.Id, raw);
        if (catalog.PropertyByPrecedence("shape.opacity", raw, inlineStyle, namedStyle) is { } opacity)
            ApplyCompoundShapeOpacity(shape, catalog.NumberToken(opacity, "opacity", $"{element.Id} opacity"), element.Id);
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
        var namedStyle = catalog.ImageStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ApplyImageCrop(image, element, raw, namedStyle, inlineStyle, catalog);
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
        ApplyAuthoredImageCompositingClip(image, raw, element.Id);
        if (catalog.PropertyByPrecedence("image.border", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } border)
        {
            var color = catalog.Color(border.GetProperty("color"));
            image.Border = new PresentationImageBorder
            {
                ColorRgb = color.Rgb,
                WidthEmu = Emu(StrokeWidth(border, catalog, $"{element.Id} image border width")),
                Style = LineStyle(OptionalString(border, "dash")),
                Cap = OptionalString(border, "cap") ?? string.Empty,
                Join = OptionalString(border, "join") ?? string.Empty,
            };
            var borderOpacity = OptionalOpacity(border, "opacity", catalog, $"{element.Id} image border opacity") ?? color.Alpha;
            if (borderOpacity < 1) image.Border.OpacityThousandthPercent = Opacity(borderOpacity);
        }
        if (catalog.PropertyByPrecedence("image.shadow", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } shadow)
            image.Shadow = BuildShadow(shadow, catalog);
        if (catalog.PropertyByPrecedence("image.glow", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } glow)
            image.Glow = BuildGlow(glow, catalog);
        if (catalog.PropertyByPrecedence("image.innerShadow", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } innerShadow)
            image.InnerShadow = BuildInnerShadow(innerShadow, catalog);
        if (catalog.PropertyByPrecedence("image.softEdge", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } softEdge)
            image.SoftEdge = BuildSoftEdge(softEdge);
        if (catalog.PropertyByPrecedence("image.reflection", raw, inlineStyle, namedStyle, includeElementWhenUndeclared: true) is { } reflection)
            image.Reflection = BuildReflection(reflection, catalog);
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

    private static void ApplyAuthoredImageCompositingClip(
        PresentationImage image,
        JsonElement raw,
        string elementId)
    {
        if (!raw.TryGetProperty("compositing", out var compositing) ||
            !compositing.TryGetProperty("clipStack", out var clipStack) ||
            clipStack.GetArrayLength() == 0)
            return;
        if (clipStack.GetArrayLength() != 1)
            throw Unsupported(elementId, "compositing.clipStack supports exactly one image clip");
        if (raw.TryGetProperty("mask", out _))
            throw Unsupported(elementId, "compositing.clipStack cannot be combined with image.mask");

        var clip = clipStack[0];
        if (clip.TryGetProperty("inverse", out var inverse) && inverse.GetBoolean())
            throw Unsupported(elementId, "inverse compositing clips have no bounded native owner");
        var geometry = clip.GetProperty("geometry");
        var geometryKind = geometry.GetProperty("kind").GetString();
        if (geometryKind == "custom")
        {
            var customMask = new PresentationShape
            {
                Geometry = "custom",
                WidthEmu = image.WidthEmu,
                HeightEmu = image.HeightEmu,
            };
            ApplyCustomGeometry(customMask, geometry, elementId + " compositing clip");
            PptxCustomGeometryCodec.Validate(customMask, elementId + " compositing clip");
            image.CustomMaskPaths.Add(customMask.CustomPaths);
            return;
        }
        if (geometryKind != "preset")
            throw Unsupported(elementId, "compositing.clipStack supports only bounded preset or custom image clips");
        var preset = geometry.GetProperty("preset").GetString();
        if (preset is null || !PptxPresetGeometryAdjustmentCodec.HasProfile(preset))
            throw Unsupported(elementId, $"compositing image clip preset {preset ?? "(missing)"} is unsupported");
        var adjustments = geometry.TryGetProperty("adjustments", out var rawAdjustments)
            ? rawAdjustments.EnumerateArray().Select(item => item.GetInt32()).ToArray()
            : [];
        PptxPresetGeometryAdjustmentCodec.Validate(preset, adjustments, elementId + " compositing clip");
        image.MaskPreset = preset;
        image.MaskPresetAdjustments.Add(adjustments);
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
        if (element.PlaybackTrigger is { } playbackTrigger) media.PlaybackTrigger = playbackTrigger;
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
        var defaultTextStyle = catalog.PropertyByPrecedence("table.defaultTextStyle", raw, inlineStyle, namedStyle);
        var defaultCellFill = catalog.PropertyByPrecedence("table.defaultCellFill", raw, inlineStyle, namedStyle);
        var headerTextStyle = catalog.PropertyByPrecedence("table.headerTextStyle", raw, inlineStyle, namedStyle);
        var headerCellFill = catalog.PropertyByPrecedence("table.headerCellFill", raw, inlineStyle, namedStyle);
        var cellStyle = catalog.PropertyByPrecedence("table.cellStyle", raw, inlineStyle, namedStyle);
        var bodyStyles = catalog.PropertyByPrecedence("table.bodyStyles", raw, inlineStyle, namedStyle);
        var firstRowStyle = catalog.PropertyByPrecedence("table.firstRowStyle", raw, inlineStyle, namedStyle);
        var lastRowStyle = catalog.PropertyByPrecedence("table.lastRowStyle", raw, inlineStyle, namedStyle);
        var firstColumnStyle = catalog.PropertyByPrecedence("table.firstColumnStyle", raw, inlineStyle, namedStyle);
        var lastColumnStyle = catalog.PropertyByPrecedence("table.lastColumnStyle", raw, inlineStyle, namedStyle);
        var rowOverColumnValue = catalog.PropertyByPrecedence("table.rowOverColumn", raw, inlineStyle, namedStyle);
        var rowOverColumn = rowOverColumnValue is null
            ? true
            : catalog.BooleanToken(rowOverColumnValue.Value, "boolean", $"table {element.Id} rowOverColumn");
        var declaredHeaderRows = catalog.PropertyByPrecedence("table.headerRows", raw, inlineStyle, namedStyle);
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
        if (catalog.PropertyByPrecedence("table.bandedRows", raw, inlineStyle, namedStyle) is { } bandedRows)
            table.BandedRows = catalog.BooleanToken(bandedRows, "boolean", $"table {element.Id} bandedRows");
        if (catalog.PropertyByPrecedence("table.bandedColumns", raw, inlineStyle, namedStyle) is { } bandedColumns)
            table.BandedColumns = catalog.BooleanToken(bandedColumns, "boolean", $"table {element.Id} bandedColumns");
        if (catalog.PropertyByPrecedence("table.firstColumnEmphasis", raw, inlineStyle, namedStyle) is { } firstColumn)
            table.FirstColumn = catalog.BooleanToken(firstColumn, "boolean", $"table {element.Id} firstColumnEmphasis");
        if (catalog.PropertyByPrecedence("table.lastColumnEmphasis", raw, inlineStyle, namedStyle) is { } lastColumn)
            table.LastColumn = catalog.BooleanToken(lastColumn, "boolean", $"table {element.Id} lastColumnEmphasis");
        if (catalog.PropertyByPrecedence("table.lastRow", raw, inlineStyle, namedStyle) is { } lastRow)
            table.LastRow = catalog.BooleanToken(lastRow, "boolean", $"table {element.Id} lastRow");
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
                    path,
                    opacity => catalog.NumberToken(opacity, "opacity", $"{path} opacity"),
                    fit => catalog.StringToken(fit, "string", $"{path} fit")),
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
            WidthPoints = StrokeWidth(stroke, catalog, "table cell border width"),
            DashStyle = ChartDash(OptionalString(stroke, "dash")),
            Cap = OptionalString(stroke, "cap") ?? string.Empty,
            Join = OptionalString(stroke, "join") ?? string.Empty,
        };
        var opacity = OptionalOpacity(stroke, "opacity", catalog, "table cell border opacity") ?? color.Alpha;
        if (opacity < 1) output.OpacityThousandthPercent = Opacity(opacity);
        return output;
    }

    private static PresentationConnector BuildConnector(PpjConnectorElementModel element, JsonElement raw, Catalog catalog,
        PpjConnectorEndpointResolver? endpoints)
    {
        RejectUnsupportedFrameTransform(element.Id, element.Frame, "connector");
        var resolved = endpoints?.Resolve(element) ?? PpjConnectorEndpointResolver.ResolveLiteral(element);
        var connector = new PresentationConnector
        {
            ConnectorType = element.ConnectorType,
            StartXEmu = resolved.StartX,
            StartYEmu = resolved.StartY,
            EndXEmu = resolved.EndX,
            EndYEmu = resolved.EndY,
            StartTargetId = element.From.ConnectionSite is not null ? element.From.ElementId! : string.Empty,
            StartConnectionSiteIndex = element.From.ConnectionSite ?? 0,
            EndTargetId = element.To.ConnectionSite is not null ? element.To.ElementId! : string.Empty,
            EndConnectionSiteIndex = element.To.ConnectionSite ?? 0,
            StartFrameAnchor = element.From.ConnectionSite is null && element.From.ElementId is { } startTarget
                ? new PresentationConnectorFrameAnchor { TargetId = startTarget, Anchor = element.From.Anchor! } : null,
            EndFrameAnchor = element.To.ConnectionSite is null && element.To.ElementId is { } endTarget
                ? new PresentationConnectorFrameAnchor { TargetId = endTarget, Anchor = element.To.Anchor! } : null,
        };
        ApplyLine(connector, raw.GetProperty("stroke"), catalog);
        connector.StartArrow = Arrow(OptionalString(raw, "startArrow"));
        connector.EndArrow = Arrow(OptionalString(raw, "endArrow"));
        connector.StartArrowWidth = OptionalString(raw, "startArrowWidth") ?? string.Empty;
        connector.StartArrowLength = OptionalString(raw, "startArrowLength") ?? string.Empty;
        connector.EndArrowWidth = OptionalString(raw, "endArrowWidth") ?? string.Empty;
        connector.EndArrowLength = OptionalString(raw, "endArrowLength") ?? string.Empty;
        if (raw.TryGetProperty("bendAdjustment", out var bend)) connector.BendAdjustment = bend.GetInt32();
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
                    Image = BuildDiagramNodeImage(item, rawNode, catalog),
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
        if (resolvedLayout != "picture" && element.Nodes.Any(node => node.Image is not null))
            throw Unsupported(element.Id, "diagram node image paint is only valid for the picture layout");
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
        if (rawNode.TryGetProperty("image", out _))
            shape.ImageFill = BuildSmartArtNodeImagePaint(item.Node, rawNode, item.Frame, catalog);
        if (geometry is { } value)
        {
            if (geometryKind == "custom") ApplyCustomGeometry(shape, value, item.Node.Id, allowShapeGraph: true);
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
            ApplyCompoundShapeOpacity(shape, catalog.NumberToken(opacity, "opacity", $"{item.Node.Id} opacity"), item.Node.Id);
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

    private static PresentationImage BuildDiagramNodeImage(
        DiagramLayoutNode item,
        JsonElement rawNode,
        Catalog catalog)
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
        if (rawNode.TryGetProperty("image", out _))
        {
            var paint = BuildSmartArtNodeImagePaint(item.Node, rawNode, frame, catalog);
            image.Crop = paint.Crop;
            image.Tiled = paint.Mode == PresentationImagePaint.Types.Mode.Tile;
            if (paint.HasOpacityThousandthPercent)
                image.OpacityThousandthPercent = paint.OpacityThousandthPercent;
        }
        else if (catalog.AssetDimensions(item.Node.AssetId!) is { } dimensions)
            image.Crop = DiagramCoverCrop(frame.Width, frame.Height, dimensions.Width, dimensions.Height);
        return image;
    }

    private static PresentationImagePaint BuildSmartArtNodeImagePaint(
        PpjSmartArtNodeModel node,
        JsonElement rawNode,
        PpjFrameModel frame,
        Catalog catalog)
    {
        if (node.AssetId is null || !rawNode.TryGetProperty("image", out var image) ||
            image.ValueKind != JsonValueKind.Object)
            throw Unsupported(node.Id, "SmartArt node image paint requires an image object and asset");
        var imageObject = JsonNode.Parse(image.GetRawText())?.AsObject() ??
            throw Unsupported(node.Id, "SmartArt node image paint must be an object");
        imageObject["type"] = "image";
        imageObject["asset"] = node.AssetId;
        using var document = JsonDocument.Parse(imageObject.ToJsonString());
        var output = PpjImagePaintLowering.Build(
            document.RootElement,
            frame.Width,
            frame.Height,
            catalog.NativeAssetId,
            catalog.AssetDimensions,
            $"{node.Id} image",
            resolveOpacity: value => catalog.NumberToken(value, "opacity", $"{node.Id} image opacity"),
            resolveFit: value => catalog.StringToken(value, "string", $"{node.Id} image fit"));
        if (output.Mode is not (PresentationImagePaint.Types.Mode.Stretch or PresentationImagePaint.Types.Mode.Tile))
            throw Unsupported(node.Id, "SmartArt node image paint fit is outside the stretch/tile profile");
        return output;
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

    private static PresentationGroup BuildGroup(
        PpjGroupElementModel element,
        JsonElement raw,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null,
        PpjConnectorEndpointResolver? endpoints = null)
    {
        var group = new PresentationGroup
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            ChildLeftEmu = Emu(element.ChildFrame.X),
            ChildTopEmu = Emu(element.ChildFrame.Y),
            ChildWidthEmu = Emu(element.ChildFrame.Width),
            ChildHeightEmu = Emu(element.ChildFrame.Height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        var childJsonById = raw.GetProperty("elements").EnumerateArray()
            .ToDictionary(child => child.GetProperty("id").GetString()!, StringComparer.Ordinal);
        var orderedElements = element.ReadingOrder.Count == 0
            ? element.Elements
            : element.ReadingOrder.Select(id => element.Elements.Single(child => child.Id.Equals(id, StringComparison.Ordinal))).ToArray();
        foreach (var child in orderedElements)
        {
            if (!childJsonById.TryGetValue(child.Id, out var childJson))
                throw Unsupported(element.Id, $"group readingOrder references missing child {child.Id}");
            group.Children.Add(BuildElement(child, childJson, catalog, textPrecedence, endpoints));
        }
        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static PresentationShape BuildPlaceholder(
        PpjPlaceholderElementModel element,
        JsonElement raw,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null)
    {
        var shape = BuildTextShape(element, raw, catalog, "textbox", textPrecedence);
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
        JsonElement? element = null)
    {
        var fill = catalog.PropertyByPrecedence("shape.fill", element, inline, named);
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
                    $"element {elementId} fill",
                    opacity => catalog.NumberToken(opacity, "opacity", $"element {elementId} fill opacity"),
                    fit => catalog.StringToken(fit, "string", $"element {elementId} fill fit"));
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
        var stroke = catalog.PropertyByPrecedence("shape.stroke", element, inline, named);
        if (stroke is { } strokeValue) ApplyLine(target, strokeValue, catalog);
        else target.LineStyle = "none";
        var shadow = catalog.PropertyByPrecedence("shape.shadow", element, inline, named);
        if (shadow is { } shadowValue)
            target.Shadow = BuildShadow(shadowValue, catalog);
        var glow = catalog.PropertyByPrecedence("shape.glow", element, inline, named);
        if (glow is { } glowValue)
            target.Glow = BuildGlow(glowValue, catalog);
        var innerShadow = catalog.PropertyByPrecedence("shape.innerShadow", element, inline, named);
        if (innerShadow is { } innerShadowValue)
            target.InnerShadow = BuildInnerShadow(innerShadowValue, catalog);
        var softEdge = catalog.PropertyByPrecedence(
            "shape.softEdge",
            element,
            inline,
            named,
            includeElementWhenUndeclared: true);
        if (softEdge is { } softEdgeValue)
            target.SoftEdge = BuildSoftEdge(softEdgeValue);
        var reflection = catalog.PropertyByPrecedence("shape.reflection", element, inline, named);
        if (reflection is { } reflectionValue)
            target.Reflection = BuildReflection(reflectionValue, catalog);
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

    internal static void SetCompoundShapeOpacity(PresentationShape target, double opacity, string elementId)
    {
        if (!double.IsFinite(opacity) || opacity is < 0 or > 1)
            throw Unsupported(elementId, "source-bound compound shape opacity must be between 0 and 1");
        var alpha = Opacity(opacity);
        if (target.FillRgb.Length > 0 || target.FillScheme.Length > 0)
        {
            if (opacity < 1) target.FillOpacityThousandthPercent = alpha;
            else target.ClearFillOpacityThousandthPercent();
        }
        if (target.GradientFill is not null)
        {
            foreach (var stop in target.GradientFill.Stops)
            {
                if (opacity < 1) stop.OpacityThousandthPercent = alpha;
                else stop.ClearOpacityThousandthPercent();
            }
        }
        if (target.ImageFill is not null)
        {
            if (opacity < 1) target.ImageFill.OpacityThousandthPercent = alpha;
            else target.ImageFill.ClearOpacityThousandthPercent();
        }
        if (target.LineStyle != "none" && (target.LineRgb.Length > 0 || target.LineScheme.Length > 0))
        {
            if (opacity < 1) target.LineOpacityThousandthPercent = alpha;
            else target.ClearLineOpacityThousandthPercent();
        }
        if (target.Shadow is not null)
        {
            if (opacity < 1) target.Shadow.OpacityThousandthPercent = alpha;
            else target.Shadow.ClearOpacityThousandthPercent();
        }
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
                "text fill",
                opacity => catalog.NumberToken(opacity, "opacity", "text fill opacity"),
                fit => catalog.StringToken(fit, "string", "text fill fit"));
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
        target.LineWidthEmu = Emu(StrokeWidth(stroke, catalog, "shape stroke width"));
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        var opacity = OptionalOpacity(stroke, "opacity", catalog, "shape stroke opacity") ?? color.Alpha;
        if (opacity < 1) target.LineOpacityThousandthPercent = Opacity(opacity);
    }

    private static void ApplyLine(PresentationConnector target, JsonElement stroke, Catalog catalog)
    {
        var color = catalog.Color(stroke.GetProperty("color"));
        target.LineRgb = color.Rgb;
        target.LineWidthEmu = Emu(StrokeWidth(stroke, catalog, "connector stroke width"));
        target.LineStyle = LineStyle(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        var opacity = OptionalOpacity(stroke, "opacity", catalog, "connector stroke opacity") ?? color.Alpha;
        if (opacity < 1) target.LineOpacityThousandthPercent = Opacity(opacity);
    }

    private static PresentationTextBody BuildTextBody(
        JsonElement text,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog,
        JsonElement? elementTextStyle = null,
        TextPrecedenceContext? textPrecedence = null) =>
        BuildTextBody(text, namedStyle, null, inlineStyle, catalog, elementTextStyle, textPrecedence);

    private static PresentationTextBody BuildTextBody(
        JsonElement text,
        JsonElement? namedStyle,
        JsonElement? middleStyle,
        JsonElement? inlineStyle,
        Catalog catalog,
        JsonElement? elementTextStyle = null,
        TextPrecedenceContext? textPrecedence = null)
    {
        var body = new PresentationTextBody();
        // Structured rich text may carry a text-container style of its own.
        // Keep this separate from paragraph/run styles, but let it override
        // the surrounding element style for the body properties it declares.
        var richTextStyle = text.ValueKind == JsonValueKind.Object ? Property(text, "style") : null;
        var bodyStyle = MergeJsonObjects(inlineStyle, richTextStyle);
        var paragraphStyle = MergeJsonObjects(true, Property(inlineStyle, "paragraph"), Property(richTextStyle, "paragraph"));
        ApplyBodyProperties(body.BodyProperties = new PresentationTextBodyProperties(), namedStyle, middleStyle, bodyStyle);
        if (text.ValueKind == JsonValueKind.String)
        {
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(middleStyle, "paragraph"),
                paragraphStyle,
                null,
                catalog);
            paragraph.Runs.Add(BuildRun(
                text.GetString()!,
                namedStyle,
                middleStyle,
                inlineStyle,
                null,
                null,
                catalog,
                textPrecedence is null && elementTextStyle is null
                    ? null
                    : (textPrecedence ?? new TextPrecedenceContext(null, null, null, null)) with
                    {
                        ElementStyle = elementTextStyle,
                    }));
            body.Paragraphs.Add(paragraph);
            return body;
        }
        foreach (var paragraphJson in text.GetProperty("paragraphs").EnumerateArray())
        {
            var runPrecedence = textPrecedence is null && elementTextStyle is null
                ? null
                : (textPrecedence ?? new TextPrecedenceContext(null, null, null, null)) with
                {
                    ElementStyle = elementTextStyle,
                    ParagraphStyle = Property(paragraphJson, "style"),
                };
            var paragraph = new PresentationTextParagraph();
            ApplyParagraphStyle(
                paragraph,
                Property(namedStyle, "paragraph"),
                Property(middleStyle, "paragraph"),
                paragraphStyle,
                Property(paragraphJson, "style"),
                catalog);
            foreach (var run in paragraphJson.GetProperty("runs").EnumerateArray())
            {
                paragraph.Runs.Add(run.TryGetProperty("field", out var field)
                    ? BuildFieldRun(
                        field,
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        Property(run, "hyperlink"),
                        catalog,
                        runPrecedence)
                    : run.TryGetProperty("formula", out var formula)
                    ? BuildFormulaRun(
                        formula.GetProperty("source").GetString()!,
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        catalog,
                        runPrecedence)
                    : run.TryGetProperty("break", out var lineBreak) && lineBreak.ValueKind == JsonValueKind.True
                    ? BuildLineBreakRun(
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        Property(run, "hyperlink"),
                        catalog,
                        runPrecedence)
                    : BuildRun(
                        run.GetProperty("text").GetString()!,
                        namedStyle,
                        middleStyle,
                        inlineStyle,
                        Property(run, "style"),
                        Property(run, "hyperlink"),
                        catalog,
                        runPrecedence));
            }
            body.Paragraphs.Add(paragraph);
        }
        return body;
    }

    private static PresentationTextRun BuildFieldRun(
        JsonElement source,
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        JsonElement? hyperlink,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null)
    {
        var type = source.GetProperty("type").GetString()!;
        var text = source.GetProperty("text").GetString()!;
        var run = BuildRun(string.Empty, namedBox, middleBox, inlineBox, inlineRun, hyperlink, catalog, textPrecedence);
        var id = OptionalString(source, "id");
        if (!Guid.TryParseExact(id, "B", out _))
        {
            var digest = MD5.HashData(Encoding.UTF8.GetBytes($"officekit-field\0{type}\0{text}"));
            id = $"{{{new Guid(digest).ToString()}}}";
        }
        run.Field = new PresentationTextField
        {
            Id = id,
            Type = type,
            Text = text,
            Automatic = source.TryGetProperty("automatic", out var automatic) && automatic.ValueKind == JsonValueKind.True,
        };
        return run;
    }

    private static PresentationTextRun BuildLineBreakRun(
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        JsonElement? hyperlink,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null)
    {
        var run = BuildRun(string.Empty, namedBox, middleBox, inlineBox, inlineRun, hyperlink, catalog, textPrecedence);
        run.LineBreak = true;
        return run;
    }

    private static PresentationTextRun BuildFormulaRun(
        string source,
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null)
    {
        var run = new PresentationTextRun { Formula = PpjLatexCompiler.Compile(source) };
        var inlineDefault = FirstProperty(inlineBox, null, "defaultText");
        var middleDefault = FirstProperty(middleBox, null, "defaultText");
        var namedDefault = FirstProperty(namedBox, null, "defaultText");
        if (TextPropertyByPrecedence(
                catalog,
                "size",
                inlineRun,
                inlineDefault,
                middleDefault,
                namedDefault,
                textPrecedence) is { } size)
            run.FontSizePoints = catalog.PositiveNumberToken(size, "size", "text size");
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
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null) =>
        BuildRun(text, namedBox, null, inlineBox, inlineRun, hyperlink, catalog, textPrecedence);

    private static PresentationTextRun BuildRun(
        string text,
        JsonElement? namedBox,
        JsonElement? middleBox,
        JsonElement? inlineBox,
        JsonElement? inlineRun,
        JsonElement? hyperlink,
        Catalog catalog,
        TextPrecedenceContext? textPrecedence = null)
    {
        var run = new PresentationTextRun { Text = text };
        var inlineDefault = FirstProperty(inlineBox, null, "defaultText");
        var middleDefault = FirstProperty(middleBox, null, "defaultText");
        var namedDefault = FirstProperty(namedBox, null, "defaultText");
        var bold = TextPropertyByPrecedence(catalog, "bold", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var italic = TextPropertyByPrecedence(catalog, "italic", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var size = TextPropertyByPrecedence(catalog, "size", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var paint = FirstTextPaint(inlineRun, inlineDefault, middleDefault, namedDefault);
        var shadow = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "shadow");
        var glow = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "glow");
        var highlight = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "highlight");
        var font = TextPropertyByPrecedence(catalog, "font", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var family = TextPropertyByPrecedence(catalog, "fontFamily", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var eastAsia = TextPropertyByPrecedence(catalog, "fontFamilyEastAsia", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var complexScript = TextPropertyByPrecedence(catalog, "fontFamilyComplexScript", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        var underline = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "underline");
        var strike = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "strike");
        var kerning = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "kerning");
        var spacing = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "letterSpacing");
        var baseline = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "baseline");
        var capitalization = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "capitalization");
        var language = TextPropertyByPrecedence(catalog, "language", inlineRun, inlineDefault, middleDefault, namedDefault, textPrecedence);
        if (bold is { } boldValue) run.Bold = catalog.BooleanToken(boldValue, "boolean", "text bold");
        if (italic is { } italicValue) run.Italic = catalog.BooleanToken(italicValue, "boolean", "text italic");
        if (size is { } sizeValue) run.FontSizePoints = catalog.PositiveNumberToken(sizeValue, "size", "text size");
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
        if (glow is { } glowValue) run.Glow = BuildGlow(glowValue, catalog);
        var innerShadow = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "innerShadow");
        if (innerShadow is { } innerShadowValue) run.InnerShadow = BuildInnerShadow(innerShadowValue, catalog);
        var reflection = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "reflection");
        if (reflection is { } reflectionValue) run.Reflection = BuildReflection(reflectionValue, catalog);
        var softEdge = FirstProperty(inlineRun, inlineDefault, middleDefault, namedDefault, "softEdge");
        if (softEdge is { } softEdgeValue) run.SoftEdge = BuildSoftEdge(softEdgeValue);
        if (highlight is { } highlightValue)
        {
            var resolved = catalog.Color(highlightValue);
            if (resolved.Alpha < 1)
                throw Unsupported("text", "highlight alpha is not part of the bounded DrawingML highlight profile");
            run.HighlightRgb = resolved.Rgb;
        }
        if (family is { } familyValue) run.FontFamily = catalog.StringToken(familyValue, "string", "text fontFamily");
        else if (font is { } fontValue) run.FontFamily = catalog.Font(fontValue);
        if (eastAsia is { } eastAsiaValue) run.FontFamilyEastAsia = catalog.StringToken(eastAsiaValue, "string", "text fontFamilyEastAsia");
        else if (run.FontFamily.Length > 0) run.FontFamilyEastAsia = run.FontFamily;
        if (complexScript is { } complexScriptValue)
            run.FontFamilyComplexScript = catalog.StringToken(complexScriptValue, "string", "text fontFamilyComplexScript");
        if (underline is { } underlineValue) run.Underline = NativeUnderline(underlineValue.GetString()!);
        if (strike is { } strikeValue) run.Strike = NativeStrike(strikeValue);
        if (kerning is { } kerningValue) run.FontKerningPoints = kerningValue.GetDouble();
        if (spacing is { } spacingValue) run.FontSpacingPoints = spacingValue.GetDouble();
        if (baseline is { } baselineValue) run.FontBaselinePercent = baselineValue.GetDouble();
        if (capitalization is { } capitalizationValue) run.FontCaps = capitalizationValue.GetString()!;
        if (language is { } languageValue) run.Language = catalog.LanguageTagToken(languageValue, "text language");
        if (hyperlink is { } link)
        {
            run.RunHyperlink = new PresentationRunHyperlink { Uri = link.GetProperty("uri").GetString()! };
            if (link.TryGetProperty("tooltip", out var tooltip)) run.RunHyperlink.Tooltip = tooltip.GetString()!;
        }
        return run;
    }

    private static JsonElement? TextPropertyByPrecedence(
        Catalog catalog,
        string field,
        JsonElement? inlineRun,
        JsonElement? inlineDefault,
        JsonElement? middleDefault,
        JsonElement? namedDefault,
        TextPrecedenceContext? textPrecedence)
    {
        var paragraphDefault = textPrecedence?.ParagraphStyle is { } paragraph
            ? Property(paragraph, "defaultText")
            : null;
        return textPrecedence is null
            ? catalog.PropertyByPrecedence(
                $"text.{field}",
                inlineRun,
                inlineDefault,
                namedDefault,
                middleDefault,
                includeElementWhenUndeclared: true)
            : catalog.TextPropertyByPrecedence(
                $"text.{field}",
                inlineRun,
                paragraphDefault,
                textPrecedence.ElementStyle,
                namedDefault,
                textPrecedence.LayoutStyle,
                textPrecedence.MasterStyle,
                inlineDefault,
                middleDefault);
    }

    private static PresentationTextStyle BuildTextStyle(JsonElement? style, Catalog catalog, bool paragraphDefaults = false)
    {
        var output = new PresentationTextStyle();
        if (style is not { } value) return output;
        if (value.TryGetProperty("bold", out var bold)) output.Bold = catalog.BooleanToken(bold, "boolean", "text style bold");
        if (value.TryGetProperty("italic", out var italic)) output.Italic = catalog.BooleanToken(italic, "boolean", "text style italic");
        if (value.TryGetProperty("size", out var size)) output.FontSizePoints = catalog.PositiveNumberToken(size, "size", "text style size");
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
        if (value.TryGetProperty("shadow", out var shadow))
            output.Shadow = paragraphDefaults
                ? BuildParagraphShadow(shadow, catalog.Color, opacity => catalog.NumberToken(opacity, "opacity", "paragraph default shadow opacity"), catalog.HasGrammarToken)
                : BuildShadow(shadow, catalog);
        if (value.TryGetProperty("glow", out var glow))
            output.Glow = paragraphDefaults
                ? BuildParagraphGlow(glow, catalog.Color, opacity => catalog.NumberToken(opacity, "opacity", "paragraph default glow opacity"), catalog.HasGrammarToken)
                : BuildGlow(glow, catalog);
        if (value.TryGetProperty("innerShadow", out var innerShadow))
            output.InnerShadow = paragraphDefaults
                ? BuildParagraphInnerShadow(innerShadow, catalog.Color, opacity => catalog.NumberToken(opacity, "opacity", "paragraph default inner shadow opacity"), catalog.HasGrammarToken)
                : BuildInnerShadow(innerShadow, catalog);
        if (value.TryGetProperty("reflection", out var reflection))
            output.Reflection = paragraphDefaults
                ? BuildChartTextReflection(reflection, opacity => catalog.NumberToken(opacity, "opacity", "paragraph default reflection opacity"))
                : BuildReflection(reflection, catalog);
        if (value.TryGetProperty("softEdge", out var softEdge)) output.SoftEdge = BuildSoftEdge(softEdge);
        if (value.TryGetProperty("highlight", out var highlight))
        {
            var resolved = catalog.Color(highlight);
            if (resolved.Alpha < 1)
                throw Unsupported("text", "highlight alpha is not part of the bounded DrawingML highlight profile");
            output.HighlightRgb = resolved.Rgb;
        }
        if (value.TryGetProperty("fontFamily", out var family)) output.FontFamily = catalog.StringToken(family, "string", "text style fontFamily");
        else if (value.TryGetProperty("font", out var font)) output.FontFamily = catalog.Font(font);
        if (value.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) output.FontFamilyEastAsia = catalog.StringToken(eastAsia, "string", "text style fontFamilyEastAsia");
        else if (output.HasFontFamily) output.FontFamilyEastAsia = output.FontFamily;
        if (value.TryGetProperty("fontFamilyComplexScript", out var complexScript))
            output.FontFamilyComplexScript = catalog.StringToken(complexScript, "string", "text style fontFamilyComplexScript");
        if (value.TryGetProperty("underline", out var underline)) output.Underline = NativeUnderline(underline.GetString()!);
        if (value.TryGetProperty("strike", out var strike)) output.Strike = NativeStrike(strike);
        if (value.TryGetProperty("kerning", out var kerning)) output.FontKerningPoints = kerning.GetDouble();
        if (value.TryGetProperty("letterSpacing", out var spacing)) output.FontSpacingPoints = spacing.GetDouble();
        if (value.TryGetProperty("baseline", out var baseline)) output.FontBaselinePercent = baseline.GetDouble();
        if (value.TryGetProperty("capitalization", out var capitalization)) output.FontCaps = capitalization.GetString()!;
        if (value.TryGetProperty("language", out var language)) output.Language = catalog.LanguageTagToken(language, "text style language");
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
        var anchorCenter = FirstProperty(inline, middle, named, "anchorCenter");
        if (anchorCenter is { } anchorCenterValue)
            target.AnchorCenter = anchorCenterValue.GetBoolean();
        var forceAntiAlias = FirstProperty(inline, middle, named, "forceAntiAlias");
        if (forceAntiAlias is { } forceAntiAliasValue)
            target.ForceAntiAlias = forceAntiAliasValue.GetBoolean();
        var spaceFirstLastParagraph = FirstProperty(inline, middle, named, "spaceFirstLastParagraph");
        if (spaceFirstLastParagraph is { } spaceFirstLastParagraphValue)
            target.SpaceFirstLastParagraph = spaceFirstLastParagraphValue.GetBoolean();
        var compatibleLineSpacing = FirstProperty(inline, middle, named, "compatibleLineSpacing");
        if (compatibleLineSpacing is { } compatibleLineSpacingValue)
            target.CompatibleLineSpacing = compatibleLineSpacingValue.GetBoolean();
        var fromWordArt = FirstProperty(inline, middle, named, "fromWordArt");
        if (fromWordArt is { } fromWordArtValue)
            target.FromWordArt = fromWordArtValue.GetBoolean();
        var textWarpPreset = FirstProperty(inline, middle, named, "textWarpPreset");
        if (textWarpPreset is { } textWarpPresetValue)
            target.TextWarpPreset = PptxBodyPropertiesCodec.ParseTextWarpPreset(textWarpPresetValue.GetString()!);
        var textWarpAdjustments = FirstProperty(inline, middle, named, "textWarpAdjustments");
        if (textWarpAdjustments is { } textWarpAdjustmentsValue)
            target.TextWarpAdjustments.Add(ParseTextWarpAdjustments(textWarpAdjustmentsValue));
        var flatTextZ = FirstProperty(inline, middle, named, "flatTextZ");
        if (flatTextZ is { } flatTextZValue)
            target.FlatTextZ = ParseFlatTextZ(flatTextZValue);
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
        var rotation = FirstProperty(inline, middle, named, "rotation");
        if (rotation is { } angle) target.RotationAngle60000 = Angle(angle.GetDouble());
        var verticalOverflow = FirstProperty(inline, middle, named, "verticalOverflow");
        if (verticalOverflow is { } overflow) target.VerticalOverflowMode = overflow.GetString()!;
        var horizontalOverflow = FirstProperty(inline, middle, named, "horizontalOverflow");
        if (horizontalOverflow is { } horizontal) target.HorizontalOverflowMode = horizontal.GetString()!;
        var upright = FirstProperty(inline, middle, named, "upright");
        if (upright is { } uprightValue) target.Upright = uprightValue.GetBoolean();
        var normalAutoFit = FirstProperty(inline, middle, named, "normalAutoFit");
        if (normalAutoFit is { } normal)
        {
            if (normal.ValueKind != JsonValueKind.Object)
                throw Unsupported("text body normalAutoFit", "normalAutoFit must be an object");
            var targetNormal = new PresentationNormalAutoFit();
            if (normal.TryGetProperty("fontScale", out var fontScale))
                targetNormal.FontScale1000 = NormalAutoFitPercentage(fontScale, "fontScale", 1, 100);
            if (normal.TryGetProperty("lineSpacingReduction", out var lineSpacingReduction))
                targetNormal.LineSpacingReduction1000 = NormalAutoFitPercentage(lineSpacingReduction, "lineSpacingReduction", 0, 13_200);
            if (targetNormal.FontScaleCase != PresentationNormalAutoFit.FontScaleOneofCase.None ||
                targetNormal.LineSpacingReductionCase != PresentationNormalAutoFit.LineSpacingReductionOneofCase.None)
                target.NormalAutoFit = targetNormal;
        }
    }

    private static IEnumerable<PresentationTextWarpAdjustment> ParseTextWarpAdjustments(JsonElement value)
    {
        foreach (var item in value.EnumerateArray())
        {
            yield return new PresentationTextWarpAdjustment
            {
                Name = item.GetProperty("name").GetString()!,
                Value = item.GetProperty("value").GetInt32(),
            };
        }
    }

    private static long ParseFlatTextZ(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt64(out var coordinate) ||
            coordinate < int.MinValue || coordinate > int.MaxValue)
            throw Unsupported("text body flatTextZ", "flatTextZ must be a signed integer in the bounded 32-bit coordinate range");
        return coordinate;
    }

    private static int NormalAutoFitPercentage(JsonElement value, string field, double minimum, double maximum)
    {
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetDouble(out var number) || !double.IsFinite(number))
            throw Unsupported("text body normalAutoFit." + field, "percentage must be a finite number");
        var scaled = number * 1_000d;
        var rounded = Math.Round(scaled, MidpointRounding.AwayFromZero);
        if (number < minimum || number > maximum || Math.Abs(scaled - rounded) > 1e-7)
            throw Unsupported("text body normalAutoFit." + field, $"percentage must be between {minimum} and {maximum} with at most three decimal places");
        return checked((int)rounded);
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
        if (FirstProperty(direct, inline, middle, named, "fontAlignment") is { } fontAlignment)
            target.FontAlignment = fontAlignment.GetString() switch
            {
                "auto" or "top" or "center" or "baseline" or "bottom" => fontAlignment.GetString(),
                _ => throw Unsupported("paragraph", "fontAlignment must be auto, top, center, baseline or bottom"),
            };
        if (FirstProperty(direct, inline, middle, named, "defaultTabSize") is { } defaultTabSize)
        {
            var points = defaultTabSize.GetDouble();
            if (!double.IsFinite(points) || points < int.MinValue / EmuPerPoint || points > int.MaxValue / EmuPerPoint)
                throw Unsupported("paragraph", "defaultTabSize must fit a signed 32-bit EMU coordinate");
            target.DefaultTabSizeEmu = checked((int)Emu(points));
        }
        if (FirstProperty(direct, inline, middle, named, "eastAsianLineBreak") is { } eastAsianLineBreak)
            target.EastAsianLineBreak = eastAsianLineBreak.GetBoolean();
        if (FirstProperty(direct, inline, middle, named, "latinLineBreak") is { } latinLineBreak)
            target.LatinLineBreak = latinLineBreak.GetBoolean();
        if (FirstProperty(direct, inline, middle, named, "hangingPunctuation") is { } hangingPunctuation)
            target.HangingPunctuation = hangingPunctuation.GetBoolean();
        if (FirstProperty(direct, inline, middle, named, "direction") is { } direction)
            target.RightToLeft = direction.GetString() switch
            {
                "left-to-right" => false,
                "right-to-left" => true,
                _ => throw Unsupported("paragraph", "direction must be left-to-right or right-to-left"),
            };
        if (FirstProperty(direct, inline, middle, named, "level") is { } level)
            target.Level = checked((uint)level.GetInt32());
        if (FirstProperty(direct, inline, middle, named, "indent") is { } indent)
            target.MarginLeftEmu = Emu(indent.GetDouble());
        if (FirstProperty(direct, inline, middle, named, "rightIndent") is { } rightIndent)
        {
            var points = rightIndent.GetDouble();
            if (!double.IsFinite(points) || points < 0 || points > 4032)
                throw Unsupported("paragraph", "rightIndent must be from 0 through 4032 points");
            target.MarginRightEmu = Emu(points);
        }
        if (FirstProperty(direct, inline, middle, named, "hanging") is { } hanging)
            target.IndentEmu = -Emu(hanging.GetDouble());
        foreach (var layer in new[] { direct, inline, middle, named })
        {
            if (layer is not { ValueKind: JsonValueKind.Object } paragraph) continue;
            var hasPoints = paragraph.TryGetProperty("spaceBefore", out var before);
            var hasMultiplier = paragraph.TryGetProperty("spaceBeforeMultiplier", out var multiplier);
            if (hasPoints && hasMultiplier) throw Unsupported("paragraph", "spaceBefore and spaceBeforeMultiplier are mutually exclusive");
            if (hasPoints) target.SpaceBeforePoints = Math.Round(before.GetDouble() * 100) / 100d;
            else if (hasMultiplier) target.SpaceBeforeMultiplier = Math.Round(multiplier.GetDouble() * 100_000) / 100_000d;
            if (hasPoints || hasMultiplier) break;
        }
        foreach (var layer in new[] { direct, inline, middle, named })
        {
            if (layer is not { ValueKind: JsonValueKind.Object } paragraph) continue;
            var hasPoints = paragraph.TryGetProperty("spaceAfter", out var before);
            var hasMultiplier = paragraph.TryGetProperty("spaceAfterMultiplier", out var multiplier);
            if (hasPoints && hasMultiplier) throw Unsupported("paragraph", "spaceAfter and spaceAfterMultiplier are mutually exclusive");
            if (hasPoints) target.SpaceAfterPoints = Math.Round(before.GetDouble() * 100) / 100d;
            else if (hasMultiplier) target.SpaceAfterMultiplier = Math.Round(multiplier.GetDouble() * 100_000) / 100_000d;
            if (hasPoints || hasMultiplier) break;
        }
        foreach (var layer in new[] { direct, inline, middle, named })
        {
            if (layer is not { ValueKind: JsonValueKind.Object } paragraph) continue;
            var hasPoints = paragraph.TryGetProperty("lineSpacing", out var spacing);
            var hasMultiplier = paragraph.TryGetProperty("lineSpacingMultiplier", out var multiplier);
            if (hasPoints && hasMultiplier) throw Unsupported("paragraph", "lineSpacing and lineSpacingMultiplier are mutually exclusive");
            if (hasPoints) target.LineSpacingPoints = Math.Round(spacing.GetDouble() * 100) / 100d;
            else if (hasMultiplier) target.LineSpacingMultiplier = Math.Round(multiplier.GetDouble() * 100_000) / 100_000d;
            if (hasPoints || hasMultiplier) break;
        }
        foreach (var layer in new[] { direct, inline, middle, named })
        {
            if (layer is not { ValueKind: JsonValueKind.Object } paragraph) continue;
            var hasTabs = paragraph.TryGetProperty("tabStops", out var tabStops);
            var hasClear = paragraph.TryGetProperty("noTabStops", out var noTabStops);
            if (hasTabs && hasClear) throw Unsupported("paragraph", "tabStops and noTabStops are mutually exclusive");
            if (hasClear)
            {
                if (noTabStops.ValueKind != JsonValueKind.True)
                    throw Unsupported("paragraph", "noTabStops must be true");
                target.NoTabStops = true;
            }
            if (hasTabs)
            {
                if (tabStops.ValueKind != JsonValueKind.Array)
                    throw Unsupported("paragraph", "tabStops must be an array");
                foreach (var tab in tabStops.EnumerateArray())
                {
                    var position = tab.GetProperty("position").GetDouble();
                    if (!double.IsFinite(position) || position < 0 || position > int.MaxValue / (double)EmuPerPoint)
                        throw Unsupported("paragraph", "tab stop position must be in the nonnegative signed-32-bit EMU range");
                    target.TabStops.Add(new PresentationTabStop
                    {
                        PositionEmu = Emu(position),
                        Alignment = OptionalString(tab, "alignment") ?? "left",
                    });
                }
                if (target.TabStops.Count == 0) target.NoTabStops = true;
            }
            if (hasTabs || hasClear) break;
        }
        if (FirstProperty(direct, inline, middle, named, "defaultText") is { } defaultText && defaultText.EnumerateObject().Any())
            target.DefaultRunProperties = BuildTextStyle(defaultText, catalog, paragraphDefaults: true);
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
            else if (kind == "picture")
            {
                if (bullet.TryGetProperty("asset", out var asset))
                    target.PictureBullet = new PresentationPictureBullet { AssetId = catalog.NativeAssetId(asset.GetString()!) };
                else if (bullet.TryGetProperty("uri", out var uri))
                    target.PictureBullet = new PresentationPictureBullet { Uri = uri.GetString()! };
                else
                    throw Unsupported("text", "picture bullet requires an asset or uri");
            }
            if (bullet.TryGetProperty("fontFamily", out var bulletFont))
            {
                if (bullet.TryGetProperty("fontFollowText", out _))
                    throw Unsupported("paragraph", "bullet fontFamily and fontFollowText are mutually exclusive");
                target.BulletFontFamily = bulletFont.GetString()!;
            }
            if (bullet.TryGetProperty("fontFollowText", out var bulletFontFollowText)) target.BulletFontFollowText = bulletFontFollowText.GetBoolean();
            if (bullet.TryGetProperty("color", out var bulletColor))
            {
                if (bullet.TryGetProperty("colorFollowText", out _))
                    throw Unsupported("paragraph", "bullet color and colorFollowText are mutually exclusive");
                SetBulletColor(target, bulletColor, catalog);
            }
            if (bullet.TryGetProperty("colorFollowText", out var colorFollowText)) target.BulletColorFollowText = colorFollowText.GetBoolean();
            if (new[] { "size", "sizePercent", "sizeFollowText" }.Count(field => bullet.TryGetProperty(field, out _)) > 1)
                throw Unsupported("paragraph", "bullet size, sizePercent and sizeFollowText are mutually exclusive");
            if (bullet.TryGetProperty("size", out var bulletSize)) target.BulletSizePoints = QuantizedBulletSize(bulletSize, 1, 768, 100);
            if (bullet.TryGetProperty("sizePercent", out var bulletSizePercent)) target.BulletSizePercent = QuantizedBulletSize(bulletSizePercent, 0.25, 4, 100_000);
            if (bullet.TryGetProperty("sizeFollowText", out var sizeFollowText)) target.BulletSizeFollowText = sizeFollowText.GetBoolean();
        }
    }

    private static double QuantizedBulletSize(JsonElement value, double minimum, double maximum, int units)
    {
        var size = value.GetDouble();
        if (!double.IsFinite(size) || size < minimum || size > maximum)
            throw Unsupported("paragraph", $"bullet size must be between {minimum} and {maximum}");
        // Requested wire state must match the native integer precision used by
        // source-preserving writeback and its post-export semantic check.
        return Math.Round(size * units) / units;
    }

    private static void SetBulletColor(PresentationTextParagraph target, JsonElement value, Catalog catalog)
    {
        var alpha = 1d;
        var explicitAlpha = value.ValueKind == JsonValueKind.String
            ? value.GetString()!.TrimStart('#').Length == 8
            : value.TryGetProperty("alpha", out _);
        if (value.ValueKind == JsonValueKind.Object && value.TryGetProperty("rgb", out var rgb))
        {
            target.BulletColorRgb = ParseHexColor(rgb.GetString()!).Rgb;
            if (value.TryGetProperty("alpha", out var opacity)) alpha = opacity.GetDouble();
        }
        else if (value.ValueKind == JsonValueKind.Object && value.TryGetProperty("token", out var token) &&
            !catalog.HasGrammarToken(token.GetString()!) && !value.TryGetProperty("tint", out _) && !value.TryGetProperty("shade", out _) &&
            PptxColor.TrySchemeToken(token.GetString()!, out var scheme))
        {
            target.BulletColorScheme = scheme;
            if (value.TryGetProperty("alpha", out var opacity)) alpha = opacity.GetDouble();
        }
        else
        {
            var color = catalog.BulletColor(value);
            target.BulletColorRgb = color.Rgb;
            alpha = color.Alpha;
        }
        if (explicitAlpha || alpha != 1) target.BulletColorOpacityThousandthPercent = Opacity(alpha);
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

    internal static PresentationBackground BuildBackground(
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
                    "background",
                    opacity => catalog.NumberToken(opacity, "opacity", "background opacity"),
                    fit => catalog.StringToken(fit, "string", "background fit")),
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
        if (source.Repeat is { } repeat) animation.RepeatCount = checked((uint)repeat);
        if (source.AutoReverse is { } autoReverse) animation.AutoReverse = autoReverse;
        if (!string.IsNullOrEmpty(source.Easing)) animation.Easing = source.Easing;
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

    private static void AddComments(
        PresentationArtifact target,
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjExpandedPageModel> expandedByPage)
    {
        if (!program.Root.TryGetProperty("comments", out var rawComments)) return;
        var pageById = target.Slides.ToDictionary(page => page.Id, StringComparer.Ordinal);
        var roots = new Dictionary<(string PageId, string CommentId), PresentationModernCommentThread>();
        var pendingReplies = new List<(JsonElement Raw, string PageId, string ParentId)>();
        foreach (var raw in rawComments.EnumerateArray())
        {
            var pageId = raw.GetProperty("page").GetString()!;
            if (!pageById.TryGetValue(pageId, out var page))
                throw Unsupported("$.comments", $"comment page {pageId} does not exist");
            var kind = OptionalString(raw, "kind") ?? "legacy";
            if (kind == "modern")
            {
                var parentId = OptionalString(raw, "parent");
                if (parentId is not null)
                {
                    pendingReplies.Add((raw, pageId, parentId));
                    continue;
                }

                var thread = BuildSourceFreeModernCommentThread(raw, pageId, program, expandedByPage);
                page.ModernComments.Add(thread);
                roots.Add((pageId, raw.GetProperty("id").GetString()!), thread);
                continue;
            }
            if (kind != "legacy")
                throw Unsupported("$.comments", $"comment kind {kind} is outside the bounded legacy/modern profile");
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
            page.LegacyComments.Add(comment);
        }

        foreach (var (raw, pageId, parentId) in pendingReplies)
        {
            if (!roots.TryGetValue((pageId, parentId), out var thread))
                throw Unsupported("$.comments", $"modern reply {raw.GetProperty("id").GetString()} does not resolve to a root on page {pageId}");
                thread.Replies.Add(BuildSourceFreeModernComment(raw, program.Meta.Id));
        }
    }

    private static PresentationModernCommentThread BuildSourceFreeModernCommentThread(
        JsonElement raw,
        string pageId,
        PpjProgramModel program,
        IReadOnlyDictionary<string, PpjExpandedPageModel> expandedByPage)
    {
        var targetId = raw.GetProperty("target").GetString()!;
        var lookupId = targetId.EndsWith("/text", StringComparison.Ordinal)
            ? targetId[..^5]
            : targetId;
        if (!expandedByPage.TryGetValue(pageId, out var expanded))
            throw Unsupported("$.comments", $"comment page {pageId} does not exist");
        var flattened = Walk(expanded.Elements).ToArray();
        var targetIndex = Array.FindIndex(flattened, element => element.Id.Equals(lookupId, StringComparison.Ordinal));
        if (targetIndex < 0)
            throw Unsupported("$.comments", $"modern comment target {targetId} does not resolve to an authored element");

        var anchor = raw.GetProperty("anchor");
        var anchorKind = anchor.GetProperty("kind").GetString()!;
        var expectedMoniker = ModernCommentMoniker(flattened[targetIndex]);
        if (!anchor.GetProperty("moniker").GetString()!.Equals(expectedMoniker, StringComparison.Ordinal))
            throw Unsupported("$.comments", $"modern comment target {targetId} requires anchor moniker {expectedMoniker}");
        var nativeSlideIndex = program.Pages
            .Select((page, index) => (page.Id, index))
            .Single(item => item.Id.Equals(pageId, StringComparison.Ordinal)).index;
        var thread = new PresentationModernCommentThread
        {
            Id = ModernCommentGuid(program.Meta.Id, raw.GetProperty("id").GetString()!),
            TargetId = targetId,
            PositionXEmu = raw.TryGetProperty("position", out var position) ? Emu(position.GetProperty("x").GetDouble()) : 0,
            PositionYEmu = raw.TryGetProperty("position", out position) ? Emu(position.GetProperty("y").GetDouble()) : 0,
            Anchor = new PresentationModernCommentAnchor
            {
                Kind = anchorKind == "textRange"
                    ? PresentationModernCommentAnchor.Types.Kind.TextRange
                    : PresentationModernCommentAnchor.Types.Kind.Element,
                NativeSlideId = checked((uint)(256 + nativeSlideIndex)),
            },
            Root = BuildSourceFreeModernComment(raw, program.Meta.Id),
        };
        thread.Anchor.Monikers.Add(new PresentationModernCommentMoniker
        {
            Type = expectedMoniker,
            NativeId = checked((uint)(targetIndex + 2)),
        });
        if (anchorKind == "textRange")
        {
            if (flattened[targetIndex] is not PpjTextElementModel and not PpjShapeElementModel and not PpjPlaceholderElementModel)
                throw Unsupported("$.comments", $"modern text-range target {targetId} is not a text shape");
            thread.Anchor.TextStart = checked((uint)anchor.GetProperty("textStart").GetInt64());
            thread.Anchor.TextLength = checked((uint)anchor.GetProperty("textLength").GetInt64());
            if (anchor.TryGetProperty("contextLength", out var contextLength))
                thread.Anchor.ContextLength = checked((uint)contextLength.GetInt64());
            if (anchor.TryGetProperty("contextHash", out var contextHash))
                thread.Anchor.ContextHash = checked((uint)contextHash.GetInt64());
        }
        else if (anchorKind != "element")
            throw Unsupported("$.comments", $"modern comment anchor kind {anchorKind} is unsupported");
        return thread;
    }

    private static PresentationModernComment BuildSourceFreeModernComment(JsonElement raw, string programId)
    {
        var id = raw.GetProperty("id").GetString()!;
        var author = raw.GetProperty("author").GetString()!;
        var authorId = ModernAuthorGuid(programId, author);
        return new PresentationModernComment
        {
            Id = ModernCommentGuid(programId, id),
            AuthorId = authorId,
            Author = author,
            Initials = ModernAuthorInitials(author),
            UserId = $"ppj:{authorId}",
            ProviderId = "OfficeKit",
            Text = raw.GetProperty("text").GetString()!,
            CreatedAt = raw.GetProperty("createdAt").GetString()!,
            Status = raw.GetProperty("status").GetString()!,
        };
    }

    private static string ModernCommentGuid(string programId, string id) => DeterministicGuid(programId, $"modern-comment\0{id}");

    private static string ModernAuthorGuid(string programId, string author) => DeterministicGuid(programId, $"modern-author\0{author}");

    private static string ModernAuthorInitials(string author)
    {
        var initials = string.Concat(author
            .Split(new[] { ' ', '\t', '-', '_', '.' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part[0])
            .Take(3)
            .Select(char.ToUpperInvariant));
        return initials.Length > 0 ? initials : "A";
    }

    private static string ModernCommentMoniker(PpjElementModel element) => element.Type switch
    {
        "image" or "media" => "picMk",
        "chart" or "table" or "smartArt" => "graphicFrameMk",
        "connector" => "cxnSpMk",
        "group" => "grpSpMk",
        _ => "spMk",
    };

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

    private static void ApplyImageCrop(
        PresentationImage target,
        PpjImageElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var source = EffectiveImagePaintSource(element, raw, namedStyle, inlineStyle, catalog);
        var paint = PpjImagePaintLowering.Build(
            source,
            element.Frame.Width,
            element.Frame.Height,
            catalog.NativeAssetId,
            catalog.AssetDimensions,
            $"element {element.Id}",
            opacity => catalog.NumberToken(opacity, "opacity", $"element {element.Id} opacity"),
            fit => catalog.StringToken(fit, "string", $"element {element.Id} fit"));
        target.Crop = paint.Crop;
        target.Tiled = paint.Mode == PresentationImagePaint.Types.Mode.Tile;
        if (paint.HasOpacityThousandthPercent)
            target.OpacityThousandthPercent = paint.OpacityThousandthPercent;
    }

    private static JsonElement EffectiveImagePaintSource(
        PpjImageElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var output = new JsonObject
        {
            ["asset"] = JsonValue.Create(element.AssetId),
        };
        foreach (var field in new[] { "fit", "crop", "focus", "opacity" })
        {
            if (catalog.PropertyByPrecedence(
                    $"image.{field}",
                    raw,
                    inlineStyle,
                    namedStyle,
                    includeElementWhenUndeclared: true) is not { } value)
                continue;
            output[field] = JsonNode.Parse(value.GetRawText());
        }
        using var document = JsonDocument.Parse(output.ToJsonString());
        return document.RootElement.Clone();
    }

    private static void ApplyChartStyle(
        PresentationChart chart,
        JsonElement? named,
        JsonElement? inline,
        Catalog catalog,
        string elementId,
        double frameWidth,
        double frameHeight,
        JsonElement? element = null)
    {
        var legendValue = catalog.PropertyByPrecedence("chart.legend", element, inline, named);
        var legend = legendValue is { } explicitLegend
            ? ChartEnumToken(explicitLegend, catalog, "chart legend", "none", "top", "topRight", "bottom", "left", "right")
            : null;
        chart.HasLegend = legend is not null and not "none";
        if (chart.HasLegend) chart.LegendPosition = legend!;
        if (catalog.PropertyByPrecedence("chart.legendOverlay", element, inline, named) is { } legendOverlay)
        {
            if (!chart.HasLegend)
                throw Unsupported(elementId, "legendOverlay requires a visible legend");
            chart.LegendOverlay = catalog.BooleanToken(legendOverlay, "boolean", "chart legendOverlay");
        }
        if (catalog.PropertyByPrecedence("chart.legendFill", element, inline, named) is { } legendFill)
        {
            if (!chart.HasLegend)
                throw Unsupported(elementId, "legendFill requires a visible legend");
            chart.LegendFill = BuildChartSurfaceFill(legendFill, catalog, $"{elementId} legend");
        }
        if (catalog.PropertyByPrecedence("chart.legendLine", element, inline, named) is { } legendLine)
        {
            if (!chart.HasLegend)
                throw Unsupported(elementId, "legendLine requires a visible legend");
            chart.LegendLine = ChartCompiler.BuildChartLine(legendLine, catalog);
        }
        if (catalog.PropertyByPrecedence("chart.stacking", element, inline, named) is { } stacking)
        {
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area or SpreadsheetChartType.Combo))
                throw Unsupported(elementId, "stacking applies only to bar, column, line, area, and combo charts");
            chart.Grouping = ChartEnumToken(stacking, catalog, "chart stacking", "none", "stacked", "percent-stacked", "stream");
        }
        if (catalog.PropertyByPrecedence("chart.gapWidth", element, inline, named) is { } gapWidth)
        {
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Combo))
                throw Unsupported(elementId, "gapWidth applies only to bar, column, and combo charts");
            chart.GapWidth = ChartIntegerToken(gapWidth, catalog, "chart gapWidth", 0, 500);
        }
        if (catalog.PropertyByPrecedence("chart.overlap", element, inline, named) is { } overlap)
        {
            if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Combo))
                throw Unsupported(elementId, "overlap applies only to bar, column, and combo charts");
            chart.Overlap = ChartSignedIntegerToken(overlap, catalog, "chart overlap", -100, 100);
        }
        if (catalog.PropertyByPrecedence("chart.startAngle", element, inline, named) is { } startAngle)
        {
            if (chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
                throw Unsupported(elementId, "startAngle applies only to pie and doughnut charts");
            chart.FirstSliceAngle = ChartIntegerToken(startAngle, catalog, "chart startAngle", 0, 360);
        }
        if (catalog.PropertyByPrecedence("chart.holeSize", element, inline, named) is { } holeSize)
        {
            if (chart.Type != SpreadsheetChartType.Doughnut)
                throw Unsupported(elementId, "holeSize applies only to doughnut charts");
            chart.DoughnutHoleSize = ChartIntegerToken(holeSize, catalog, "chart holeSize", 10, 90);
        }
        if (catalog.PropertyByPrecedence("chart.bubbleScale", element, inline, named) is { } bubbleScale)
        {
            if (chart.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(elementId, "bubbleScale applies only to bubble charts");
            chart.BubbleScale = ChartIntegerToken(bubbleScale, catalog, "chart bubbleScale", 0, 300);
        }
        if (catalog.PropertyByPrecedence("chart.bubbleSizeMode", element, inline, named) is { } bubbleSizeMode)
        {
            if (chart.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(elementId, "bubbleSizeMode applies only to bubble charts");
            chart.BubbleSizeMode = ChartEnumToken(bubbleSizeMode, catalog, "chart bubbleSizeMode", "area", "width");
        }
        if (catalog.PropertyByPrecedence("chart.scatterStyle", element, inline, named) is { } scatterStyle)
        {
            if (chart.Type != SpreadsheetChartType.Scatter)
                throw Unsupported(elementId, "scatterStyle applies only to scatter charts");
            chart.ScatterStyle = ChartEnumToken(scatterStyle, catalog, "chart scatterStyle", "line", "lineWithMarkers", "marker", "smooth", "smoothWithMarkers");
        }
        if (catalog.PropertyByPrecedence("chart.radarStyle", element, inline, named) is { } radarStyle)
        {
            if (chart.Type != SpreadsheetChartType.Radar)
                throw Unsupported(elementId, "radarStyle applies only to radar charts");
            chart.RadarStyle = ChartEnumToken(radarStyle, catalog, "chart radarStyle", "standard", "marker", "filled");
        }
        var axisBearing = chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut);
        if (catalog.PropertyByPrecedence("chart.showCategoryAxis", element, inline, named) is { } showCategoryAxis)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have category axes");
            chart.ShowCategoryAxis = catalog.BooleanToken(showCategoryAxis, "boolean", "chart showCategoryAxis");
        }
        if (catalog.PropertyByPrecedence("chart.showValueAxis", element, inline, named) is { } showValueAxis)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have value axes");
            chart.ShowValueAxis = catalog.BooleanToken(showValueAxis, "boolean", "chart showValueAxis");
        }
        if (catalog.PropertyByPrecedence("chart.showGridlines", element, inline, named) is { } showGridlines)
        {
            if (!axisBearing) throw Unsupported(elementId, "circular charts do not have value-axis gridlines");
            // DrawingML represents "no major gridlines" by omitting the node.
            // Keep false canonical rather than inventing a presence-only wire
            // distinction that the package cannot round-trip.
            if (catalog.BooleanToken(showGridlines, "boolean", "chart showGridlines")) chart.ShowGridlines = true;
        }
        if (catalog.PropertyByPrecedence("chart.chartAreaFill", element, inline, named) is { } chartAreaFill)
            chart.ChartAreaFill = BuildChartSurfaceFill(chartAreaFill, catalog, $"{elementId} chart area");
        if (catalog.PropertyByPrecedence("chart.plotAreaFill", element, inline, named) is { } plotAreaFill)
            chart.PlotAreaFill = BuildChartSurfaceFill(plotAreaFill, catalog, $"{elementId} plot area");
        if (catalog.PropertyByPrecedence("chart.plotAreaLine", element, inline, named) is { } plotAreaLine)
            chart.PlotAreaLine = ChartCompiler.BuildChartLine(plotAreaLine, catalog);
        if (catalog.PropertyByPrecedence("chart.frame", element, inline, named) is { } frame)
        {
            chart.Frame = new PresentationChartFrame();
            if (frame.TryGetProperty("fill", out var frameFill))
            {
                if (frameFill.TryGetProperty("type", out var frameFillType) && frameFillType.GetString() == "image")
                    chart.Frame.ImageFill = PpjImagePaintLowering.Build(
                        frameFill,
                        frameWidth,
                        frameHeight,
                        catalog.NativeAssetId,
                        catalog.AssetDimensions,
                        $"{elementId} chart frame",
                        opacity => catalog.NumberToken(opacity, "opacity", $"{elementId} chart frame opacity"),
                        fit => catalog.StringToken(fit, "string", $"{elementId} chart frame fit"));
                else
                    chart.Frame.Fill = BuildChartSurfaceFill(frameFill, catalog, $"{elementId} chart frame");
            }
            if (frame.TryGetProperty("stroke", out var frameStroke))
                chart.Frame.Line = ChartCompiler.BuildChartLine(frameStroke, catalog);
            if (frame.TryGetProperty("shadow", out var frameShadow))
                chart.Frame.Shadow = BuildShadow(frameShadow, catalog);
        }
        if (catalog.PropertyByPrecedence("chart.titleTextStyle", element, inline, named) is { } titleTextStyle)
        {
            if (chart.Title.Length == 0)
                throw Unsupported(elementId, "titleTextStyle requires a non-empty chart title");
            chart.TitleTextStyle = BuildChartTextStyle(
                "chart.titleTextStyle", element, inline, named, titleTextStyle, catalog);
        }
        if (catalog.PropertyByPrecedence("chart.legendTextStyle", element, inline, named) is { } legendTextStyle)
        {
            if (!chart.HasLegend)
                throw Unsupported(elementId, "legendTextStyle requires a visible legend");
            chart.LegendTextStyle = BuildChartTextStyle(
                "chart.legendTextStyle", element, inline, named, legendTextStyle, catalog);
        }
        var smooth = catalog.PropertyByPrecedence("chart.smooth", element, inline, named);
        var varyColors = catalog.PropertyByPrecedence("chart.varyColors", element, inline, named);
        if (smooth is not null && chart.Type != SpreadsheetChartType.Line)
            throw Unsupported(elementId, "smooth applies only to line charts");
        if (varyColors is not null && chart.Type is not (SpreadsheetChartType.Line or SpreadsheetChartType.Bar or SpreadsheetChartType.Area or SpreadsheetChartType.Scatter or SpreadsheetChartType.Combo))
            throw Unsupported(elementId, "varyColors applies only to line, column, area, scatter, and categorical combo charts");
        if (smooth is { } explicitSmooth)
        {
            chart.LineOptions = new SpreadsheetChartLineOptionsArtifact
            {
                Smooth = catalog.BooleanToken(explicitSmooth, "boolean", "chart smooth"),
            };
        }
        if (varyColors is { } explicitVaryColors)
        {
            var resolvedVaryColors = catalog.BooleanToken(explicitVaryColors, "boolean", "chart varyColors");
            if (chart.Type == SpreadsheetChartType.Line)
            {
                if (resolvedVaryColors)
                    (chart.LineOptions ??= new SpreadsheetChartLineOptionsArtifact()).VaryColors = true;
            }
            else
                chart.VaryColors = resolvedVaryColors;
        }
        var structuredLabels = catalog.PropertyByPrecedence("chart.dataLabels", element, inline, named);
        var labels = catalog.PropertyByPrecedence("chart.showDataLabels", element, inline, named);
        var legacyPosition = catalog.PropertyByPrecedence("chart.dataLabelPosition", element, inline, named);
        var showLabels = labels is { } explicitLabels
            ? catalog.BooleanToken(explicitLabels, "boolean", "chart showDataLabels")
            : (bool?)null;
        var labelPosition = legacyPosition is { } explicitPosition
            ? ChartEnumToken(explicitPosition, catalog, "chart dataLabelPosition", "center", "inside-end", "outside-end", "above", "below")
            : null;
        if (structuredLabels is not null && (labels is not null || legacyPosition is not null))
            throw Unsupported(elementId, "structured dataLabels cannot be combined with legacy data-label fields");
        if (structuredLabels is { } dataLabels)
        {
            chart.DataLabels = new SpreadsheetChartDataLabelsArtifact
            {
                ShowValue = dataLabels.TryGetProperty("showValue", out var showValue) &&
                    catalog.BooleanToken(showValue, "boolean", "chart dataLabels showValue"),
                ShowCategoryName = dataLabels.TryGetProperty("showCategory", out var showCategory) &&
                    catalog.BooleanToken(showCategory, "boolean", "chart dataLabels showCategory"),
            };
            if (dataLabels.TryGetProperty("showSeries", out var showSeries))
                chart.DataLabels.ShowSeriesName = catalog.BooleanToken(showSeries, "boolean", "chart dataLabels showSeries");
            if (dataLabels.TryGetProperty("showPercent", out var showPercent))
                chart.DataLabels.ShowPercent = catalog.BooleanToken(showPercent, "boolean", "chart dataLabels showPercent");
            if (dataLabels.TryGetProperty("showBubbleSize", out var showBubbleSize))
                chart.DataLabels.ShowBubbleSize = catalog.BooleanToken(showBubbleSize, "boolean", "chart dataLabels showBubbleSize");
            if (dataLabels.TryGetProperty("showLeaderLines", out var showLeaderLines))
                chart.DataLabels.ShowLeaderLines = catalog.BooleanToken(showLeaderLines, "boolean", "chart dataLabels showLeaderLines");
            if (dataLabels.TryGetProperty("position", out var position))
                chart.DataLabels.Position = LabelPosition(ChartEnumToken(position, catalog, "chart dataLabels position",
                    "best-fit", "bottom", "center", "inside-base", "inside-end", "left", "outside-end", "right", "top"));
            if (dataLabels.TryGetProperty("numberFormat", out var numberFormat))
                chart.DataLabels.NumberFormatCode = catalog.StringToken(
                    numberFormat,
                    "string",
                    "chart dataLabels numberFormat");
            if (dataLabels.TryGetProperty("textStyle", out var textStyle))
                chart.DataLabels.TextStyle = BuildChartTextStyle(
                    "chart.dataLabels.textStyle", element, inline, named, textStyle, catalog);
            if (dataLabels.TryGetProperty("fill", out var fill))
                chart.DataLabels.Fill = BuildChartSurfaceFill(fill, catalog, "chart dataLabels fill");
            if (dataLabels.TryGetProperty("line", out var line))
                chart.DataLabels.Line = ChartCompiler.BuildChartLine(line, catalog);
        }
        else if (showLabels == true)
        {
            chart.DataLabels = new SpreadsheetChartDataLabelsArtifact
            {
                ShowValue = true,
                ShowSeriesName = false,
            };
            if (labelPosition is not null) chart.DataLabels.Position = LabelPosition(labelPosition);
        }
        else if (legacyPosition is not null)
            throw Unsupported("chart", "dataLabelPosition requires showDataLabels: true");
    }

    private static string ChartEnumToken(JsonElement value, Catalog catalog, string owner, params string[] allowed)
    {
        var resolved = catalog.StringToken(value, "string", owner);
        if (!allowed.Contains(resolved, StringComparer.Ordinal))
            throw Unsupported(owner, $"value {resolved} is outside the bounded chart vocabulary");
        return resolved;
    }

    private static uint ChartIntegerToken(JsonElement value, Catalog catalog, string owner, int minimum, int maximum)
    {
        var resolved = catalog.NumberToken(value, "size", owner);
        if (!double.IsFinite(resolved) || resolved != Math.Truncate(resolved) || resolved < minimum || resolved > maximum)
            throw Unsupported(owner, $"value must resolve to an integer from {minimum} through {maximum}");
        return checked((uint)resolved);
    }

    private static int ChartSignedIntegerToken(JsonElement value, Catalog catalog, string owner, int minimum, int maximum)
    {
        var resolved = catalog.NumberToken(value, "size", owner);
        if (!double.IsFinite(resolved) || resolved != Math.Truncate(resolved) || resolved < minimum || resolved > maximum)
            throw Unsupported(owner, $"value must resolve to an integer from {minimum} through {maximum}");
        return checked((int)resolved);
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
        if (source.TryGetProperty("fontSize", out var fontSize)) output.FontSizePoints = catalog.PositiveNumberToken(fontSize, "size", "chart text fontSize");
        if (source.TryGetProperty("fontFamily", out var fontFamily)) output.FontFamily = catalog.StringToken(fontFamily, "string", "chart text fontFamily");
        if (source.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) output.FontFamilyEastAsia = catalog.StringToken(eastAsia, "string", "chart text fontFamilyEastAsia");
        if (source.TryGetProperty("fontFamilyComplexScript", out var complexScript)) output.FontFamilyComplexScript = catalog.StringToken(complexScript, "string", "chart text fontFamilyComplexScript");
        if (source.TryGetProperty("language", out var language)) output.Language = catalog.LanguageTagToken(language, "chart text language");
        if (source.TryGetProperty("strike", out var strike)) output.Strike = NativeStrike(strike);
        if (source.TryGetProperty("baseline", out var baseline)) output.BaselineThousandthPercent = XlsxChartTextStyleCodec.BaselineThousandthPercent(baseline.GetDouble());
        if (source.TryGetProperty("letterSpacing", out var spacing)) output.LetterSpacingHundredthPoints = XlsxChartTextStyleCodec.LetterSpacingHundredthPoints(spacing.GetDouble());
        if (source.TryGetProperty("kerning", out var kerning)) output.KerningHundredthPoints = XlsxChartTextStyleCodec.KerningHundredthPoints(kerning.GetDouble());
        if (source.TryGetProperty("capitalization", out var capitalization)) output.Capitalization = capitalization.GetString()!;
        if (source.TryGetProperty("softEdge", out var softEdge)) output.SoftEdge = BuildChartTextSoftEdge(softEdge);
        if (source.TryGetProperty("reflection", out var reflection)) output.Reflection = BuildChartTextReflection(reflection, catalog);
        if (source.TryGetProperty("innerShadow", out var innerShadow)) output.InnerShadow = BuildChartTextInnerShadow(innerShadow, catalog);
        if (source.TryGetProperty("glow", out var glow)) output.Glow = BuildChartTextGlow(glow, catalog);
        if (source.TryGetProperty("shadow", out var shadow)) output.Shadow = BuildChartTextShadow(shadow, catalog);
        if (source.TryGetProperty("highlight", out var highlight)) output.HighlightRgb = BuildChartHighlight(highlight, catalog);
        if (source.TryGetProperty("bold", out var bold)) output.Bold = catalog.BooleanToken(bold, "boolean", "chart text bold");
        if (source.TryGetProperty("italic", out var italic)) output.Italic = catalog.BooleanToken(italic, "boolean", "chart text italic");
        if (source.TryGetProperty("underline", out var underline)) output.Underline = NativeUnderline(underline.GetString()!);
        if (source.TryGetProperty("alignment", out var alignment)) output.Alignment = NativeChartAlignment(alignment.GetString()!);
        if (source.TryGetProperty("fill", out var fill))
            ApplyChartTextFill(output, BuildChartSurfaceFill(fill, catalog, "chart text fill"));
        if (source.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            output.Fill = null;
            output.ColorRgb = resolved.Rgb;
            if (resolved.Alpha == 1) output.ClearOpacityThousandthPercent();
            else output.OpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        return output;
    }

    private static SpreadsheetChartTextStyleArtifact BuildChartTextStyle(
        string target,
        JsonElement? element,
        JsonElement? inline,
        JsonElement? named,
        JsonElement fallback,
        Catalog catalog)
    {
        // Preserve the historical whole-object precedence unless the grammar
        // explicitly declares a nested field rule. This keeps an existing
        // `chart.legendTextStyle` declaration source-stable while allowing a
        // bounded Kimi-style shallow merge such as
        // `chart.legendTextStyle.fontFamily` from inline and
        // `chart.legendTextStyle.fontSize` from styleRef.
        var output = BuildChartTextStyle(fallback, catalog);
        foreach (var field in ChartTextStyleFields)
        {
            var nestedTarget = $"{target}.{field}";
            if (!catalog.HasStylePrecedence(nestedTarget)) continue;
            if (catalog.PropertyByPrecedence(nestedTarget, element, inline, named) is { } value)
                ApplyChartTextStyleProperty(output, field, value, catalog);
        }
        return output;
    }

    private static readonly string[] ChartTextStyleFields =
    ["fontSize", "fontFamily", "fontFamilyEastAsia", "fontFamilyComplexScript", "language", "strike", "baseline", "capitalization", "letterSpacing", "kerning", "highlight", "shadow", "glow", "innerShadow", "reflection", "softEdge", "bold", "italic", "underline", "alignment", "fill", "color"];

    private static void ApplyChartTextStyleProperty(
        SpreadsheetChartTextStyleArtifact output,
        string field,
        JsonElement value,
        Catalog catalog)
    {
        switch (field)
        {
            case "fontSize":
                output.FontSizePoints = catalog.PositiveNumberToken(value, "size", "chart text fontSize");
                break;
            case "fontFamily":
                output.FontFamily = catalog.StringToken(value, "string", "chart text fontFamily");
                break;
            case "fontFamilyEastAsia":
                output.FontFamilyEastAsia = catalog.StringToken(value, "string", "chart text fontFamilyEastAsia");
                break;
            case "fontFamilyComplexScript":
                output.FontFamilyComplexScript = catalog.StringToken(value, "string", "chart text fontFamilyComplexScript");
                break;
            case "language":
                output.Language = catalog.LanguageTagToken(value, "chart text language");
                break;
            case "strike":
                output.Strike = NativeStrike(value);
                break;
            case "baseline":
                output.BaselineThousandthPercent = XlsxChartTextStyleCodec.BaselineThousandthPercent(value.GetDouble());
                break;
            case "letterSpacing":
                output.LetterSpacingHundredthPoints = XlsxChartTextStyleCodec.LetterSpacingHundredthPoints(value.GetDouble());
                break;
            case "kerning":
                output.KerningHundredthPoints = XlsxChartTextStyleCodec.KerningHundredthPoints(value.GetDouble());
                break;
            case "reflection":
                output.Reflection = BuildChartTextReflection(value, catalog);
                break;
            case "innerShadow":
                output.InnerShadow = BuildChartTextInnerShadow(value, catalog);
                break;
            case "softEdge":
                output.SoftEdge = BuildChartTextSoftEdge(value);
                break;
            case "glow":
                output.Glow = BuildChartTextGlow(value, catalog);
                break;
            case "shadow":
                output.Shadow = BuildChartTextShadow(value, catalog);
                break;
            case "highlight":
                output.HighlightRgb = BuildChartHighlight(value, catalog);
                break;
            case "capitalization":
                output.Capitalization = value.GetString()!;
                break;
            case "bold":
                output.Bold = catalog.BooleanToken(value, "boolean", "chart text bold");
                break;
            case "italic":
                output.Italic = catalog.BooleanToken(value, "boolean", "chart text italic");
                break;
            case "underline":
                output.Underline = NativeUnderline(value.GetString()!);
                break;
            case "alignment":
                output.Alignment = NativeChartAlignment(value.GetString()!);
                break;
            case "fill":
                ApplyChartTextFill(output, BuildChartSurfaceFill(value, catalog, "chart text fill"));
                break;
            case "color":
                var resolved = catalog.Color(value);
                output.Fill = null;
                output.ColorRgb = resolved.Rgb;
                if (resolved.Alpha == 1) output.ClearOpacityThousandthPercent();
                else output.OpacityThousandthPercent = Opacity(resolved.Alpha);
                break;
            default:
                throw new InvalidOperationException($"Unknown chart text style field {field}.");
        }
    }

    internal static PresentationSoftEdge BuildChartTextSoftEdge(JsonElement value)
    {
        var output = new PresentationSoftEdge
        {
            RadiusEmu = Emu(ChartEffectNumber(value.GetProperty("radius").GetDouble(), 0, 1000)),
        };
        PptxSoftEdgeValueCodec.Validate(output, "chart-text", "chart text");
        return output;
    }

    private static PresentationReflection BuildChartTextReflection(JsonElement value, Catalog catalog) =>
        BuildChartTextReflection(value, opacity => catalog.NumberToken(opacity, "opacity", "chart text reflection opacity"));

    internal static PresentationReflection BuildChartTextReflection(JsonElement value, Func<JsonElement, double> resolveOpacity)
    {
        var output = new PresentationReflection();
        if (value.TryGetProperty("fadeAngle", out var fadeAngle))
            output.FadeDirectionAngle60000 = Angle(((ChartEffectNumber(fadeAngle.GetDouble(), -360, 360) % 360) + 360) % 360) % 21_600_000;
        if (value.TryGetProperty("scaleX", out var scaleX)) output.ScaleXThousandthPercent = checked((int)Math.Round(ChartEffectNumber(scaleX.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d, MidpointRounding.ToEven));
        if (value.TryGetProperty("scaleY", out var scaleY)) output.ScaleYThousandthPercent = checked((int)Math.Round(ChartEffectNumber(scaleY.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d, MidpointRounding.ToEven));
        if (value.TryGetProperty("skewX", out var skewX)) output.SkewXAngle60000 = checked((int)Angle(ChartEffectNumber(skewX.GetDouble(), -90, 90)));
        if (value.TryGetProperty("skewY", out var skewY)) output.SkewYAngle60000 = checked((int)Angle(ChartEffectNumber(skewY.GetDouble(), -90, 90)));
        if (value.TryGetProperty("alignment", out var alignment)) output.Alignment = alignment.GetString()!;
        if (value.TryGetProperty("rotateWithShape", out var rotateWithShape)) output.RotateWithShape = rotateWithShape.GetBoolean();
        if (value.TryGetProperty("startPosition", out var startPosition)) output.StartPositionThousandthPercent = Opacity(ChartEffectNumber(startPosition.GetDouble(), 0, 1));
        if (value.TryGetProperty("endPosition", out var endPosition)) output.EndPositionThousandthPercent = Opacity(ChartEffectNumber(endPosition.GetDouble(), 0, 1));
        if (value.TryGetProperty("blur", out var blur)) output.BlurRadiusEmu = Emu(ChartEffectNumber(blur.GetDouble(), 0, 1000));
        if (value.TryGetProperty("distance", out var distance)) output.DistanceEmu = Emu(ChartEffectNumber(distance.GetDouble(), 0, 100000));
        if (value.TryGetProperty("angle", out var angle))
            output.DirectionAngle60000 = Angle(((ChartEffectNumber(angle.GetDouble(), -360, 360) % 360) + 360) % 360) % 21_600_000;
        if (value.TryGetProperty("startOpacity", out var startOpacity)) output.StartOpacityThousandthPercent = Opacity(ChartEffectNumber(resolveOpacity(startOpacity), 0, 1));
        if (value.TryGetProperty("endOpacity", out var endOpacity)) output.EndOpacityThousandthPercent = Opacity(ChartEffectNumber(resolveOpacity(endOpacity), 0, 1));
        PptxReflectionCodec.Validate(output, "chart-text", "chart text");
        return output;
    }

    private static PresentationInnerShadow BuildChartTextInnerShadow(JsonElement value, Catalog catalog) =>
        BuildChartTextInnerShadow(value, catalog.Color,
            opacity => catalog.NumberToken(opacity, "opacity", "chart text inner shadow opacity"), catalog.HasGrammarToken);

    internal static PresentationInnerShadow BuildChartTextInnerShadow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var paint = BuildChartEffectColor(value, resolveColor, resolveOpacity, declaredToken);
        var output = new PresentationInnerShadow { ColorRgb = paint.Rgb };
        if (paint.Scheme is not null) output.ColorScheme = paint.Scheme;
        if (paint.Opacity is { } opacity) output.OpacityThousandthPercent = opacity;
        if (value.TryGetProperty("blur", out var blur)) output.BlurRadiusEmu = Emu(ChartEffectNumber(blur.GetDouble(), 0, 1000));
        if (value.TryGetProperty("distance", out var distance)) output.DistanceEmu = Emu(ChartEffectNumber(distance.GetDouble(), 0, 100000));
        if (value.TryGetProperty("angle", out var angle))
            output.DirectionAngle60000 = Angle(((ChartEffectNumber(angle.GetDouble(), -360, 360) % 360) + 360) % 360) % 21_600_000;
        PptxInnerShadowCodec.Validate(output, "chart-text", "chart text");
        return output;
    }

    private static PresentationShadow BuildChartTextShadow(JsonElement value, Catalog catalog) =>
        BuildChartTextShadow(value, catalog.Color,
            opacity => catalog.NumberToken(opacity, "opacity", "chart text shadow opacity"), catalog.HasGrammarToken);

    internal static PresentationShadow BuildChartTextShadow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var paint = BuildChartEffectColor(value, resolveColor, resolveOpacity, declaredToken);
        var output = new PresentationShadow { ColorRgb = paint.Rgb };
        if (value.TryGetProperty("scaleX", out var scaleX)) output.ScaleXThousandthPercent = checked((int)Math.Round(ChartEffectNumber(scaleX.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d, MidpointRounding.ToEven));
        if (value.TryGetProperty("scaleY", out var scaleY)) output.ScaleYThousandthPercent = checked((int)Math.Round(ChartEffectNumber(scaleY.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d, MidpointRounding.ToEven));
        if (value.TryGetProperty("skewX", out var skewX)) output.SkewXAngle60000 = Angle(ChartEffectNumber(skewX.GetDouble(), -90, 90));
        if (value.TryGetProperty("skewY", out var skewY)) output.SkewYAngle60000 = Angle(ChartEffectNumber(skewY.GetDouble(), -90, 90));
        if (paint.Scheme is not null) output.ColorScheme = paint.Scheme;
        if (paint.Opacity is { } opacity) output.OpacityThousandthPercent = opacity;
        if (value.TryGetProperty("blur", out var blur)) output.BlurRadiusEmu = Emu(ChartEffectNumber(blur.GetDouble(), 0, 1000));
        if (value.TryGetProperty("distance", out var distance)) output.DistanceEmu = Emu(ChartEffectNumber(distance.GetDouble(), 0, 100000));
        if (value.TryGetProperty("angle", out var angle))
            output.DirectionAngle60000 = Angle(((ChartEffectNumber(angle.GetDouble(), -360, 360) % 360) + 360) % 360) % 21_600_000;
        if (value.TryGetProperty("alignment", out var alignment)) output.Alignment = alignment.GetString()!;
        if (value.TryGetProperty("rotateWithShape", out var rotate)) output.RotateWithShape = rotate.GetBoolean();
        PptxShadowCodec.Validate(output, "chart-text", "chart text");
        return output;
    }

    private static PresentationGlow BuildChartTextGlow(JsonElement value, Catalog catalog) =>
        BuildChartTextGlow(value, catalog.Color,
            opacity => catalog.NumberToken(opacity, "opacity", "chart text glow opacity"), catalog.HasGrammarToken);

    internal static PresentationGlow BuildChartTextGlow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var paint = BuildChartEffectColor(value, resolveColor, resolveOpacity, declaredToken);
        var output = new PresentationGlow
        {
            ColorRgb = paint.Rgb,
            RadiusEmu = Emu(ChartEffectNumber(value.GetProperty("radius").GetDouble(), 0, 1000)),
        };
        if (paint.Scheme is not null) output.ColorScheme = paint.Scheme;
        if (paint.Opacity is { } opacity) output.OpacityThousandthPercent = opacity;
        PptxGlowCodec.Validate(output, "chart-text", "chart text");
        return output;
    }

    private static (string Rgb, string? Scheme, uint? Opacity) BuildChartEffectColor(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var color = value.GetProperty("color");
        var rgb = string.Empty;
        string? scheme = null;
        var alpha = 1d;
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            !declaredToken(token.GetString()!) && PptxColor.TrySchemeToken(token.GetString()!, out var recognizedScheme))
        {
            if (color.TryGetProperty("tint", out _) || color.TryGetProperty("shade", out _))
                throw new CodecException("invalid_presentation_effects", "Chart text theme effects do not support tint/shade transforms.");
            scheme = recognizedScheme;
            if (color.TryGetProperty("alpha", out var colorAlpha)) alpha = colorAlpha.GetDouble();
        }
        else
        {
            var resolved = resolveColor(color);
            rgb = resolved.Rgb;
            alpha = resolved.Alpha;
        }
        uint? opacityValue = null;
        if (value.TryGetProperty("opacity", out var opacity))
            opacityValue = Opacity(ChartEffectNumber(resolveOpacity(opacity), 0, 1));
        else if (alpha != 1 || color.ValueKind == JsonValueKind.Object && color.TryGetProperty("alpha", out _))
            opacityValue = Opacity(ChartEffectNumber(alpha, 0, 1));
        return (rgb, scheme, opacityValue);
    }

    private static double ChartEffectNumber(double number, double min, double max)
    {
        if (!double.IsFinite(number) || number < min || number > max)
            throw new CodecException("invalid_presentation_effects", "Chart text effect has an out-of-range number.");
        return number;
    }

    private static string BuildChartHighlight(JsonElement value, Catalog catalog)
    {
        var resolved = catalog.Color(value);
        return XlsxChartTextStyleCodec.HighlightRgb(resolved.Rgb, resolved.Alpha);
    }

    private static void ApplyChartTextFill(
        SpreadsheetChartTextStyleArtifact output,
        SpreadsheetChartSurfaceFill fill)
    {
        output.Fill = null;
        output.ColorRgb = string.Empty;
        output.ClearOpacityThousandthPercent();
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb)
        {
            output.ColorRgb = fill.SolidRgb;
            if (fill.HasOpacityThousandthPercent)
                output.OpacityThousandthPercent = fill.OpacityThousandthPercent;
        }
        else
            output.Fill = fill;
    }

    internal static void ApplyCustomGeometry(PresentationShape target, JsonElement geometry, string elementId, bool allowShapeGraph = false)
    {
        if (!geometry.TryGetProperty("viewBox", out var viewBox) ||
            !geometry.TryGetProperty("paths", out var paths))
            throw Unsupported(elementId, "custom geometry has no compiler-owned path graph");
        if (geometry.TryGetProperty("adjustmentHandles", out var handles))
        {
            if (!allowShapeGraph) throw Unsupported(elementId, "adjustment handles belong to custom shapes, not masks or clips");
            foreach (var handle in handles.EnumerateArray())
                target.CustomAdjustmentHandles.Add(PpjCustomGeometryHandleCodec.Read(handle));
        }
        if (geometry.TryGetProperty("connectionSites", out var sites))
        {
            if (!allowShapeGraph) throw Unsupported(elementId, "connection sites belong to custom shapes, not masks or clips");
            foreach (var source in sites.EnumerateArray())
            {
                var site = new PresentationCustomGeometryConnectionSite();
                var angle = source.GetProperty("angle");
                var x = source.GetProperty("x");
                var y = source.GetProperty("y");
                if (angle.ValueKind == JsonValueKind.String) site.AngleReference = angle.GetString()!;
                else site.Angle60000 = Angle(angle.GetDouble());
                if (x.ValueKind == JsonValueKind.String) site.XReference = x.GetString()!;
                else site.XEmu = Emu(x.GetDouble());
                if (y.ValueKind == JsonValueKind.String) site.YReference = y.GetString()!;
                else site.YEmu = Emu(y.GetDouble());
                target.CustomConnectionSites.Add(site);
            }
        }
        if (geometry.TryGetProperty("adjustments", out var adjustments))
        {
            if (!allowShapeGraph) throw Unsupported(elementId, "custom adjustment formulas belong to shapes, not masks or clips");
            foreach (var adjustment in adjustments.EnumerateArray())
                target.CustomAdjustments.Add(new PresentationCustomGeometryGuide
                {
                    Name = adjustment.GetProperty("name").GetString()!,
                    Formula = adjustment.GetProperty("formula").GetString()!,
                });
        }
        if (geometry.TryGetProperty("guides", out var guides))
        {
            if (!allowShapeGraph) throw Unsupported(elementId, "geometry guides belong to custom shapes, not masks or clips");
            foreach (var guide in guides.EnumerateArray())
                target.CustomGuides.Add(new PresentationCustomGeometryGuide
                {
                    Name = guide.GetProperty("name").GetString()!,
                    Formula = guide.GetProperty("formula").GetString()!,
                });
        }
        target.TextRectangle = null;
        if (geometry.TryGetProperty("textRectangle", out var rectangle))
        {
            if (!allowShapeGraph) throw Unsupported(elementId, "textRectangle belongs to custom shapes, not masks or clips");
            var result = new PresentationCustomGeometryTextRectangle();
            foreach (var edge in rectangle.EnumerateObject())
            {
                var reference = edge.Value.ValueKind == JsonValueKind.String ? edge.Value.GetString() : null;
                var value = reference is null ? Emu(edge.Value.GetDouble()) : 0;
                switch (edge.Name)
                {
                    case "left": if (reference is null) result.LeftEmu = value; else result.LeftReference = reference; break;
                    case "top": if (reference is null) result.TopEmu = value; else result.TopReference = reference; break;
                    case "right": if (reference is null) result.RightEmu = value; else result.RightReference = reference; break;
                    case "bottom": if (reference is null) result.BottomEmu = value; else result.BottomReference = reference; break;
                }
            }
            target.TextRectangle = result;
        }
        var originX = viewBox.GetProperty("x").GetDouble();
        var originY = viewBox.GetProperty("y").GetDouble();
        var width = CustomPathCoordinate(viewBox.GetProperty("width").GetDouble());
        var height = CustomPathCoordinate(viewBox.GetProperty("height").GetDouble());
        foreach (var sourcePath in paths.EnumerateArray())
        {
            var path = new PresentationCustomGeometryPath { Width = width, Height = height };
            if (sourcePath.TryGetProperty("viewport", out var viewport))
            {
                if (!allowShapeGraph) throw Unsupported(elementId, "path viewport overrides belong to custom shapes, not masks or clips");
                path.Width = CustomPathCoordinate(viewport.GetProperty("width").GetDouble());
                path.Height = CustomPathCoordinate(viewport.GetProperty("height").GetDouble());
            }
            if (sourcePath.TryGetProperty("fill", out var fill))
                path.FillMode = fill.GetBoolean()
                    ? PresentationCustomGeometryPath.Types.FillMode.Normal
                    : PresentationCustomGeometryPath.Types.FillMode.None;
            if (sourcePath.TryGetProperty("stroke", out var stroke)) path.Stroke = stroke.GetBoolean();
            if (sourcePath.TryGetProperty("extrusionOk", out var extrusion)) path.ExtrusionAllowed = extrusion.GetBoolean();
            var hasCurrentPoint = false;
            var hasSubpathStart = false;
            foreach (var sourceCommand in sourcePath.GetProperty("commands").EnumerateArray())
            {
                if (!allowShapeGraph && sourceCommand.EnumerateObject().Any(property =>
                    property.Name != "op" && property.Value.ValueKind == JsonValueKind.String))
                    throw Unsupported(elementId, "referenced paths belong to custom shapes, not masks or clips");
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
                        command.ArcTo = CustomArc(sourceCommand);
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
        string yName)
    {
        var point = new PresentationCustomGeometryPoint();
        var x = command.GetProperty(xName); var y = command.GetProperty(yName);
        if (x.ValueKind == JsonValueKind.String) point.XReference = x.GetString()!;
        else point.X = CustomPathCoordinate(x.GetDouble() - originX);
        if (y.ValueKind == JsonValueKind.String) point.YReference = y.GetString()!;
        else point.Y = CustomPathCoordinate(y.GetDouble() - originY);
        return point;
    }

    private static PresentationCustomGeometryArc CustomArc(JsonElement command)
    {
        var arc = new PresentationCustomGeometryArc();
        var rx = command.GetProperty("radiusX"); var ry = command.GetProperty("radiusY");
        var start = command.GetProperty("startAngle"); var sweep = command.GetProperty("sweepAngle");
        if (rx.ValueKind == JsonValueKind.String) arc.WidthRadiusReference = rx.GetString()!;
        else arc.WidthRadius = CustomPathCoordinate(rx.GetDouble());
        if (ry.ValueKind == JsonValueKind.String) arc.HeightRadiusReference = ry.GetString()!;
        else arc.HeightRadius = CustomPathCoordinate(ry.GetDouble());
        if (start.ValueKind == JsonValueKind.String) arc.StartAngleReference = start.GetString()!;
        else arc.StartAngle = Angle(start.GetDouble());
        if (sweep.ValueKind == JsonValueKind.String) arc.SweepAngleReference = sweep.GetString()!;
        else arc.SweepAngle = Angle(sweep.GetDouble());
        return arc;
    }

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
        "custom" => SpreadsheetChartErrorBarValueType.Custom,
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

    private static string Arrow(string? value) => value is null or "none" ? string.Empty : value == "open" ? "arrow" : value;

    private static int CropValue(JsonElement crop, string property) =>
        crop.TryGetProperty(property, out var value) ? checked((int)Math.Round(value.GetDouble() * 100_000)) : 0;

    private static void RejectProperties(JsonElement value, string elementId, params string[] names)
    {
        if (value.ValueKind != JsonValueKind.Object) return;
        foreach (var name in names)
            if (value.TryGetProperty(name, out _))
                throw Unsupported(elementId, $"{name} is valid PPJ but not yet compiler-owned for this element");
    }

    internal static string NativeUnderline(string value) => value switch
    {
        "single" => "sng",
        "double" => "dbl",
        _ => value,
    };

    internal static string NativeChartAlignment(string value) => value switch
    {
        "left" => "l",
        "center" => "ctr",
        "right" => "r",
        "justify" => "just",
        _ => value,
    };

    internal static string NativeStrike(JsonElement value) => value.ValueKind == JsonValueKind.True
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
        (run.Field?.Text ?? (run.Formula is null ? (run.LineBreak ? "\n" : string.Empty) : PpjLatexCompiler.Compile(run.Formula.Source).PlainText));

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

    private static double? OptionalOpacity(JsonElement value, string name, Catalog catalog, string owner)
    {
        if (!value.TryGetProperty(name, out var property)) return null;
        var opacity = catalog.NumberToken(property, "opacity", owner);
        if (opacity is < 0 or > 1)
            throw new CodecException("ppj.opacity", $"PPJ {owner} must be between 0 and 1.");
        return opacity;
    }

    private static double StrokeWidth(JsonElement stroke, Catalog catalog, string owner)
    {
        var width = catalog.NumberToken(stroke.GetProperty("width"), "size", owner);
        if (width is < 0 or > 1000)
            throw new CodecException("ppj.stroke.width", $"PPJ {owner} must be between 0 and 1000 points.");
        return width;
    }

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

    private static JsonElement? MergeJsonObjects(params JsonElement?[] sources) => MergeJsonObjects(false, sources);

    private static JsonElement? MergeJsonObjects(bool mergeParagraphSpacing, params JsonElement?[] sources)
    {
        Dictionary<string, JsonElement>? merged = null;
        foreach (var source in sources)
        {
            if (source is not { ValueKind: JsonValueKind.Object } value) continue;
            merged ??= new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            if (mergeParagraphSpacing)
            {
                if (value.TryGetProperty("tabStops", out _) || value.TryGetProperty("noTabStops", out _))
                {
                    merged.Remove("tabStops");
                    merged.Remove("noTabStops");
                }
                foreach (var slot in new[] { "spaceBefore", "spaceAfter", "lineSpacing" })
                    if (value.TryGetProperty(slot, out _) || value.TryGetProperty(slot + "Multiplier", out _))
                    {
                        merged.Remove(slot);
                        merged.Remove(slot + "Multiplier");
                    }
            }
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

    internal sealed class Catalog
    {
        private readonly Dictionary<string, (string Rgb, double Alpha)> _colors;
        private readonly Dictionary<string, string> _fonts;
        private readonly Dictionary<string, JsonElement> _textStyles;
        private readonly Dictionary<string, JsonElement> _shapeStyles;
        private readonly Dictionary<string, JsonElement> _imageStyles;
        private readonly Dictionary<string, JsonElement> _chartStyles;
        private readonly Dictionary<string, JsonElement> _tableStyles;
        private readonly Dictionary<string, (double Width, double Height)> _assetDimensions;
        private readonly Dictionary<string, string> _nativeAssetIds;
        private readonly Dictionary<string, PpjSmartArtDefinition> _smartArtDefinitions;
        private readonly Dictionary<string, (string Rgb, double Alpha)> _grammarColors;
        private readonly Dictionary<string, (string Kind, JsonElement Value)> _grammarTokens;
        private readonly Dictionary<string, string[]> _stylePrecedence;
        private readonly Dictionary<string, JsonElement> _masterTextStyles;
        private readonly Dictionary<string, JsonElement> _layoutTextStyles;
        private readonly Dictionary<string, string> _layoutMasters;
        private readonly JsonElement _theme;

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
            _imageStyles = Styles(styles, "image");
            _chartStyles = Styles(styles, "chart");
            _tableStyles = Styles(styles, "table");
            _theme = design.GetProperty("theme").Clone();
            _stylePrecedence = ParseStylePrecedence(design);
            _masterTextStyles = DirectTextStyles(design, "masters");
            _layoutTextStyles = DirectTextStyles(design, "layouts");
            _layoutMasters = DirectLayoutMasters(design);
            _grammarTokens = new Dictionary<string, (string Kind, JsonElement Value)>(StringComparer.Ordinal);
            _grammarColors = new Dictionary<string, (string Rgb, double Alpha)>(StringComparer.Ordinal);
            if (design.TryGetProperty("grammar", out var grammar) &&
                grammar.ValueKind == JsonValueKind.Object &&
                grammar.TryGetProperty("tokens", out var grammarTokens) &&
                grammarTokens.ValueKind == JsonValueKind.Object)
            {
                foreach (var token in grammarTokens.EnumerateObject())
                {
                    if (token.Value.ValueKind != JsonValueKind.Object ||
                        !token.Value.TryGetProperty("kind", out var kind) ||
                        !token.Value.TryGetProperty("value", out var value)) continue;
                    var kindName = kind.GetString();
                    if (kindName is null) continue;
                    _grammarTokens[token.Name] = (kindName, value.Clone());
                    if (!string.Equals(kindName, "color", StringComparison.Ordinal) || value.ValueKind != JsonValueKind.String) continue;
                    var raw = value.GetString();
                    if (raw is null) continue;
                    try { _grammarColors[token.Name] = ParseHexColor(raw); }
                    catch (FormatException) { }
                }
            }
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
            if (theme.TryGetProperty("accentColors", out var accentColors))
            {
                output.AccentRgb.Add(new[] { "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" }
                    .Select(role => NormalizeThemeColor(accentColors.GetProperty(role).GetString()!)));
            }
            else
            {
                output.AccentRgb.Add(theme.GetProperty("colors").EnumerateArray()
                    .Take(6)
                    .Select(color => NormalizeThemeColor(color.GetProperty("value").GetString()!)));
            }
            if (theme.TryGetProperty("accentTransforms", out var accentTransforms))
            {
                foreach (var role in new[] { "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" })
                {
                    if (!accentTransforms.TryGetProperty(role, out var transform)) continue;
                    var authored = new PresentationThemeColorTransform { Role = role };
                    if (transform.TryGetProperty("tint", out var tint))
                        authored.TintThousandth = ThemeTransformThousandth(tint, role, "tint");
                    if (transform.TryGetProperty("shade", out var shade))
                        authored.ShadeThousandth = ThemeTransformThousandth(shade, role, "shade");
                    if (transform.TryGetProperty("lumMod", out var luminanceModulation))
                        authored.LuminanceModulationThousandth = ThemeTransformThousandth(luminanceModulation, role, "lumMod");
                    if (transform.TryGetProperty("lumOff", out var luminanceOffset))
                        authored.LuminanceOffsetThousandth = ThemeTransformSignedThousandth(luminanceOffset, role, "lumOff");
                    if (transform.TryGetProperty("alphaMod", out var alphaModulation))
                        authored.AlphaModulationThousandth = ThemeTransformThousandth(alphaModulation, role, "alphaMod");
                    if (transform.TryGetProperty("alphaOff", out var alphaOffset))
                        authored.AlphaOffsetThousandth = ThemeTransformSignedThousandth(alphaOffset, role, "alphaOff");
                    if (transform.TryGetProperty("satMod", out var saturationModulation))
                        authored.SaturationModulationThousandth = ThemeTransformThousandth(saturationModulation, role, "satMod");
                    if (transform.TryGetProperty("satOff", out var saturationOffset))
                        authored.SaturationOffsetThousandth = ThemeTransformSignedThousandth(saturationOffset, role, "satOff");
                    if (transform.TryGetProperty("redMod", out var redModulation))
                        authored.RedModulationThousandth = ThemeTransformThousandth(redModulation, role, "redMod");
                    if (transform.TryGetProperty("redOff", out var redOffset))
                        authored.RedOffsetThousandth = ThemeTransformSignedThousandth(redOffset, role, "redOff");
                    if (transform.TryGetProperty("greenMod", out var greenModulation))
                        authored.GreenModulationThousandth = ThemeTransformThousandth(greenModulation, role, "greenMod");
                    if (transform.TryGetProperty("greenOff", out var greenOffset))
                        authored.GreenOffsetThousandth = ThemeTransformSignedThousandth(greenOffset, role, "greenOff");
                    if (transform.TryGetProperty("blueMod", out var blueModulation))
                        authored.BlueModulationThousandth = ThemeTransformThousandth(blueModulation, role, "blueMod");
                    if (transform.TryGetProperty("blueOff", out var blueOffset))
                        authored.BlueOffsetThousandth = ThemeTransformSignedThousandth(blueOffset, role, "blueOff");
                    if (transform.TryGetProperty("hueMod", out var hueModulation))
                        authored.HueModulationThousandth = ThemeTransformThousandth(hueModulation, role, "hueMod");
                    if (transform.TryGetProperty("hueOff", out var hueOffset))
                        authored.HueOffsetAngleThousandth = ThemeTransformHueOffset(hueOffset, role, "hueOff");
                    if (transform.TryGetProperty("gray", out var gray))
                        authored.Gray = ThemeTransformFlag(gray, role, "gray");
                    if (transform.TryGetProperty("comp", out var complement))
                        authored.Comp = ThemeTransformFlag(complement, role, "comp");
                    if (transform.TryGetProperty("inv", out var inverse))
                        authored.Inv = ThemeTransformFlag(inverse, role, "inv");
                    if (transform.TryGetProperty("gamma", out var gamma))
                        authored.Gamma = ThemeTransformFlag(gamma, role, "gamma");
                    if (transform.TryGetProperty("invGamma", out var inverseGamma))
                        authored.InvGamma = ThemeTransformFlag(inverseGamma, role, "invGamma");
                    output.AccentTransforms.Add(authored);
                }
            }
            if (theme.TryGetProperty("fontScheme", out var fontScheme))
            {
                output.MajorFontFamily = fontScheme.GetProperty("major").GetString()!;
                output.MinorFontFamily = fontScheme.GetProperty("minor").GetString()!;
                if (fontScheme.TryGetProperty("majorEastAsia", out var majorEastAsia))
                    output.MajorFontFamilyEastAsia = majorEastAsia.GetString()!;
                if (fontScheme.TryGetProperty("majorComplexScript", out var majorComplexScript))
                    output.MajorFontFamilyComplexScript = majorComplexScript.GetString()!;
                if (fontScheme.TryGetProperty("minorEastAsia", out var minorEastAsia))
                    output.MinorFontFamilyEastAsia = minorEastAsia.GetString()!;
                if (fontScheme.TryGetProperty("minorComplexScript", out var minorComplexScript))
                    output.MinorFontFamilyComplexScript = minorComplexScript.GetString()!;
            }
            else
            {
                var fonts = design.GetProperty("fonts").EnumerateArray().ToArray();
                if (fonts.Length > 0)
                {
                    output.MajorFontFamily = fonts[0].GetProperty("family").GetString()!;
                    output.MinorFontFamily = fonts.Length > 1
                        ? fonts[1].GetProperty("family").GetString()!
                        : output.MajorFontFamily;
                }
            }
            if (theme.TryGetProperty("colorRoles", out var colorRoles))
            {
                output.Dark1Rgb = NormalizeThemeColor(colorRoles.GetProperty("dark1").GetString()!);
                output.Light1Rgb = NormalizeThemeColor(colorRoles.GetProperty("light1").GetString()!);
                output.Dark2Rgb = NormalizeThemeColor(colorRoles.GetProperty("dark2").GetString()!);
                output.Light2Rgb = NormalizeThemeColor(colorRoles.GetProperty("light2").GetString()!);
                output.HyperlinkRgb = NormalizeThemeColor(colorRoles.GetProperty("hyperlink").GetString()!);
                output.FollowedHyperlinkRgb = NormalizeThemeColor(colorRoles.GetProperty("followedHyperlink").GetString()!);
            }
            return output;
        }

        private static uint ThemeTransformThousandth(JsonElement value, string role, string operation)
        {
            var fraction = value.GetDouble();
            if (!double.IsFinite(fraction) || fraction < 0 || fraction > 1)
                throw new CodecException(
                    "ppj.theme.transform",
                    $"PPJ theme accent {role} {operation} must be between 0 and 1.",
                    $"$.design.theme.accentTransforms.{role}.{operation}");
            return checked((uint)Math.Round(fraction * 100_000d, MidpointRounding.AwayFromZero));
        }

        private static int ThemeTransformSignedThousandth(JsonElement value, string role, string operation)
        {
            var fraction = value.GetDouble();
            if (!double.IsFinite(fraction) || fraction < -1 || fraction > 1)
                throw new CodecException(
                    "ppj.theme.transform",
                    $"PPJ theme accent {role} {operation} must be between -1 and 1.",
                    $"$.design.theme.accentTransforms.{role}.{operation}");
            return checked((int)Math.Round(fraction * 100_000d, MidpointRounding.AwayFromZero));
        }

        private static bool ThemeTransformFlag(JsonElement value, string role, string operation)
        {
            if (value.ValueKind != JsonValueKind.True)
                throw new CodecException(
                    "ppj.theme.transform",
                    $"PPJ theme accent {role} {operation} must be true when declared.",
                    $"$.design.theme.accentTransforms.{role}.{operation}");
            return true;
        }

        private static int ThemeTransformHueOffset(JsonElement value, string role, string operation)
        {
            var degrees = value.GetDouble();
            if (!double.IsFinite(degrees) || degrees < -360 || degrees > 360)
                throw new CodecException(
                    "ppj.theme.transform",
                    $"PPJ theme accent {role} {operation} must be between -360 and 360 degrees.",
                    $"$.design.theme.accentTransforms.{role}.{operation}");
            return checked((int)Math.Round(degrees * 60_000d, MidpointRounding.AwayFromZero));
        }

        internal bool HasGrammarToken(string name) => _grammarTokens.ContainsKey(name);

        internal (string Rgb, double Alpha) Color(JsonElement color) => ResolveColor(color, preferGrammar: false);

        internal (string Rgb, double Alpha) BulletColor(JsonElement color) => ResolveColor(color, preferGrammar: true);

        private (string Rgb, double Alpha) ResolveColor(JsonElement color, bool preferGrammar)
        {
            if (color.ValueKind == JsonValueKind.String) return ParseHexColor(color.GetString()!);
            var token = color.GetProperty("token").GetString()!;
            var value = (Rgb: string.Empty, Alpha: 1d);
            if (preferGrammar && _grammarTokens.ContainsKey(token) || !_colors.TryGetValue(token, out value))
            {
                if (!_grammarTokens.TryGetValue(token, out var declared))
                    throw new CodecException("ppj.color.unknown", $"PPJ color token {token} is not declared.");
                if (!string.Equals(declared.Kind, "color", StringComparison.Ordinal))
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ color token {token} declares kind {declared.Kind}, expected color.");
                if (!_grammarColors.TryGetValue(token, out value))
                    throw new CodecException("ppj.grammar.tokenValue", $"PPJ color token {token} must resolve to a valid color.");
            }
            var rgb = value.Rgb;
            if (color.TryGetProperty("tint", out var tint)) rgb = MixRgb(rgb, "FFFFFF", tint.GetDouble());
            if (color.TryGetProperty("shade", out var shade)) rgb = MixRgb(rgb, "000000", shade.GetDouble());
            return (rgb, color.TryGetProperty("alpha", out var alpha) ? alpha.GetDouble() : value.Alpha);
        }

        internal string Font(string id) => _fonts.TryGetValue(id, out var value)
            ? value
            : throw new CodecException("ppj.font.unknown", $"PPJ font token {id} is not declared.");

        internal string Font(JsonElement value) => value.ValueKind == JsonValueKind.String
            ? Font(value.GetString()!)
            : StringToken(value, "font", "font");

        internal double NumberToken(JsonElement value, string expectedKind, string owner)
        {
            var resolved = ResolveToken(value, expectedKind, owner);
            if (resolved.ValueKind != JsonValueKind.Number || !resolved.TryGetDouble(out var number) || !double.IsFinite(number))
                throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token for {owner} must resolve to a finite number.");
            return number;
        }

        internal double PositiveNumberToken(JsonElement value, string expectedKind, string owner)
        {
            var number = NumberToken(value, expectedKind, owner);
            if (number <= 0)
                throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token for {owner} must resolve to a positive number.");
            return number;
        }

        internal bool BooleanToken(JsonElement value, string expectedKind, string owner)
        {
            var resolved = ResolveToken(value, expectedKind, owner);
            if (resolved.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token for {owner} must resolve to a boolean.");
            return resolved.GetBoolean();
        }

        internal string StringToken(JsonElement value, string expectedKind, string owner, bool allowEmpty = false)
        {
            var resolved = ResolveToken(value, expectedKind, owner);
            if (resolved.ValueKind != JsonValueKind.String || !allowEmpty && string.IsNullOrEmpty(resolved.GetString()))
                throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token for {owner} must resolve to {(allowEmpty ? "a string" : "a non-empty string")}.");
            return resolved.GetString()!;
        }

        internal string LanguageTagToken(JsonElement value, string owner) =>
            PptxLanguageTag.Validate(StringToken(value, "string", owner));

        private JsonElement ResolveToken(JsonElement value, string expectedKind, string owner)
        {
            if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty("token", out var tokenValue)) return value;
            var token = tokenValue.GetString();
            if (token is null || !_grammarTokens.TryGetValue(token, out var definition))
                throw new CodecException("ppj.grammar.tokenUnknown", $"PPJ grammar token {token ?? "(missing)"} for {owner} is not declared.");
            if (!string.Equals(definition.Kind, expectedKind, StringComparison.Ordinal))
                throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {token} has kind {definition.Kind}, but {owner} requires {expectedKind}.");
            return definition.Value;
        }

        internal JsonElement? TextStyle(string? id) => Style(_textStyles, id);
        internal JsonElement? ShapeStyle(string? id) => Style(_shapeStyles, id);
        internal JsonElement? ImageStyle(string? id) => Style(_imageStyles, id);
        internal JsonElement? ChartStyle(string? id) => Style(_chartStyles, id);
        internal JsonElement? TableStyle(string? id) => Style(_tableStyles, id);
        internal bool HasStylePrecedence(string target) => _stylePrecedence.ContainsKey(target);
        internal TextPrecedenceContext TextPrecedenceForLayout(string? layoutId)
        {
            if (layoutId is null)
                return new(null, null, null, null);
            var layoutStyle = _layoutTextStyles.TryGetValue(layoutId, out var layoutSource)
                ? (JsonElement?)layoutSource
                : null;
            var masterStyle = _layoutMasters.TryGetValue(layoutId, out var masterId) &&
                               _masterTextStyles.TryGetValue(masterId, out var source)
                ? source
                : (JsonElement?)null;
            return new(layoutStyle, masterStyle, null, null);
        }

        internal JsonElement? TextPropertyByPrecedence(
            string target,
            JsonElement? run,
            JsonElement? paragraph,
            JsonElement? element,
            JsonElement? styleRef,
            JsonElement? layout,
            JsonElement? master,
            JsonElement? legacyInline,
            JsonElement? legacyMaster,
            JsonElement? defaultValue = null)
        {
            var field = target;
            var separator = field.IndexOf('.');
            if (separator >= 0) field = field[(separator + 1)..];

            if (!_stylePrecedence.TryGetValue(target, out var sources))
                return PathProperty(run, field) ??
                    PathProperty(legacyInline, field) ??
                    PathProperty(legacyMaster, field) ??
                    PathProperty(styleRef, field) ??
                    defaultValue;

            foreach (var source in sources)
            {
                var value = source switch
                {
                    "run" => PathProperty(run, field),
                    "paragraph" => PathProperty(paragraph, field),
                    "element" => PathProperty(element, field),
                    "inline" => PathProperty(run, field) ??
                        PathProperty(element, field) ??
                        PathProperty(legacyInline, field),
                    "styleRef" => PathProperty(styleRef, field),
                    "layout" => PathProperty(layout, field),
                    "master" => PathProperty(master, field) ?? PathProperty(legacyMaster, field),
                    "theme" => ThemeTextProperty(field),
                    "default" => defaultValue ?? GrammarDefault(target, field),
                    _ => null,
                };
                if (value is not null) return value;
            }
            return null;
        }
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

        /// <summary>
        /// Resolve a bounded authored property using the declared grammar
        /// source order. If no rule targets the property, preserve the
        /// historical inline-style then named-style behavior. For a declared
        /// rule, the inline source checks the element before its style object,
        /// matching the read-only grammar evaluator.
        /// </summary>
        internal JsonElement? PropertyByPrecedence(
            string target,
            JsonElement? element,
            JsonElement? inlineStyle,
            JsonElement? styleRef,
            JsonElement? master = null,
            JsonElement? defaultValue = null,
            bool includeElementWhenUndeclared = false)
        {
            var field = target;
            var separator = field.IndexOf('.');
            if (separator >= 0) field = field[(separator + 1)..];

            if (!_stylePrecedence.TryGetValue(target, out var sources))
                return (includeElementWhenUndeclared ? PathProperty(element, field) : null) ??
                    PathProperty(inlineStyle, field) ??
                    PathProperty(master, field) ??
                    PathProperty(styleRef, field) ??
                    defaultValue;

            foreach (var source in sources)
            {
                var value = source switch
                {
                    "inline" => PathProperty(element, field) ?? PathProperty(inlineStyle, field),
                    "styleRef" => PathProperty(styleRef, field),
                    "theme" => ThemeProperty(field),
                    "master" => PathProperty(master, field),
                    "default" => defaultValue ?? GrammarDefault(target, field),
                    _ => null,
                };
                if (value is not null) return value;
            }
            return null;
        }

        private JsonElement? ThemeProperty(string field)
        {
            var value = PathProperty(_theme, field);
            if (value is not null) return value;
            var leaf = field.Split('.').LastOrDefault();
            if (leaf is null || !_theme.TryGetProperty("colors", out var colors) || colors.ValueKind != JsonValueKind.Array)
                return null;
            foreach (var color in colors.EnumerateArray())
            {
                if (color.TryGetProperty("id", out var id) && id.GetString() == leaf && color.TryGetProperty("value", out var colorValue))
                    return colorValue.Clone();
            }
            return null;
        }

        private JsonElement? ThemeTextProperty(string field)
        {
            return _theme.TryGetProperty("textStyle", out var textStyle)
                ? PathProperty(textStyle, field)
                : null;
        }

        private JsonElement? GrammarDefault(string target, string field)
        {
            if (_grammarTokens.TryGetValue(target, out var exact)) return exact.Value.Clone();
            var leaf = field.Split('.').LastOrDefault();
            return leaf is not null && _grammarTokens.TryGetValue(leaf, out var fallback)
                ? fallback.Value.Clone()
                : null;
        }

        private static JsonElement? PathProperty(JsonElement? value, string path)
        {
            var cursor = value;
            foreach (var part in path.Split('.', StringSplitOptions.RemoveEmptyEntries))
            {
                if (cursor is not { ValueKind: JsonValueKind.Object } objectValue ||
                    !objectValue.TryGetProperty(part, out var next)) return null;
                cursor = next;
            }
            return cursor?.Clone();
        }

        private static Dictionary<string, string[]> ParseStylePrecedence(JsonElement design)
        {
            var output = new Dictionary<string, string[]>(StringComparer.Ordinal);
            if (!design.TryGetProperty("grammar", out var grammar) || grammar.ValueKind != JsonValueKind.Object ||
                !grammar.TryGetProperty("stylePrecedence", out var precedence) || precedence.ValueKind != JsonValueKind.Array)
                return output;
            foreach (var rule in precedence.EnumerateArray())
            {
                if (!rule.TryGetProperty("target", out var target) || target.GetString() is not { } targetName ||
                    !rule.TryGetProperty("sources", out var sources) || sources.ValueKind != JsonValueKind.Array) continue;
                output[targetName] = sources.EnumerateArray()
                    .Select(source => source.GetString())
                    .Where(source => source is not null)
                    .Select(source => source!)
                    .ToArray();
            }
            return output;
        }

        private static Dictionary<string, JsonElement> DirectTextStyles(JsonElement design, string collection)
        {
            if (!design.TryGetProperty(collection, out var definitions) || definitions.ValueKind != JsonValueKind.Array)
                return new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            return definitions.EnumerateArray()
                .Where(item => item.TryGetProperty("id", out _) && item.TryGetProperty("style", out _))
                .ToDictionary(
                    item => item.GetProperty("id").GetString()!,
                    item => item.GetProperty("style").Clone(),
                    StringComparer.Ordinal);
        }

        private static Dictionary<string, string> DirectLayoutMasters(JsonElement design)
        {
            if (!design.TryGetProperty("layouts", out var definitions) || definitions.ValueKind != JsonValueKind.Array)
                return new Dictionary<string, string>(StringComparer.Ordinal);
            return definitions.EnumerateArray()
                .Where(item => item.TryGetProperty("id", out _) && item.TryGetProperty("master", out _))
                .ToDictionary(
                    item => item.GetProperty("id").GetString()!,
                    item => item.GetProperty("master").GetString()!,
                    StringComparer.Ordinal);
        }

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
        var opacity = fill.TryGetProperty("opacity", out var opacityValue)
            ? catalog.NumberToken(opacityValue, "opacity", "fill opacity")
            : color.Alpha;
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
            output.Angle60000 = Angle(normalized) % (360 * 60_000);
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
            if (item.TryGetProperty("opacity", out _) || alpha < 1) stop.OpacityThousandthPercent = Opacity(alpha);
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
            OpacityThousandthPercent = Opacity(value.TryGetProperty("opacity", out var opacity)
                ? catalog.NumberToken(opacity, "opacity", "shadow opacity")
                : color.Alpha),
        };
        if (schemeToken is not null)
            output.ColorScheme = schemeToken;
        if (value.TryGetProperty("alignment", out var alignment)) output.Alignment = alignment.GetString()!;
        if (value.TryGetProperty("rotateWithShape", out var rotateWithShape)) output.RotateWithShape = rotateWithShape.GetBoolean();
        return output;
    }

    internal static PresentationGlow BuildParagraphGlow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var color = value.GetProperty("color");
        var transformed = color.ValueKind == JsonValueKind.Object &&
            (color.TryGetProperty("tint", out _) || color.TryGetProperty("shade", out _));
        var glow = BuildChartTextGlow(value, resolveColor, resolveOpacity, name => transformed || declaredToken(name));
        if (!glow.HasOpacityThousandthPercent && color.ValueKind == JsonValueKind.String &&
            color.GetString()!.TrimStart('#').Length == 8)
            glow.OpacityThousandthPercent = 100_000;
        return glow;
    }

    internal static PresentationInnerShadow BuildParagraphInnerShadow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var color = value.GetProperty("color");
        var transformed = color.ValueKind == JsonValueKind.Object &&
            (color.TryGetProperty("tint", out _) || color.TryGetProperty("shade", out _));
        var shadow = BuildChartTextInnerShadow(value, resolveColor, resolveOpacity, name => transformed || declaredToken(name));
        if (!shadow.HasOpacityThousandthPercent && color.ValueKind == JsonValueKind.String &&
            color.GetString()!.TrimStart('#').Length == 8)
            shadow.OpacityThousandthPercent = 100_000;
        return shadow;
    }

    internal static PresentationShadow BuildParagraphShadow(
        JsonElement value,
        Func<JsonElement, (string Rgb, double Alpha)> resolveColor,
        Func<JsonElement, double> resolveOpacity,
        Func<string, bool> declaredToken)
    {
        var color = value.GetProperty("color");
        var transformed = color.ValueKind == JsonValueKind.Object &&
            (color.TryGetProperty("tint", out _) || color.TryGetProperty("shade", out _));
        var shadow = BuildChartTextShadow(value, resolveColor, resolveOpacity, name => transformed || declaredToken(name));
        if (!shadow.HasOpacityThousandthPercent && color.ValueKind == JsonValueKind.String &&
            color.GetString()!.TrimStart('#').Length == 8)
            shadow.OpacityThousandthPercent = 100_000;
        return shadow;
    }

    private static PresentationGlow BuildGlow(JsonElement value, Catalog catalog)
    {
        var colorValue = value.GetProperty("color");
        var schemeToken = colorValue.ValueKind == JsonValueKind.Object && colorValue.TryGetProperty("token", out var token) &&
                          PptxColor.TrySchemeToken(token.GetString() ?? string.Empty, out var recognizedScheme)
            ? recognizedScheme
            : null;
        var color = schemeToken is null
            ? catalog.Color(colorValue)
            : (Rgb: string.Empty, Alpha: colorValue.TryGetProperty("alpha", out var alpha) ? alpha.GetDouble() : 1d);
        var output = new PresentationGlow
        {
            ColorRgb = color.Rgb,
            RadiusEmu = Emu(value.GetProperty("radius").GetDouble()),
            OpacityThousandthPercent = Opacity(value.TryGetProperty("opacity", out var opacity)
                ? catalog.NumberToken(opacity, "opacity", "glow opacity")
                : color.Alpha),
        };
        if (schemeToken is not null)
            output.ColorScheme = schemeToken;
        return output;
    }

    private static PresentationSoftEdge BuildSoftEdge(JsonElement value) => new()
    {
        RadiusEmu = Emu(value.GetProperty("radius").GetDouble()),
    };

    private static PresentationInnerShadow BuildInnerShadow(JsonElement value, Catalog catalog)
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
        var output = new PresentationInnerShadow
        {
            ColorRgb = color.Rgb,
            BlurRadiusEmu = Emu(value.GetProperty("blur").GetDouble()),
            DistanceEmu = Emu(value.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = Angle(normalized),
            OpacityThousandthPercent = Opacity(value.TryGetProperty("opacity", out var opacity)
                ? catalog.NumberToken(opacity, "opacity", "inner shadow opacity")
                : color.Alpha),
        };
        if (schemeToken is not null)
            output.ColorScheme = schemeToken;
        return output;
    }

    private static PresentationReflection BuildReflection(JsonElement value, Catalog catalog)
    {
        var startOpacity = catalog.NumberToken(value.GetProperty("startOpacity"), "opacity", "reflection start opacity");
        var endOpacity = catalog.NumberToken(value.GetProperty("endOpacity"), "opacity", "reflection end opacity");
        if (startOpacity is < 0 or > 1 || endOpacity is < 0 or > 1)
            throw new CodecException("ppj.opacity", "PPJ reflection opacity must be between 0 and 1.");
        var degrees = value.GetProperty("angle").GetDouble();
        var normalized = ((degrees % 360) + 360) % 360;
        var output = new PresentationReflection
        {
            StartPositionThousandthPercent = 0,
            EndPositionThousandthPercent = 100_000,
            BlurRadiusEmu = Emu(value.GetProperty("blur").GetDouble()),
            StartOpacityThousandthPercent = Opacity(startOpacity),
            EndOpacityThousandthPercent = Opacity(endOpacity),
            DistanceEmu = Emu(value.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = Angle(normalized),
        };
        if (value.TryGetProperty("fadeAngle", out var fadeAngle))
        {
            var fadeDegrees = fadeAngle.GetDouble();
            output.FadeDirectionAngle60000 = Angle(((fadeDegrees % 360) + 360) % 360) % 21_600_000;
        }
        if (value.TryGetProperty("scaleX", out var scaleX))
            output.ScaleXThousandthPercent = checked((int)Math.Round(
                ChartEffectNumber(scaleX.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d,
                MidpointRounding.ToEven));
        if (value.TryGetProperty("scaleY", out var scaleY))
            output.ScaleYThousandthPercent = checked((int)Math.Round(
                ChartEffectNumber(scaleY.GetDouble(), int.MinValue / 100000d, int.MaxValue / 100000d) * 100000d,
                MidpointRounding.ToEven));
        if (value.TryGetProperty("skewX", out var skewX))
            output.SkewXAngle60000 = checked((int)Math.Round(
                ChartEffectNumber(skewX.GetDouble(), -90, 90) * 60_000d,
                MidpointRounding.ToEven));
        if (value.TryGetProperty("skewY", out var skewY))
            output.SkewYAngle60000 = checked((int)Math.Round(
                ChartEffectNumber(skewY.GetDouble(), -90, 90) * 60_000d,
                MidpointRounding.ToEven));
        if (value.TryGetProperty("alignment", out var alignment))
            output.Alignment = alignment.GetString()!;
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

    private static string NormalizeThemeColor(string value)
    {
        var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
        if (normalized.Length is not (6 or 8) || normalized.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_presentation_theme", $"Presentation theme color {value} must be a six- or eight-digit RGB/RGBA value.");
        return "#" + normalized;
    }

    private static string MixRgb(string source, string target, double amount)
    {
        var t = Math.Clamp(amount, 0, 1);
        var r = MixChannel(source[..2], target[..2], t);
        var g = MixChannel(source[2..4], target[2..4], t);
        var b = MixChannel(source[4..6], target[4..6], t);
        return $"{r:X2}{g:X2}{b:X2}";
    }

    private static byte MixChannel(string source, string target, double amount)
    {
        var from = byte.Parse(source, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var to = byte.Parse(target, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return checked((byte)Math.Round(from + (to - from) * amount, MidpointRounding.AwayFromZero));
    }

    private static string NativeAssetId(string mimeType, string sha256) =>
        PptxAssetCatalog.NativeAssetIdFor(mimeType, sha256);
}
