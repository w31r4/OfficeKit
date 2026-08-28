using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeKit.Codec;

internal sealed record PptxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record PptxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record PptxLayoutGraphEntry(int Index, string Id, string RelationshipId, SlideLayoutPart Part);
internal sealed record PptxMasterGraphEntry(int Index, string Id, string RelationshipId, SlideMasterPart Part, IReadOnlyList<PptxLayoutGraphEntry> Layouts);
internal sealed record PptxSourceSlideEntry(int Index, P.SlideId SlideId, string RelationshipId, SlidePart Part);
internal sealed class PptxTargetSlideEntry
{
    internal PptxTargetSlideEntry(int targetIndex, PresentationSlide target, PptxSourceSlideEntry source, bool isClone)
    {
        TargetIndex = targetIndex;
        Target = target;
        Source = source;
        IsClone = isClone;
        OutputSlideId = source.SlideId;
        OutputPart = source.Part;
    }

    internal int TargetIndex { get; }
    internal PresentationSlide Target { get; }
    internal PptxSourceSlideEntry Source { get; }
    internal bool IsClone { get; }
    internal P.SlideId OutputSlideId { get; set; }
    internal SlidePart OutputPart { get; set; }
}

internal static class PptxCodec
{
    internal static bool SupportsBoundTextLeaf(P.Shape shape) =>
        shape.TextBody is not null && PptxTextCodec.SupportsEditing(shape.TextBody);

    internal static int ValidateEditPlanOutput(
        byte[] sourceBytes,
        byte[] outputBytes,
        EffectiveCodecLimits limits)
    {
        ValidateOutputBudget(outputBytes, limits);
        return ValidateOffice2021AgainstSource(sourceBytes, outputBytes);
    }

    private const long DefaultSlideWidthEmu = 12_192_000;
    private const long DefaultSlideHeightEmu = 6_858_000;
    internal static PptxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx, out var packagePaths);
        var nativeObjects = new PptxNativeObjectCatalog(opaque, packagePaths, limits);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships for source-bound, fail-closed export from the validated package snapshot.",
                opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));

        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        var presentationRoot = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no presentation root.", "ppt/presentation.xml");
        var slideIds = presentationRoot.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
        if ((uint)slideIds.Length > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"PPTX presentation has {slideIds.Length} slides and exceeds max_sheets ({limits.MaxSheets}).", "ppt/presentation.xml");
        var slideParts = ResolveSlideParts(presentationPart, slideIds);
        var publicSlideIds = slideIds.Select((_, index) => $"presentation/slide/{index + 1}").ToArray();
        var publicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(slideIds
            .Select((slideId, index) => (
                RelationshipId: slideId.RelationshipId?.Value ?? string.Empty,
                PublicId: publicSlideIds[index])));
        var publicSlideIdByNativeId = BuildSectionSlideIdMap(slideIds
            .Select((slideId, index) => (
                NativeId: slideId.Id?.Value,
                PublicId: publicSlideIds[index])));
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: $"presentation/slide/{index + 1}"))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var assetCatalog = new PptxAssetCatalog([], limits, nativeObjects.ValidatedPartSha256);
        var masterGraph = ReadMasterGraph(presentationPart);
        var layoutIdByPartPath = masterGraph
            .SelectMany(master => master.Layouts)
            .ToDictionary(layout => PartPath(layout.Part), layout => layout.Id, StringComparer.OrdinalIgnoreCase);

        var artifact = new PresentationArtifact
        {
            Id = "presentation/1",
            Name = "Imported presentation",
            SlideWidthEmu = presentationRoot.SlideSize?.Cx?.Value ?? DefaultSlideWidthEmu,
            SlideHeightEmu = presentationRoot.SlideSize?.Cy?.Value ?? DefaultSlideHeightEmu,
        };
        artifact.ViewProperties = PptxViewPropertiesCodec.Read(presentationPart);
        ulong semanticItems = 0;
        var customShows = PptxCustomShowCodec.Read(presentationPart, publicSlideIdByRelationshipId, limits);
        semanticItems = checked(semanticItems + customShows.SemanticItems);
        if (semanticItems > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", "ppt/presentation.xml");
        artifact.CustomShows.Add(customShows.Shows);
        artifact.CustomShowsOpaque = customShows.Opaque;
        if (customShows.Opaque)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_presentation_custom_shows_retained",
                $"Retained an unsupported custom-show graph without exposing incomplete editable semantics: {customShows.Reason}.",
                "ppt/presentation.xml"));
        var sections = PptxSectionCodec.Read(presentationPart, publicSlideIdByNativeId, publicSlideIds, limits);
        semanticItems = checked(semanticItems + sections.SemanticItems);
        if (semanticItems > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", "ppt/presentation.xml");
        artifact.Sections.Add(sections.Sections);
        artifact.SectionsOpaque = sections.Opaque;
        if (sections.Opaque)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_presentation_sections_retained",
                $"Retained an unsupported PowerPoint section graph without exposing incomplete editable semantics: {sections.Reason}.",
                "ppt/presentation.xml"));
        var customShowCatalog = PptxCustomShowCatalog.From(customShows.Shows);
        foreach (var master in masterGraph)
        {
            var masterRoot = master.Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"Presentation master {master.Index + 1} has no slide master root.", PartPath(master.Part));
            var masterCommon = masterRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation master {master.Index + 1} has no common slide data.", PartPath(master.Part));
            var masterShapeTree = masterCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation master {master.Index + 1} has no shape tree.", PartPath(master.Part));
            var masterContext = new PptxPartContext(master.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
            var textStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
            var background = PptxBackgroundCodec.Read(masterCommon);
            var masterArtifact = new PresentationMaster
            {
                Id = master.Id,
                Name = masterCommon.Name?.Value ?? $"Master {master.Index + 1}",
                TextStyles = textStyles,
                Source = new PresentationMasterSourceBinding
                {
                    MasterIndex = checked((uint)master.Index),
                    PartPath = PartPath(master.Part),
                    RelationshipId = master.RelationshipId,
                    MasterXmlSha256 = HashElement(masterRoot),
                    TextStylesSemanticSha256 = MasterTextStylesSemanticHash(textStyles),
                    TextStylesEditable = PptxMasterTextStylesCodec.Supports(masterRoot),
                    BackgroundSemanticSha256 = BackgroundSemanticHash(background),
                    BackgroundEditable = PptxBackgroundCodec.Supports(masterCommon),
                },
            };
            if (background is not null) masterArtifact.Background = background;
            masterArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(masterShapeTree, master.Id, masterContext));
            artifact.Masters.Add(masterArtifact);
            foreach (var layout in master.Layouts)
            {
                var layoutRoot = layout.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no slide layout root.", PartPath(layout.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no common slide data.", PartPath(layout.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no shape tree.", PartPath(layout.Part));
                var layoutContext = new PptxPartContext(layout.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
                var layoutBackground = PptxBackgroundCodec.Read(layoutCommon);
                var layoutArtifact = new PresentationLayout
                {
                    Id = layout.Id,
                    Name = layoutCommon.Name?.Value ?? $"Layout {layout.Index + 1}",
                    MasterId = master.Id,
                    Type = LayoutTypeName(layoutRoot),
                    Source = new PresentationLayoutSourceBinding
                    {
                        LayoutIndex = checked((uint)layout.Index),
                        PartPath = PartPath(layout.Part),
                        RelationshipId = layout.RelationshipId,
                        LayoutXmlSha256 = HashElement(layoutRoot),
                        BackgroundSemanticSha256 = BackgroundSemanticHash(layoutBackground),
                        BackgroundEditable = PptxBackgroundCodec.Supports(layoutCommon),
                    },
                };
                if (layoutBackground is not null) layoutArtifact.Background = layoutBackground;
                layoutArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(layoutShapeTree, layout.Id, layoutContext));
                artifact.Layouts.Add(layoutArtifact);
                layout.Part.UnloadRootElement();
            }
            master.Part.UnloadRootElement();
        }
        for (var slideIndex = 0; slideIndex < slideIds.Length; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var relationshipId = slideId.RelationshipId?.Value ?? string.Empty;
            var slidePart = slideParts[slideIndex];
            var slideRoot = slidePart.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
            var slideCommon = slideRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
            var shapeTree = slideCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
            var slideBackground = PptxBackgroundCodec.Read(slideCommon);
            var slideTransition = PptxTransitionCodec.Read(slideRoot);
            var slideVisibility = PptxSlideVisibilityCodec.Read(slideRoot);
            var elements = ShapeElements(shapeTree);
            var deletionAnalysis = PptxElementDeletionCodec.AnalyzeSlide(slidePart);
            var zOrderPlan = AnalyzeElementZOrder(elements);
            var slideArtifactId = $"presentation/slide/{slideIndex + 1}";
            var elementIdsByNativeId = NativeElementIds(elements, slideArtifactId);
            P.Slide? previousSlideRoot = null;
            IReadOnlyDictionary<uint, string>? previousElementIdsByNativeId = null;
            string? previousSlideArtifactId = null;
            if (slideIndex > 0)
            {
                previousSlideRoot = slideParts[slideIndex - 1].Slide;
                var previousShapeTree = previousSlideRoot?.CommonSlideData?.ShapeTree;
                previousSlideArtifactId = $"presentation/slide/{slideIndex}";
                if (previousShapeTree is not null)
                    previousElementIdsByNativeId = NativeElementIds(ShapeElements(previousShapeTree), previousSlideArtifactId);
            }
            var slideTiming = PptxTimingCodec.Read(
                slideRoot,
                elementIdsByNativeId,
                previousSlideRoot,
                previousElementIdsByNativeId,
                previousSlideArtifactId);
            var sourceEntry = new PptxSourceSlideEntry(slideIndex, slideId, relationshipId, slidePart);
            var deletionPlan = PptxSlideDeletionCodec.Analyze(presentationPart, sourceEntry, opaque);
            var clonePlan = PptxSlideCloneCodec.Analyze(presentationPart, sourceEntry, slideParts.ToHashSet());
            var target = new PresentationSlide
            {
                Id = slideArtifactId,
                // An absent p:cSld/@name is a real source value. Do not invent
                // a display name here: source-preserving slide clone and edit
                // checks must be able to distinguish absent from authored text.
                Name = slideRoot.CommonSlideData?.Name?.Value ?? string.Empty,
                LayoutId = slidePart.SlideLayoutPart is { } layoutPart
                    ? layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart)) ??
                      throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart))
                    : string.Empty,
                Source = new PresentationSlideSourceBinding
                {
                    SlideIndex = checked((uint)slideIndex),
                    PartPath = PartPath(slidePart),
                    RelationshipId = relationshipId,
                    SlideXmlSha256 = HashElement(slideRoot),
                    LayoutRelationshipId = slidePart.SlideLayoutPart is { } boundLayout ? slidePart.GetIdOfPart(boundLayout) : string.Empty,
                    BackgroundSemanticSha256 = BackgroundSemanticHash(slideBackground),
                    BackgroundEditable = PptxBackgroundCodec.Supports(slideCommon),
                    SpeakerNotesAddable = PptxSpeakerNotesCodec.CanAddSourceBound(presentationPart, slidePart),
                    LegacyCommentsAddable = PptxLegacyCommentsCodec.CanAddSourceBound(presentationPart, slidePart),
                    LegacyCommentsEditable = PptxLegacyCommentsCodec.CanEditSourceBound(presentationPart, slidePart, slideIndex),
                    CommentPartPresent = PptxLegacyCommentsCodec.CommentPartPresent(slidePart),
                    CommentFamily = PptxLegacyCommentsCodec.CommentFamily(presentationPart),
                    TransitionSemanticSha256 = PptxTransitionCodec.SemanticHash(slideTransition),
                    TransitionEditable = PptxTransitionCodec.Supports(slideRoot),
                    TransitionPresent = PptxTransitionCodec.HasTransition(slideRoot),
                    TransitionAddable = PptxTransitionCodec.CanAdd(slideRoot),
                    TimingPresent = slideTiming.Present,
                    TimingEditable = slideTiming.Editable,
                    TimingAddable = slideTiming.Addable,
                    TimingSemanticSha256 = slideTiming.SemanticSha256,
                    VisibilitySemanticSha256 = slideVisibility.SemanticSha256,
                    VisibilityEditable = slideVisibility.Editable,
                    DeletionCapability = new PresentationSlideDeletionCapability
                    {
                        Supported = deletionPlan.Supported,
                        BlockedReason = deletionPlan.BlockedReason,
                        OwnedPartCount = deletionPlan.OwnedPartCount,
                    },
                    CloneCapability = new PresentationSlideCloneCapability
                    {
                        Supported = clonePlan.Supported,
                        BlockedReason = clonePlan.BlockedReason,
                        ClonedPartCount = clonePlan.ClonedPartCount,
                        SharedPartCount = clonePlan.SharedPartCount,
                    },
                },
            };
            if (slideVisibility.Hidden is { } hidden) target.Hidden = hidden;
            if (slideBackground is not null) target.Background = slideBackground;
            if (slideTransition is not null) target.Transition = slideTransition;
            target.Animations.AddRange(slideTiming.Animations);
            if (slideTiming.Morph is not null) target.Morph = slideTiming.Morph;
            if (PptxSpeakerNotesCodec.Read(slidePart) is { } speakerNotes)
                target.SpeakerNotes = speakerNotes;
            target.LegacyComments.Add(PptxLegacyCommentsCodec.Read(presentationPart, slidePart, slideIndex, diagnostics));
            target.ModernComments.Add(PptxModernCommentsCodec.Read(
                presentationPart,
                slideId,
                slidePart,
                elements,
                elementIdsByNativeId,
                slideIndex,
                diagnostics));
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                semanticItems++;
                if (semanticItems > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                var importedElement = ReadElement(elements[elementIndex], slideIndex, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                if (importedElement.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    semanticItems += checked((ulong)importedElement.Table.Rows.Sum(row => row.Cells.Count));
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                }
                SetElementDeletionCapability(
                    importedElement,
                    PptxElementDeletionCodec.Analyze(slidePart, elements[elementIndex], elements, deletionAnalysis));
                SetElementZOrderCapability(importedElement, zOrderPlan);
                target.Elements.Add(importedElement);
            }
            artifact.Slides.Add(target);
            if (slideIndex > 0)
                slideParts[slideIndex - 1].UnloadRootElement();
        }
        if (slideParts.Length > 0)
            slideParts[^1].UnloadRootElement();
        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Presentation = artifact,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "pptx",
                PackageSha256 = Hash(bytes),
                Producer = "office-kit/OfficeKit",
            },
        };
        envelope.Assets.Add(assetCatalog.ImportedAssets);
        envelope.Diagnostics.Add(diagnostics);
        return new PptxImportResult(envelope, diagnostics);
    }

    internal static PptxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        var requiresSourcePreservation =
            envelope.ProtocolVersion == CodecProtocol.ProtocolVersion &&
            envelope.Family == ArtifactFamily.Presentation &&
            envelope.PayloadCase == ArtifactEnvelope.PayloadOneofCase.Presentation &&
            RequiresSourcePreservation(envelope);
        if (requiresSourcePreservation && envelope.OpaqueOpc?.SourcePackage is not { Data.IsEmpty: false })
            throw new CodecException(
                "missing_source_package",
                "Source-bound PPTX export requires its validated source package snapshot.");

        var assetCatalog = ValidateEnvelope(envelope, limits);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (requiresSourcePreservation)
            return ExportPreservingSource(envelope, limits, opaqueCount, assetCatalog);

        var diagnostics = new List<Diagnostic>();

        using var stream = new MemoryStream();
        using (var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, autoSave: true))
            BuildPresentation(package, envelope.Presentation, assetCatalog);
        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new PptxExportResult(bytes, diagnostics);
    }

    private static bool RequiresSourcePreservation(ArtifactEnvelope envelope)
    {
        if (envelope.Source is not null) return true;
        if (envelope.OpaqueOpc is { } opaque &&
            (opaque.SourcePackage is not null || opaque.Parts.Count > 0 || opaque.PackageRelationships.Count > 0))
            return true;

        var presentation = envelope.Presentation;
        return presentation.Masters.Any(master =>
                   master.Source is not null || master.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Layouts.Any(layout =>
                   layout.Source is not null || layout.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Slides.Any(slide =>
                   slide.Source is not null || slide.Elements.Any(element =>
                       element.Source is not null || element.ContentCase == PresentationElement.ContentOneofCase.Opaque)) ||
               presentation.ViewProperties?.Source is not null ||
               presentation.CustomShowsOpaque ||
               presentation.CustomShows.Any(show => show.Source is not null) ||
               presentation.SectionsOpaque ||
               presentation.Sections.Any(section => section.Source is not null);
    }

    private static PptxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount, PptxAssetCatalog assetCatalog)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(envelope.OpaqueOpc, envelope.Source, limits, OpcPackageProfile.Pptx);
        var nativeObjects = new PptxNativeObjectCatalog(envelope.OpaqueOpc, sourceBytes, limits);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        var changedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var addedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var replacedOpaquePartHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var removedSourcePartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var removedSourceRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        var removedElementPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var removedElementRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        var clonedPartSourcePaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var clonedPackageEntryPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var authoredOverlayXmlByPartPath = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        using (var package = PresentationDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var presentationPart = package.PresentationPart ??
                throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
            var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
            var targetSlides = BindSourcePreservingSlides(presentationPart, slideIds, envelope.Presentation.Slides);
            AssertCloneOriginsRetained(targetSlides);
            var retainedTargets = targetSlides.Where(target => !target.IsClone).ToArray();
            var slideParts = retainedTargets.Select(target => target.Source.Part).ToArray();
            var slideIdByPartPath = retainedTargets
                .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
                .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
            var slidePartById = retainedTargets
                .Select(target => (Part: target.Source.Part, Id: target.Target.Id))
                .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
            var customShowCatalog = PptxCustomShowCatalog.From(envelope.Presentation.CustomShows);
            var retainedPublicSlideIdByRelationshipId = retainedTargets
                .Select(target => (target.Source.RelationshipId, target.Target.Id))
                .ToDictionary(item => item.RelationshipId, item => item.Id, StringComparer.Ordinal);
            var sourcePublicSlideIds = slideIds
                .Select(slideId => retainedPublicSlideIdByRelationshipId.GetValueOrDefault(slideId.RelationshipId?.Value ?? string.Empty) ?? string.Empty)
                .ToArray();
            var sourcePublicSlideIdByNativeId = BuildSectionSlideIdMap(slideIds
                .Select(slideId => (
                    NativeId: slideId.Id?.Value,
                    PublicId: retainedPublicSlideIdByRelationshipId.GetValueOrDefault(slideId.RelationshipId?.Value ?? string.Empty) ?? string.Empty)));
            if (targetSlides.Any(target => target.IsClone))
            {
                var sourcePublicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(retainedTargets.Select(target => (
                    RelationshipId: target.Source.RelationshipId,
                    PublicId: target.Target.Id)));
                PptxCustomShowCodec.AssertMembershipUnchangedForSlideClone(
                    presentationPart,
                    envelope.Presentation,
                    sourcePublicSlideIdByRelationshipId,
                    limits);
                PptxSectionCodec.AssertNoSectionCloneCombination(
                    presentationPart,
                    envelope.Presentation,
                    sourcePublicSlideIdByNativeId,
                    sourcePublicSlideIds,
                    limits);
            }
            var masterGraph = ReadMasterGraph(presentationPart);
            if (masterGraph.Length != envelope.Presentation.Masters.Count)
                throw new CodecException(
                    "presentation_master_topology_changed",
                    $"Source-preserving PPTX export requires the original {masterGraph.Length}-master topology; the artifact contains {envelope.Presentation.Masters.Count} masters.",
                    "ppt/presentation.xml");
            var layoutGraph = masterGraph.SelectMany(master => master.Layouts.Select(layout => (Master: master, Layout: layout))).ToArray();
            if (layoutGraph.Length != envelope.Presentation.Layouts.Count)
                throw new CodecException(
                    "presentation_layout_topology_changed",
                    $"Source-preserving PPTX export requires the original {layoutGraph.Length}-layout topology; the artifact contains {envelope.Presentation.Layouts.Count} layouts.",
                    "ppt/presentation.xml");
            var layoutIdByPartPath = layoutGraph.ToDictionary(item => PartPath(item.Layout.Part), item => item.Layout.Id, StringComparer.OrdinalIgnoreCase);
            if (PptxViewPropertiesCodec.ApplySourceBound(presentationPart, envelope.Presentation.ViewProperties) is { } viewPropertiesChange)
            {
                changedParts.Add(viewPropertiesChange.PartPath);
                replacedOpaquePartHashes.Add(viewPropertiesChange.PartPath, viewPropertiesChange.Sha256);
            }
            PptxLegacyCommentsCodec.AssertSourceUnchanged(presentationPart, slideParts, retainedTargets.Select(target => target.Target).ToArray());
            CloneRequestedSourceSlides(
                presentationPart,
                targetSlides,
                layoutIdByPartPath,
                slideIdByPartPath,
                assetCatalog,
                customShowCatalog,
                nativeObjects,
                changedParts,
                addedRelationshipIds,
                addedPartPaths,
                clonedPackageEntryPaths,
                clonedPartSourcePaths);
            DeleteUnrequestedSourceSlides(
                presentationPart,
                slideIds,
                targetSlides,
                envelope.OpaqueOpc,
                changedParts,
                removedSourcePartPaths);
            if (ReorderSourceSlideIdList(presentationPart, targetSlides))
                changedParts.Add(PartPath(presentationPart));
            if (ApplySourceBoundSlideSize(presentationPart, envelope.Presentation))
                changedParts.Add(PartPath(presentationPart));
            if (presentationPart.Presentation?.CustomShowList is not null ||
                envelope.Presentation.CustomShowsOpaque ||
                envelope.Presentation.CustomShows.Count > 0)
            {
                var outputSlideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
                if (outputSlideIds.Length != envelope.Presentation.Slides.Count)
                    throw new CodecException(
                        "presentation_slide_topology_changed",
                        "Source-preserving PPTX export could not bind custom shows to the requested slide topology.",
                        "ppt/presentation.xml");
                var publicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(outputSlideIds
                    .Select((slideId, index) => (
                        RelationshipId: slideId.RelationshipId?.Value ?? string.Empty,
                        PublicId: envelope.Presentation.Slides[index].Id)));
                if (PptxCustomShowCodec.ApplySourceBound(presentationPart, envelope.Presentation, publicSlideIdByRelationshipId, limits))
                    changedParts.Add(PartPath(presentationPart));
            }
            var outputSectionSlideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
            if (outputSectionSlideIds.Length != envelope.Presentation.Slides.Count)
                throw new CodecException(
                    "presentation_slide_topology_changed",
                    "Source-preserving PPTX export could not bind PowerPoint sections to the requested slide topology.",
                    "ppt/presentation.xml");
            var outputPublicSlideIdByNativeId = BuildSectionSlideIdMap(outputSectionSlideIds
                .Select((slideId, index) => (
                    NativeId: slideId.Id?.Value,
                    PublicId: envelope.Presentation.Slides[index].Id)));
            if (PptxSectionCodec.ApplySourceBound(
                    presentationPart,
                    envelope.Presentation,
                    outputPublicSlideIdByNativeId,
                    envelope.Presentation.Slides.Select(slide => slide.Id).ToArray(),
                    limits))
                changedParts.Add(PartPath(presentationPart));
            assetCatalog.IndexExistingParts(slideParts.SelectMany(part => part.ImageParts)
                .Concat(masterGraph.SelectMany(master => master.Part.Parts.Select(pair => pair.OpenXmlPart).OfType<ImagePart>())));

            for (var masterIndex = 0; masterIndex < masterGraph.Length; masterIndex++)
            {
                var graph = masterGraph[masterIndex];
                var masterRoot = graph.Part.SlideMaster ??
                    throw new CodecException("missing_slide_master_root", $"Presentation master {masterIndex + 1} has no slide master root.", PartPath(graph.Part));
                var masterCommon = masterRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation master {masterIndex + 1} has no common slide data.", PartPath(graph.Part));
                var masterShapeTree = masterCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation master {masterIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Masters[masterIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_master_binding",
                    $"Presentation master {masterIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                if (target.Id != graph.Id ||
                    binding.MasterIndex != masterIndex ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.MasterXmlSha256.Equals(HashElement(masterRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_binding_mismatch",
                        $"Presentation master {masterIndex + 1} does not match its hash-bound source master.",
                        PartPath(graph.Part));
                var sourceName = masterCommon.Name?.Value ?? $"Master {masterIndex + 1}";
                if (!target.Name.Equals(sourceName, StringComparison.Ordinal))
                    throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export cannot rename master {masterIndex + 1}.", PartPath(graph.Part));
                var masterContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
                var originalStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
                var originalSemanticHash = MasterTextStylesSemanticHash(originalStyles);
                if (!binding.TextStylesSemanticSha256.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_semantics_mismatch",
                        $"Presentation master {masterIndex + 1} text styles do not match their source binding.",
                        PartPath(graph.Part));
                var requestedSemanticHash = MasterTextStylesSemanticHash(target.TextStyles);
                if (!requestedSemanticHash.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.TextStylesEditable || !PptxMasterTextStylesCodec.Supports(masterRoot))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} text styles are preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxMasterTextStylesCodec.Apply(masterRoot, target.TextStyles ?? new PresentationMasterTextStyles(), masterContext);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                var originalBackground = PptxBackgroundCodec.Read(masterCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_background_mismatch",
                        $"Presentation master {masterIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(masterCommon))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(masterCommon, target.Background);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(masterShapeTree, graph.Id, target.Placeholders, masterContext, PartPath(graph.Part)))
                {
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, masterContext, changedParts, addedRelationshipIds, addedPartPaths);
            }

            for (var layoutIndex = 0; layoutIndex < layoutGraph.Length; layoutIndex++)
            {
                var (master, graph) = layoutGraph[layoutIndex];
                var layoutRoot = graph.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layoutIndex + 1} has no slide layout root.", PartPath(graph.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layoutIndex + 1} has no common slide data.", PartPath(graph.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layoutIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Layouts[layoutIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_layout_binding",
                    $"Presentation layout {layoutIndex + 1} is missing its source binding.",
                    PartPath(graph.Part));
                var sourceName = layoutCommon.Name?.Value ?? $"Layout {graph.Index + 1}";
                if (target.Id != graph.Id || target.MasterId != master.Id || target.Name != sourceName || target.Type != LayoutTypeName(layoutRoot) ||
                    binding.LayoutIndex != graph.Index ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.LayoutXmlSha256.Equals(HashElement(layoutRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_binding_mismatch",
                        $"Presentation layout {layoutIndex + 1} does not match its hash-bound read-only source layout.",
                        PartPath(graph.Part));
                var layoutContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
                var originalBackground = PptxBackgroundCodec.Read(layoutCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_source_background_mismatch",
                        $"Presentation layout {layoutIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(layoutCommon))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation layout {layoutIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(layoutCommon, target.Background);
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(layoutShapeTree, graph.Id, target.Placeholders, layoutContext, PartPath(graph.Part)))
                {
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, layoutContext, changedParts, addedRelationshipIds, addedPartPaths);
            }

            ulong semanticItems = 0;
            for (var targetPosition = 0; targetPosition < retainedTargets.Length; targetPosition++)
            {
                var targetSlide = retainedTargets[targetPosition];
                var slideIndex = targetSlide.TargetIndex;
                var relationshipId = targetSlide.Source.RelationshipId;
                var slidePart = targetSlide.Source.Part;
                var slideRoot = slidePart.Slide ??
                    throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
                var target = targetSlide.Target;
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_slide_binding",
                    $"Presentation slide {slideIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                var sourceVisibility = PptxSlideVisibilityCodec.Read(slideRoot);
                if (binding.SlideIndex != targetSlide.Source.Index ||
                    !binding.PartPath.Equals(PartPath(slidePart), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(relationshipId, StringComparison.Ordinal) ||
                    !binding.SlideXmlSha256.Equals(HashElement(slideRoot), StringComparison.OrdinalIgnoreCase) ||
                    binding.SpeakerNotesAddable != PptxSpeakerNotesCodec.CanAddSourceBound(presentationPart, slidePart) ||
                    binding.LegacyCommentsAddable != PptxLegacyCommentsCodec.CanAddSourceBound(presentationPart, slidePart) ||
                    binding.LegacyCommentsEditable != PptxLegacyCommentsCodec.CanEditSourceBound(presentationPart, slidePart, targetSlide.Source.Index) ||
                    binding.CommentPartPresent != PptxLegacyCommentsCodec.CommentPartPresent(slidePart) ||
                    !binding.CommentFamily.Equals(PptxLegacyCommentsCodec.CommentFamily(presentationPart), StringComparison.Ordinal) ||
                    binding.TransitionEditable != PptxTransitionCodec.Supports(slideRoot) ||
                    binding.TransitionPresent != PptxTransitionCodec.HasTransition(slideRoot) ||
                    binding.TransitionAddable != PptxTransitionCodec.CanAdd(slideRoot) ||
                    binding.VisibilityEditable != sourceVisibility.Editable ||
                    !binding.VisibilitySemanticSha256.Equals(sourceVisibility.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} does not match its hash-bound source slide.",
                        PartPath(slidePart));
                var sourceLayoutPart = slidePart.SlideLayoutPart;
                var sourceLayoutId = sourceLayoutPart is null ? string.Empty :
                    layoutIdByPartPath.GetValueOrDefault(PartPath(sourceLayoutPart)) ??
                    throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart));
                var sourceLayoutRelationshipId = sourceLayoutPart is null ? string.Empty : slidePart.GetIdOfPart(sourceLayoutPart);
                if (target.LayoutId != sourceLayoutId || binding.LayoutRelationshipId != sourceLayoutRelationshipId)
                    throw new CodecException(
                        "presentation_slide_layout_binding_changed",
                        $"Source-preserving PPTX export cannot change slide {slideIndex + 1}'s layout binding.",
                        PartPath(slidePart));

                var slideCommon = slideRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
                var shapeTree = slideCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
                var originalBackground = PptxBackgroundCodec.Read(slideCommon);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_background_mismatch",
                        $"Presentation slide {slideIndex + 1} background does not match its source binding.",
                        PartPath(slidePart));
                var changed = false;
                if (sourceVisibility.Editable != target.HasHidden)
                    throw new CodecException(
                        "presentation_slide_visibility_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} visibility no longer matches its source capability.",
                        PartPath(slidePart));
                if (target.HasHidden && target.Hidden != sourceVisibility.Hidden && PptxSlideVisibilityCodec.ApplySourceBound(slideRoot, target))
                    changed = true;
                var sourceName = slideCommon.Name?.Value ?? string.Empty;
                if (!string.Equals(target.Name, sourceName, StringComparison.Ordinal))
                {
                    // This is deliberately the only source-bound slide metadata
                    // mutation: p:cSld/@name belongs to the existing SlidePart
                    // and does not alter its relationship graph or shape tree.
                    slideCommon.Name = target.Name;
                    changed = true;
                }
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(slideCommon))
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} background is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxBackgroundCodec.Apply(slideCommon, target.Background);
                    changed = true;
                }
                var originalTransition = PptxTransitionCodec.Read(slideRoot);
                var originalTransitionHash = PptxTransitionCodec.SemanticHash(originalTransition);
                if (!binding.TransitionSemanticSha256.Equals(originalTransitionHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_transition_mismatch",
                        $"Presentation slide {slideIndex + 1} transition does not match its source binding.",
                        PartPath(slidePart));
                if (!PptxTransitionCodec.SemanticHash(target.Transition).Equals(originalTransitionHash, StringComparison.OrdinalIgnoreCase))
                {
                    var canAddTransition = binding.TransitionAddable && PptxTransitionCodec.CanAdd(slideRoot);
                    if ((!binding.TransitionEditable || !PptxTransitionCodec.Supports(slideRoot)) && !canAddTransition)
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} transition is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxTransitionCodec.Apply(slideRoot, target.Transition);
                    changed = true;
                }
                var sourceElements = ShapeElements(shapeTree);
                var zOrderPlan = AnalyzeElementZOrder(sourceElements);
                // Timing edits can change whether a native shape is safe to
                // delete, but the imported element binding is about the
                // source revision. Capture that contract before replacing the
                // timing tree so an animation edit does not masquerade as a
                // shape-capability mutation.
                var deletionAnalysis = PptxElementDeletionCodec.AnalyzeSlide(slidePart);
                var sourceDeletionPlans = sourceElements
                    .Select(element => PptxElementDeletionCodec.Analyze(slidePart, element, sourceElements, deletionAnalysis))
                    .ToArray();
                var (retainedElements, authoredElements) = SplitSourceBoundElements(target, sourceElements.Length, slideIndex, slidePart);
                var elementIdsByNativeId = NativeElementIds(sourceElements, target.Id);
                var nativeIdsByElementId = elementIdsByNativeId.ToDictionary(item => item.Value, item => item.Key, StringComparer.Ordinal);
                P.Slide? previousSourceRoot = null;
                IReadOnlyDictionary<uint, string>? previousSourceElementIds = null;
                string? previousSourceSlideId = null;
                if (targetSlide.Source.Index > 0 && PptxTimingCodec.HasMorph(slideRoot))
                {
                    var previousSourceTarget = retainedTargets.FirstOrDefault(candidate =>
                        candidate.Source.Index == targetSlide.Source.Index - 1);
                    if (previousSourceTarget is null)
                        throw new CodecException(
                            "invalid_presentation_morph",
                            $"Presentation slide {slideIndex + 1} cannot retain Morph after its source predecessor was removed.",
                            PartPath(slidePart));
                    previousSourceRoot = previousSourceTarget.Source.Part.Slide;
                    var previousSourceShapeTree = previousSourceRoot?.CommonSlideData?.ShapeTree;
                    previousSourceSlideId = previousSourceTarget.Target.Id;
                    if (previousSourceShapeTree is not null)
                        previousSourceElementIds = NativeElementIds(ShapeElements(previousSourceShapeTree), previousSourceSlideId);
                }
                var originalTiming = PptxTimingCodec.Read(
                    slideRoot,
                    elementIdsByNativeId,
                    previousSourceRoot,
                    previousSourceElementIds,
                    previousSourceSlideId);
                if (!binding.TimingSemanticSha256.Equals(originalTiming.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_timing_mismatch",
                        $"Presentation slide {slideIndex + 1} timing does not match its source binding.",
                        PartPath(slidePart));
                PptxTimingCodec.ValidateMorphContext(target, slideIndex > 0 ? targetSlides[slideIndex - 1].Target : null);
                var requestedTimingHash = PptxTimingCodec.SemanticHash(target.Animations, target.Morph);
                var requestedOpaqueNoop = originalTiming.Present && !originalTiming.Editable && target.Animations.Count == 0 && target.Morph is null;
                if (!requestedOpaqueNoop && !requestedTimingHash.Equals(originalTiming.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                {
                    var canReplaceTiming = originalTiming.Editable || (originalTiming.Addable && (target.Animations.Count > 0 || target.Morph is not null));
                    if (!canReplaceTiming)
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} timing is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxTimingCodec.Apply(slideRoot, target, nativeIdsByElementId, allowOpaqueReplacement: originalTiming.Editable || originalTiming.Addable);
                    changed = true;
                }
                foreach (var (elementId, nativeId) in AuthoredOverlayNativeIds(sourceElements, authoredElements, slideIndex, slidePart))
                    if (!nativeIdsByElementId.TryAdd(elementId, nativeId))
                        throw new CodecException(
                            "invalid_presentation_element",
                            $"Presentation slide {slideIndex + 1} authored overlay element {elementId} reuses an existing source element identity.",
                            PartPath(slidePart));

                var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
                var requestedBySourceIndex = new Dictionary<int, PresentationElement>();
                foreach (var requested in retainedElements)
                {
                    var requestedBinding = requested.Source ?? throw new CodecException(
                        "missing_presentation_element_binding",
                        $"Presentation slide {slideIndex + 1} retained element {requested.Id} is missing its source binding.",
                        PartPath(slidePart));
                    if (requestedBinding.ShapeTreeIndex >= (uint)sourceElements.Length)
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} retained element {requested.Id} identifies a source shape-tree index outside the slide.",
                            PartPath(slidePart));
                    var sourceIndex = (int)requestedBinding.ShapeTreeIndex;
                    if (!requestedBySourceIndex.TryAdd(sourceIndex, requested))
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} retained elements must identify unique source shape-tree nodes.",
                            PartPath(slidePart));
                }
                var deletionsBySourceIndex = new Dictionary<int, PresentationElementDeletion>();
                foreach (var deletion in target.ElementDeletions)
                {
                    var deletionBinding = deletion.Source ?? throw new CodecException(
                        "missing_presentation_element_deletion_binding",
                        $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} is missing its source binding.",
                        PartPath(slidePart));
                    if (deletionBinding.ShapeTreeIndex >= (uint)sourceElements.Length)
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} identifies a source shape-tree index outside the slide.",
                            PartPath(slidePart));
                    var sourceIndex = (int)deletionBinding.ShapeTreeIndex;
                    if (requestedBySourceIndex.ContainsKey(sourceIndex) || !deletionsBySourceIndex.TryAdd(sourceIndex, deletion))
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} does not identify one omitted source element.",
                            PartPath(slidePart));
                }
                var requestedSourceOrder = retainedElements
                    .Select(element => checked((int)element.Source!.ShapeTreeIndex))
                    .ToArray();
                var sourceOrderChanged = !requestedSourceOrder.SequenceEqual(Enumerable.Range(0, sourceElements.Length));
                if (sourceOrderChanged && (authoredElements.Length > 0 || deletionsBySourceIndex.Count > 0))
                    throw new CodecException(
                        "unsupported_presentation_element_reorder",
                        $"Presentation slide {slideIndex + 1} cannot combine direct-element reordering with authored overlays or deletions in one export; commit and reopen between bounded edits.",
                        PartPath(slidePart));
                if (sourceOrderChanged && !zOrderPlan.Supported)
                    throw new CodecException(
                        "unsupported_presentation_element_reorder",
                        $"Presentation slide {slideIndex + 1} cannot safely reorder its direct elements: {zOrderPlan.BlockedReason}.",
                        PartPath(slidePart));
                var pendingElementDeletions = new List<(OpenXmlElement Source, PptxElementDeletionPlan Plan)>();
                for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
                {
                    semanticItems++;
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    var sourceElement = sourceElements[elementIndex];
                    // Element identities are owned by the source-bound slide, not
                    // by its current presentation position. A slide reorder must
                    // therefore retain the imported owner ID while proving each
                    // element binding against the original SlidePart.
                    var original = ReadElement(sourceElement, target.Id, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                    var deletionPlan = sourceDeletionPlans[elementIndex];
                    SetElementDeletionCapability(original, deletionPlan);
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Table)
                    {
                        semanticItems += checked((ulong)original.Table.Rows.Sum(row => row.Cells.Count));
                        if (semanticItems > limits.MaxCells)
                            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    }
                    if (!requestedBySourceIndex.TryGetValue(elementIndex, out var requested))
                    {
                        if (!deletionsBySourceIndex.TryGetValue(elementIndex, out var deletion))
                            throw new CodecException(
                                "presentation_element_topology_changed",
                                $"Presentation slide {slideIndex + 1} source element {elementIndex + 1} is neither retained nor explicitly deleted.",
                                PartPath(slidePart));
                        SetElementZOrderCapability(original, zOrderPlan);
                        AssertElementBinding(deletion.Id, deletion.Source, sourceElement, original, deletionPlan, zOrderPlan, slideIndex, elementIndex, slidePart);
                        if (!deletion.Id.Equals(original.Id, StringComparison.Ordinal))
                            throw new CodecException(
                                "presentation_element_deletion_binding_mismatch",
                                $"Presentation slide {slideIndex + 1} deletion {elementIndex + 1} changed its source element identity.",
                                PartPath(slidePart));
                        if (!deletionPlan.Supported)
                            throw new CodecException(
                                "unsupported_presentation_element_delete",
                                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} cannot be safely deleted: {deletionPlan.BlockedReason}.",
                                PartPath(slidePart));
                        pendingElementDeletions.Add((sourceElement, deletionPlan));
                        continue;
                    }
                    var elementBinding = requested.Source!;
                    SetElementZOrderCapability(original, zOrderPlan);
                    AssertElementBinding(requested.Id, elementBinding, sourceElement, original, deletionPlan, zOrderPlan, slideIndex, elementIndex, slidePart);
                    PptxOleWorkbookReplacement? oleWorkbookReplacement = null;
                    PptxOleOfficePackageReplacement? oleOfficePackageReplacement = null;
                    PptxDiagramTextReplacement? diagramTextReplacement = null;
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        ValidateNativeObjectRequest(original, requested);
                        oleWorkbookReplacement = PptxOleWorkbookCodec.PrepareReplacement(original.Opaque, requested.Opaque, assetCatalog, limits);
                        oleOfficePackageReplacement = PptxOleOfficePackageCodec.PrepareReplacement(original.Opaque, requested.Opaque, assetCatalog, limits);
                        diagramTextReplacement = PptxDiagramTextCodec.PrepareReplacement(slidePart, sourceElement, original.Opaque, requested.Opaque);
                    }
                    if (SemanticHash(requested).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!elementBinding.Editable)
                    {
                        if (elementBinding.TextEditable &&
                            sourceElement is P.Shape sourcePlaceholder &&
                            requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                            PptxPlaceholderCodec.SupportsSlideTextEditing(sourcePlaceholder))
                        {
                            PptxPlaceholderCodec.ApplySlideText(sourcePlaceholder, original, requested, slideContext);
                            changed = true;
                            continue;
                        }
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                    if (sourceElement is P.Shape sourceShape &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                        IsSimpleShape(sourceShape))
                    {
                        ApplyShape(sourceShape, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.Picture sourcePicture &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Image &&
                             PptxPictureCodec.TryRead(sourcePicture, slideContext, out _))
                    {
                        PptxPictureCodec.Apply(sourcePicture, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceTable &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Table &&
                             PptxTableCodec.TryRead(sourceTable, out _))
                    {
                        PptxTableCodec.Apply(sourceTable, requested);
                        changed = true;
                    }
                    else if (sourceElement is P.ConnectionShape sourceConnector &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Connector &&
                             PptxConnectorCodec.TryRead(sourceConnector, elementIdsByNativeId, out _))
                    {
                        PptxConnectorCodec.Apply(sourceConnector, requested, nativeIdsByElementId);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceChart &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Chart &&
                             PptxChartCodec.TryRead(sourceChart, slideContext, out _, out var chartEditable) && chartEditable)
                    {
                        var replacement = PptxChartCodec.Apply(sourceChart, requested, slideContext);
                        changedParts.Add(replacement.PartPath);
                        replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
                        changed = true;
                    }
                    else if (sourceElement is P.GroupShape sourceGroup &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             original.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             TryReadGroup(sourceGroup, original.Id, slideContext, elementIdsByNativeId, out _))
                    {
                        if (ApplyGroup(sourceGroup, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"element {elementIndex + 1}"))
                            changed = true;
                    }
                    else if (requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                             PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        if (oleWorkbookReplacement is not null)
                        {
                            PptxOleWorkbookCodec.Apply(slidePart, sourceElement, original.Opaque.OleWorkbook, oleWorkbookReplacement);
                            changedParts.Add(oleWorkbookReplacement.PartPath);
                            replacedOpaquePartHashes.Add(oleWorkbookReplacement.PartPath, oleWorkbookReplacement.Sha256);
                        }
                        if (oleOfficePackageReplacement is not null)
                        {
                            PptxOleOfficePackageCodec.Apply(slidePart, sourceElement, original.Opaque.OleOfficePackage, oleOfficePackageReplacement);
                            changedParts.Add(oleOfficePackageReplacement.PartPath);
                            replacedOpaquePartHashes.Add(oleOfficePackageReplacement.PartPath, oleOfficePackageReplacement.Sha256);
                        }
                        if (diagramTextReplacement is not null)
                        {
                            PptxDiagramTextCodec.Apply(slidePart, original.Opaque.DiagramText, diagramTextReplacement);
                            changedParts.Add(diagramTextReplacement.PartPath);
                            replacedOpaquePartHashes.Add(diagramTextReplacement.PartPath, diagramTextReplacement.Sha256);
                        }
                        if (NativePlacementChanged(original, requested))
                        {
                            ApplyNativePlacement(sourceElement, requested);
                            changed = true;
                        }
                    }
                    else
                    {
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                }
                if (sourceOrderChanged)
                {
                    ApplyElementZOrder(shapeTree, sourceElements, requestedSourceOrder);
                    changed = true;
                }
                foreach (var deletion in pendingElementDeletions)
                {
                    PptxElementDeletionCodec.Apply(slidePart, deletion.Source, deletion.Plan);
                    foreach (var removedRelationshipId in deletion.Plan.RelationshipIds)
                    {
                        removedSourceRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                        removedElementRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                    }
                    if (deletion.Plan.RelationshipIds.Count > 0)
                        changedParts.Add(RelationshipPartPath(slidePart));
                    if (deletion.Plan.RemovedPackagePartPaths.Count > 0)
                    {
                        changedParts.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        removedSourcePartPaths.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        removedElementPartPaths.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        changedParts.Add("[Content_Types].xml");
                    }
                    changed = true;
                }
                if (authoredElements.Length > 0)
                {
                    if (changed)
                        throw new CodecException(
                            "unsupported_presentation_authored_overlay",
                            $"Presentation slide {slideIndex + 1} cannot combine an authored overlay with another SlidePart mutation in one export; commit and reopen between bounded edits.",
                            PartPath(slidePart));
                    var relationshipCount = slideContext.AddedRelationshipIds.Count;
                    var partCount = slideContext.AddedPartPaths.Count;
                    var authoredXml = authoredElements
                        .Select(authored => BuildElement(authored, nativeIdsByElementId, slideContext, slidePart).OuterXml)
                        .ToArray();
                    var authoredImageCount = authoredElements.Count(element => element.ContentCase == PresentationElement.ContentOneofCase.Image);
                    var addedRelationshipCount = slideContext.AddedRelationshipIds.Count - relationshipCount;
                    var addedPartCount = slideContext.AddedPartPaths.Count - partCount;
                    if (addedRelationshipCount < 0 || addedPartCount < 0 ||
                        addedRelationshipCount > authoredImageCount || addedPartCount > addedRelationshipCount)
                        throw new CodecException(
                            "unsupported_presentation_authored_overlay",
                            $"Presentation slide {slideIndex + 1} authored overlay changed relationships outside its embedded-image allowance.",
                            PartPath(slidePart));
                    authoredOverlayXmlByPartPath.Add(PartPath(slidePart), authoredXml);
                    changedParts.Add(PartPath(slidePart));
                }
                if (changed)
                {
                    slideRoot.Save();
                    changedParts.Add(PartPath(slidePart));
                }
                if (PptxSpeakerNotesCodec.ApplySourceBound(presentationPart, slidePart, target.SpeakerNotes, slideIndex) is { } notesChange)
                {
                    changedParts.UnionWith(notesChange.ChangedPartPaths);
                    addedPartPaths.UnionWith(notesChange.AddedPartPaths);
                    addedRelationshipIds.UnionWith(notesChange.AddedRelationshipKeys);
                    foreach (var (partPath, sha256) in notesChange.ReplacedPartHashes)
                        replacedOpaquePartHashes.Add(partPath, sha256);
                }
                if (PptxModernCommentsCodec.ApplySourceBound(
                        presentationPart,
                        targetSlide.Source.SlideId,
                        slidePart,
                        sourceElements,
                        elementIdsByNativeId,
                        target,
                        slideIndex) is { } modernCommentsChange)
                {
                    changedParts.Add(modernCommentsChange.PartPath);
                    replacedOpaquePartHashes.Add(modernCommentsChange.PartPath, modernCommentsChange.Sha256);
                }
                TrackContextChanges(slidePart, slideContext, changedParts, addedRelationshipIds, addedPartPaths);
            }
            if (PptxLegacyCommentsCodec.ApplySourceBoundEdits(
                    presentationPart,
                    slideParts,
                    retainedTargets.Select(target => target.Target).ToArray()) is { } legacyCommentsEdit)
            {
                changedParts.UnionWith(legacyCommentsEdit.ChangedPartPaths);
                addedPartPaths.UnionWith(legacyCommentsEdit.AddedPartPaths);
                addedRelationshipIds.UnionWith(legacyCommentsEdit.AddedRelationshipKeys);
                foreach (var (partPath, sha256) in legacyCommentsEdit.ReplacedPartHashes)
                    replacedOpaquePartHashes.Add(partPath, sha256);
            }
            if (PptxLegacyCommentsCodec.ApplySourceBoundAdditions(
                    presentationPart,
                    slideParts,
                    retainedTargets.Select(target => target.Target).ToArray()) is { } legacyCommentsChange)
            {
                changedParts.UnionWith(legacyCommentsChange.ChangedPartPaths);
                addedPartPaths.UnionWith(legacyCommentsChange.AddedPartPaths);
                addedRelationshipIds.UnionWith(legacyCommentsChange.AddedRelationshipKeys);
                foreach (var (partPath, sha256) in legacyCommentsChange.ReplacedPartHashes)
                    replacedOpaquePartHashes.Add(partPath, sha256);
            }
        }

        // Opening an OPC package can rewrite ZIP container metadata even with
        // AutoSave disabled. When no modeled part, relationship, or opaque
        // graph changed, the validated source bytes are the only lossless
        // export: returning the stream would create a spurious package-level
        // diff while every retained part remains identical.
        var noModeledChanges = changedParts.Count == 0 &&
            addedRelationshipIds.Count == 0 &&
            addedPartPaths.Count == 0 &&
            replacedOpaquePartHashes.Count == 0 &&
            removedSourcePartPaths.Count == 0 &&
            removedSourceRelationshipKeys.Count == 0;
        var bytes = noModeledChanges
            ? sourceBytes
            : NormalizeChangedPartTimestamps(
                stream.ToArray(),
                sourceBytes,
                addedPartPaths.Concat(clonedPackageEntryPaths).ToHashSet(StringComparer.OrdinalIgnoreCase),
                changedParts);
        if (authoredOverlayXmlByPartPath.Count > 0)
        {
            var replacements = authoredOverlayXmlByPartPath.ToDictionary(
                entry => entry.Key,
                entry => PptxEditPlanCodec.AppendShapeTreeChildren(
                    PptxEditPlanCodec.ReadPart(bytes, entry.Key),
                    entry.Value,
                    entry.Key),
                StringComparer.OrdinalIgnoreCase);
            bytes = PptxEditPlanCodec.ReplaceParts(bytes, replacements);
        }
        ValidateOutputBudget(bytes, limits);
        AssertPlannedPartsRemoved(sourceBytes, bytes, removedSourcePartPaths);
        var retainedValidationErrorCount = ValidateOffice2021AgainstSource(sourceBytes, bytes, clonedPartSourcePaths);
        AssertPackagePartsUnchangedExcept(sourceBytes, bytes, changedParts);
        ValidatePreservedSlideElements(sourceBytes, bytes, envelope.Presentation, limits);
        ValidatePreservedMasterAndLayoutContent(sourceBytes, bytes, envelope.Presentation, limits);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        AssertOpaqueGraphMatchesWithModeledAdditions(
            envelope.OpaqueOpc,
            outputOpaque,
            addedRelationshipIds,
            addedPartPaths,
            replacedOpaquePartHashes,
            removedSourcePartPaths,
            removedSourceRelationshipKeys);
        var diagnostics = new List<Diagnostic>();
        var removedElementOpaqueCount = envelope.OpaqueOpc.Parts.Count(part => removedElementPartPaths.Contains(part.Path)) +
                                        envelope.OpaqueOpc.PackageRelationships.Count(relationship =>
                                            removedElementPartPaths.Contains(relationship.SourcePath) ||
                                            removedElementRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"));
        var removedOpaqueCount = envelope.OpaqueOpc.Parts.Count(part => removedSourcePartPaths.Contains(part.Path)) +
                                 envelope.OpaqueOpc.PackageRelationships.Count(relationship =>
                                     removedSourcePartPaths.Contains(relationship.SourcePath) ||
                                     removedSourceRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"));
        var removedSlideOpaqueCount = removedOpaqueCount - removedElementOpaqueCount;
        var retainedOpaqueCount = opaqueCount - removedOpaqueCount;
        if (retainedOpaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_preserved",
                $"Preserved {retainedOpaqueCount} opaque OPC parts or relationships while updating modeled presentation content."));
        if (removedSlideOpaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_deleted_with_slide",
                $"Removed {removedSlideOpaqueCount} opaque OPC parts or relationships because they belonged exclusively to an explicitly deleted source slide graph."));
        if (removedElementOpaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_deleted_with_element",
                $"Removed {removedElementOpaqueCount} opaque OPC parts or relationships because they belonged exclusively to explicitly deleted source slide elements."));
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; export introduced none."));
        return new PptxExportResult(bytes, diagnostics);
    }

    private static byte[] NormalizeChangedPartTimestamps(
        byte[] bytes,
        byte[] sourceBytes,
        IReadOnlyCollection<string> addedPartPaths,
        IReadOnlyCollection<string> changedPartPaths)
    {
        if (addedPartPaths.Count == 0 && changedPartPaths.Count == 0) return bytes;
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false);
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var timestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var added = addedPartPaths.ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var path in added.Order(StringComparer.OrdinalIgnoreCase))
            {
                var entry = archive.GetEntry(path) ?? throw new CodecException(
                    "presentation_added_part_missing",
                    $"PPTX declared added part {path} is missing from the output package.",
                    path);
                entry.LastWriteTime = timestamp;
            }
            foreach (var path in changedPartPaths.Where(path => !added.Contains(path)).Order(StringComparer.OrdinalIgnoreCase))
            {
                var outputEntry = archive.GetEntry(path);
                var sourceEntry = sourceArchive.GetEntry(path);
                if (outputEntry is not null && sourceEntry is not null)
                    outputEntry.LastWriteTime = sourceEntry.LastWriteTime;
            }
        }
        return stream.ToArray();
    }

    // This is deliberately a canvas-only mutation. PresentationML leaves all
    // slide, layout, and master coordinates untouched when p:sldSz changes;
    // callers that want a reflow must explicitly use their layout primitives.
    // Clearing the preset type avoids claiming that an arbitrary pair of EMU
    // dimensions still matches the old Office preset.
    private static bool ApplySourceBoundSlideSize(PresentationPart presentationPart, PresentationArtifact requested)
    {
        var presentation = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var source = presentation.SlideSize;
        var sourceWidth = source?.Cx?.Value ?? DefaultSlideWidthEmu;
        var sourceHeight = source?.Cy?.Value ?? DefaultSlideHeightEmu;
        if (requested.SlideWidthEmu == sourceWidth && requested.SlideHeightEmu == sourceHeight)
            return false;
        if (requested.SlideWidthEmu <= 0 || requested.SlideHeightEmu <= 0 ||
            requested.SlideWidthEmu > int.MaxValue || requested.SlideHeightEmu > int.MaxValue)
            throw new CodecException(
                "invalid_slide_size",
                "Source-bound PPTX canvas dimensions must be positive signed 32-bit EMU values.",
                "ppt/presentation.xml");

        if (source is null)
        {
            source = new P.SlideSize();
            var slideIdList = presentation.SlideIdList ??
                throw new CodecException("missing_slide_id_list", "PPTX presentation has no slide ID list.", "ppt/presentation.xml");
            presentation.InsertAfter(source, slideIdList);
        }
        source.Cx = checked((int)requested.SlideWidthEmu);
        source.Cy = checked((int)requested.SlideHeightEmu);
        source.Type = null;
        presentation.Save();
        return true;
    }

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        int slideIndex,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null)
        => ReadElement(source, $"presentation/slide/{slideIndex + 1}", elementIndex, slideContext, nativeObjects, elementIdsByNativeId);

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        string ownerId,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null)
    {
        var element = new PresentationElement
        {
            Id = $"{ownerId}/element/{elementIndex + 1}",
            Name = ElementName(source, elementIndex),
        };
        var nativeMediaPicture = PptxNativeObjectCatalog.IsMediaPicture(source);
        var editable = false;
        var modeled = false;
        if (source is P.Shape sourceShape)
        {
            editable = IsSimpleShape(sourceShape);
            element.Shape = ReadShape(sourceShape, slideContext);
            modeled = true;
        }
        else if (source is P.Picture sourcePicture && !nativeMediaPicture && PptxPictureCodec.TryRead(sourcePicture, slideContext, out var image))
        {
            element.Image = image;
            editable = true;
            modeled = true;
        }
        else if (source is P.GraphicFrame sourceFrame)
        {
            var tableModeled = PptxTableCodec.TryRead(sourceFrame, out var table);
            var chartModeled = PptxChartCodec.TryRead(sourceFrame, slideContext, out var chart, out var chartEditable);
            editable = chartModeled ? chartEditable : tableModeled;
            if (tableModeled)
                element.Table = table;
            else if (chartModeled)
                element.Chart = chart;
            modeled = tableModeled || chartModeled;
        }
        else if (source is P.ConnectionShape sourceConnector && PptxConnectorCodec.TryRead(sourceConnector, elementIdsByNativeId, out var connector))
        {
            element.Connector = connector;
            editable = true;
            modeled = true;
        }
        else if (source is P.GroupShape sourceGroup && TryReadGroup(sourceGroup, element.Id, slideContext, elementIdsByNativeId, out var group))
        {
            element.Group = group;
            editable = true;
            modeled = true;
        }
        if (!modeled)
        {
            var frame = ReadFrame(source);
            element.Opaque = new PresentationOpaqueElement
            {
                ElementName = source.LocalName,
                Text = DescendantText(source),
                RawXml = source.OuterXml,
                LeftEmu = frame.Left,
                TopEmu = frame.Top,
                WidthEmu = frame.Width,
                HeightEmu = frame.Height,
            };
            nativeObjects?.Populate(element.Opaque, source, slideContext.Owner);
            var nativeKind = string.IsNullOrEmpty(element.Opaque.NativeKind)
                ? PptxNativeObjectCatalog.Classify(source)
                : element.Opaque.NativeKind;
            editable = PptxNativeObjectCatalog.SupportsPlacementEditing(source, nativeKind);
        }
        element.Source = new PresentationElementSourceBinding
        {
            ShapeTreeIndex = checked((uint)elementIndex),
            ElementSha256 = HashElement(source),
            Editable = editable,
            TextEditable = source is P.Shape textShape && textShape.TextBody is not null && PptxTextCodec.SupportsEditing(textShape.TextBody),
            AccessibilityEditable = editable && (
                (source is P.Picture accessibilityPicture && element.ContentCase == PresentationElement.ContentOneofCase.Image &&
                 PptxNonVisualAccessibilityCodec.SupportsResidual(accessibilityPicture.NonVisualPictureProperties?.NonVisualDrawingProperties)) ||
                PptxNonVisualAccessibilityCodec.Supports(source switch
                {
                    P.Shape accessibilityShape => accessibilityShape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                    P.ConnectionShape accessibilityConnector => accessibilityConnector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
                    P.GroupShape accessibilityGroup => accessibilityGroup.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
                    P.GraphicFrame accessibilityFrame when element.ContentCase is PresentationElement.ContentOneofCase.Table or PresentationElement.ContentOneofCase.Chart =>
                        accessibilityFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
                    _ => null,
                })),
        };
        element.Source.SemanticSha256 = SemanticHash(element);
        return element;
    }

    private static void SetElementDeletionCapability(
        PresentationElement element,
        PptxElementDeletionPlan plan)
    {
        element.Source.DeletionCapability = new PresentationElementDeletionCapability
        {
            Supported = plan.Supported,
            BlockedReason = plan.BlockedReason,
            NativeId = plan.NativeId,
        };
    }

    private static PresentationElementZOrderCapability AnalyzeElementZOrder(IReadOnlyList<OpenXmlElement> elements)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            if (elements[index] is P.Shape or P.Picture or P.GraphicFrame or P.ConnectionShape or P.GroupShape or P.ContentPart) continue;
            return new PresentationElementZOrderCapability
            {
                Supported = false,
                BlockedReason = $"direct shape-tree child {index + 1} ({elements[index].LocalName}) is not a movable drawing node",
            };
        }
        return new PresentationElementZOrderCapability { Supported = true };
    }

    private static void SetElementZOrderCapability(
        PresentationElement element,
        PresentationElementZOrderCapability plan)
    {
        element.Source.ZOrderCapability = plan.Clone();
    }

    private static void ApplyElementZOrder(
        P.ShapeTree shapeTree,
        IReadOnlyList<OpenXmlElement> sourceElements,
        IReadOnlyList<int> requestedSourceOrder)
    {
        if (sourceElements.Count != requestedSourceOrder.Count || requestedSourceOrder.Distinct().Count() != sourceElements.Count ||
            requestedSourceOrder.Any(index => index < 0 || index >= sourceElements.Count))
            throw new CodecException("presentation_element_topology_changed", "Presentation direct-element reorder must be a complete permutation of the source shape tree.");
        OpenXmlElement anchor = shapeTree.GetFirstChild<P.GroupShapeProperties>() ??
            throw new CodecException("missing_shape_tree", "Presentation shape tree has no group-shape properties anchor.");
        foreach (var sourceElement in sourceElements) sourceElement.Remove();
        foreach (var sourceIndex in requestedSourceOrder)
        {
            var sourceElement = sourceElements[sourceIndex];
            shapeTree.InsertAfter(sourceElement, anchor);
            anchor = sourceElement;
        }
    }

    private static void AssertElementBinding(
        string requestedId,
        PresentationElementSourceBinding binding,
        OpenXmlElement sourceElement,
        PresentationElement original,
        PptxElementDeletionPlan deletionPlan,
        PresentationElementZOrderCapability zOrderPlan,
        int slideIndex,
        int elementIndex,
        SlidePart slidePart)
    {
        if (!requestedId.Equals(original.Id, StringComparison.Ordinal) ||
            binding.ShapeTreeIndex != elementIndex ||
            !binding.ElementSha256.Equals(HashElement(sourceElement), StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_element_binding_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} does not match its source element.",
                PartPath(slidePart));
        if (binding.Editable != original.Source.Editable ||
            binding.DirectFramePresenceEditable != original.Source.DirectFramePresenceEditable ||
            binding.TextEditable != original.Source.TextEditable ||
            binding.AccessibilityEditable != original.Source.AccessibilityEditable ||
            binding.DeletionCapability is null ||
            binding.DeletionCapability.Supported != deletionPlan.Supported ||
            binding.DeletionCapability.NativeId != deletionPlan.NativeId ||
            !binding.DeletionCapability.BlockedReason.Equals(deletionPlan.BlockedReason, StringComparison.Ordinal) ||
            binding.ZOrderCapability is null ||
            binding.ZOrderCapability.Supported != zOrderPlan.Supported ||
            !binding.ZOrderCapability.BlockedReason.Equals(zOrderPlan.BlockedReason, StringComparison.Ordinal))
            throw new CodecException(
                "presentation_element_binding_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} changed its source capability contract.",
                PartPath(slidePart));
        if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_source_semantics_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} source semantics do not match its binding.",
                PartPath(slidePart));
    }

    private static bool TryReadGroup(
        P.GroupShape source,
        string groupId,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId,
        out PresentationGroup group)
    {
        group = new PresentationGroup();
        var nonVisual = source.GetFirstChild<P.NonVisualGroupShapeProperties>();
        var properties = source.GetFirstChild<P.GroupShapeProperties>();
        var transform = properties?.GetFirstChild<A.TransformGroup>();
        if (nonVisual is null || properties is null || transform is null ||
            source.Elements<P.NonVisualGroupShapeProperties>().Count() != 1 ||
            source.Elements<P.GroupShapeProperties>().Count() != 1 ||
            nonVisual.ChildElements.Count != 3 ||
            nonVisual.ChildElements[0] is not P.NonVisualDrawingProperties drawing ||
            nonVisual.ChildElements[1] is not P.NonVisualGroupShapeDrawingProperties groupDrawing ||
            nonVisual.ChildElements[2] is not P.ApplicationNonVisualDrawingProperties application ||
            groupDrawing.ChildElements.Count != 0 || application.ChildElements.Count != 0 ||
            drawing.Id?.Value is null or 0 || drawing.Name?.Value is not { Length: <= 1_024 } ||
            !HasOnlyAttributes(groupDrawing) || !HasOnlyAttributes(application) ||
            properties.ChildElements.Count != 1 || properties.FirstChild != transform || !HasOnlyAttributes(properties) ||
            !HasOnlyAttributes(transform) || transform.ChildElements.Count != 4 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            transform.ChildElements[2] is not A.ChildOffset childOffset ||
            transform.ChildElements[3] is not A.ChildExtents childExtents ||
            !HasOnlyAttributes(offset, "x", "y") || !HasOnlyAttributes(extents, "cx", "cy") ||
            !HasOnlyAttributes(childOffset, "x", "y") || !HasOnlyAttributes(childExtents, "cx", "cy") ||
            extents.Cx?.Value <= 0 || extents.Cy?.Value <= 0 || childExtents.Cx?.Value <= 0 || childExtents.Cy?.Value <= 0 ||
            offset.X?.Value < 0 || offset.Y?.Value < 0)
            return false;

        group.LeftEmu = offset.X?.Value ?? 0;
        group.TopEmu = offset.Y?.Value ?? 0;
        group.WidthEmu = extents.Cx?.Value ?? 0;
        group.HeightEmu = extents.Cy?.Value ?? 0;
        group.ChildLeftEmu = childOffset.X?.Value ?? 0;
        group.ChildTopEmu = childOffset.Y?.Value ?? 0;
        group.ChildWidthEmu = childExtents.Cx?.Value ?? 0;
        group.ChildHeightEmu = childExtents.Cy?.Value ?? 0;
        group.Accessibility = PptxNonVisualAccessibilityCodec.Read(drawing);
        var children = GroupElements(source);
        if (children.Length == 0) return false;
        for (var index = 0; index < children.Length; index++)
        {
            var child = ReadElement(children[index], groupId, index, slideContext, elementIdsByNativeId: elementIdsByNativeId);
            if (child.ContentCase is PresentationElement.ContentOneofCase.Opaque or PresentationElement.ContentOneofCase.None || child.Source?.Editable != true)
                return false;
            group.Children.Add(child);
        }
        return true;
    }

    private static PresentationShape ReadShape(P.Shape shape, PptxPartContext slideContext)
    {
        var frame = ReadFrame(shape);
        var properties = shape.ShapeProperties;
        var textBody = PptxTextCodec.Read(shape.TextBody, slideContext);
        var placeholder = PptxPlaceholderCodec.ReadIdentity(shape);
        var transform = properties?.Transform2D;
        var geometry = Geometry(shape);
        var solidFill = properties?.GetFirstChild<A.SolidFill>();
        var result = new PresentationShape
        {
            Geometry = geometry,
            LeftEmu = frame.Left,
            TopEmu = frame.Top,
            WidthEmu = frame.Width,
            HeightEmu = frame.Height,
            Text = PptxTextCodec.Flatten(textBody),
            TextBody = textBody,
            FillRgb = PptxColor.SolidRgb(solidFill),
            Placeholder = placeholder,
            DirectFrame = placeholder is null ? null : PptxPlaceholderCodec.ReadDirectFrame(shape),
            Transform = placeholder is null && PptxShapeTransformCodec.Supports(transform, allowSingleZeroExtent: geometry == "line")
                ? PptxShapeTransformCodec.Read(transform!)
                : null,
            Shadow = ReadShadow(properties),
        };
        if (ReadFillOpacity(solidFill) is { } fillOpacity)
            result.FillOpacityThousandthPercent = fillOpacity;
        PptxLineStyleCodec.ReadForProjection(properties?.GetFirstChild<A.Outline>(), result);
        if (!string.Equals(geometry, "line", StringComparison.Ordinal))
        {
            // Non-line arrowheads remain source-bound. Keep the historic
            // ordinary-shape preview projection valid without implying that
            // those endpoint semantics are editable.
            result.StartArrow = result.EndArrow = result.StartArrowWidth = result.StartArrowLength =
                result.EndArrowWidth = result.EndArrowLength = string.Empty;
        }
        if (shape.UseBackgroundFill?.HasValue == true)
            result.UseBackgroundFill = shape.UseBackgroundFill.Value;
        PptxCustomGeometryCodec.Read(properties?.GetFirstChild<A.CustomGeometry>(), frame.Width, frame.Height, result);
        result.Accessibility = PptxNonVisualAccessibilityCodec.Read(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);
        return result;
    }

    private static bool IsSimpleShape(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is not null) return false;
        if (shape.ShapeStyle is not null) return false;
        var properties = shape.ShapeProperties;
        var transform = properties?.Transform2D;
        var geometry = Geometry(shape);
        if (properties is null || properties.Elements<A.Transform2D>().Count() != 1 ||
            !PptxShapeTransformCodec.Supports(transform, allowSingleZeroExtent: geometry == "line")) return false;
        if (geometry is not ("rect" or "ellipse" or "roundRect" or "textbox" or "line" or "custom")) return false;
        if (geometry == "custom")
        {
            var frame = ReadFrame(shape);
            if (!PptxCustomGeometryCodec.Supports(properties.GetFirstChild<A.CustomGeometry>(), frame.Width, frame.Height)) return false;
        }
        if (!SimpleFill(properties)) return false;
        var outline = properties.GetFirstChild<A.Outline>();
        if (!PptxLineStyleCodec.TryRead(outline, out var lineStyle)) return false;
        if (!string.Equals(geometry, "line", StringComparison.Ordinal) &&
            (lineStyle.StartArrow.Length > 0 || lineStyle.EndArrow.Length > 0)) return false;
        if (!SupportsShadow(properties)) return false;
        if (properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.CustomGeometry and not A.NoFill and not A.SolidFill and not A.Outline and not A.EffectList)) return false;
        return PptxTextCodec.SupportsEditing(shape.TextBody);
    }

    private static bool SimpleFill(OpenXmlCompositeElement element)
    {
        var fills = element.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
        if (fills.Length > 1) return false;
        if (fills.Length == 0 || fills[0] is A.NoFill) return true;
        var solid = (A.SolidFill)fills[0];
        if (solid.ChildElements.Count != 1 || solid.FirstChild is not A.RgbColorModelHex color || !HasOnlyAttributes(color, "val")) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == alphas.Length && alphas.Length <= 1 &&
               (alphas.Length == 0 || alphas[0].Val?.Value is >= 0 and <= 100_000 && HasOnlyAttributes(alphas[0], "val"));
    }

    private static uint? ReadFillOpacity(A.SolidFill? solid)
    {
        var alpha = solid?.GetFirstChild<A.RgbColorModelHex>()?.GetFirstChild<A.Alpha>();
        return alpha?.Val?.Value is { } value ? checked((uint)value) : null;
    }

    private static PresentationShadow? ReadShadow(P.ShapeProperties? properties)
    {
        if (!SupportsShadow(properties)) return null;
        var outer = properties?.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>();
        if (outer?.GetFirstChild<A.RgbColorModelHex>() is not { } color) return null;
        return new PresentationShadow
        {
            ColorRgb = PptxColor.Normalize(color.Val?.Value ?? string.Empty),
            BlurRadiusEmu = outer.BlurRadius?.Value ?? 0L,
            DistanceEmu = outer.Distance?.Value ?? 0L,
            DirectionAngle60000 = outer.Direction?.Value ?? 0,
            OpacityThousandthPercent = checked((uint)(color.GetFirstChild<A.Alpha>()?.Val?.Value ?? 100_000)),
        };
    }

    private static bool SupportsShadow(P.ShapeProperties? properties)
    {
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            !HasOnlyAttributes(outer, "blurRad", "dist", "dir") || outer.BlurRadius?.Value is < 0 || outer.Distance?.Value is < 0 ||
            outer.Direction?.Value is < 0 or >= 21_600_000 || outer.ChildElements.Count != 1 || outer.FirstChild is not A.RgbColorModelHex color ||
            color.Val?.Value is not { Length: 6 } rgb || !rgb.All(Uri.IsHexDigit) || !HasOnlyAttributes(color, "val")) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == alphas.Length && alphas.Length <= 1 &&
               (alphas.Length == 0 || alphas[0].Val?.Value is >= 0 and <= 100_000 && HasOnlyAttributes(alphas[0], "val"));
    }

    private static void ApplyShadow(P.ShapeProperties properties, PresentationShadow? shadow)
    {
        properties.GetFirstChild<A.EffectList>()?.Remove();
        if (shadow is null) return;
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(shadow.ColorRgb) };
        color.Append(new A.Alpha { Val = checked((int)shadow.OpacityThousandthPercent) });
        var outer = new A.OuterShadow(color)
        {
            BlurRadius = shadow.BlurRadiusEmu,
            Distance = shadow.DistanceEmu,
            Direction = shadow.DirectionAngle60000,
        };
        properties.Append(new A.EffectList(outer));
    }

    private static void ValidateShadow(PresentationShadow? shadow, string elementId)
    {
        if (shadow is null) return;
        PptxColor.Normalize(shadow.ColorRgb);
        if (shadow.BlurRadiusEmu < 0 || shadow.DistanceEmu < 0 || shadow.DirectionAngle60000 is < 0 or >= 21_600_000 || shadow.OpacityThousandthPercent > 100_000)
            throw new CodecException("invalid_presentation_shadow", $"Presentation shape {elementId} has invalid shadow geometry or opacity.");
    }

    private static void ApplyShape(P.Shape shape, PresentationElement source, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        var sourceHasBackgroundFill = shape.UseBackgroundFill?.HasValue == true;
        if (sourceHasBackgroundFill != semantic.HasUseBackgroundFill ||
            sourceHasBackgroundFill && shape.UseBackgroundFill!.Value != semantic.UseBackgroundFill)
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation shape {source.Id} cannot change its source-bound useBgFill attribute.");
        var properties = shape.ShapeProperties ??= new P.ShapeProperties();
        var transform = properties.Transform2D ??= new A.Transform2D();
        var offset = transform.Offset ??= new A.Offset();
        offset.X = semantic.LeftEmu;
        offset.Y = semantic.TopEmu;
        var extents = transform.Extents ??= new A.Extents();
        extents.Cx = semantic.WidthEmu;
        extents.Cy = semantic.HeightEmu;
        PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        PptxCustomGeometryCodec.Apply(properties, semantic, source.Id);
        PptxNonVisualAccessibilityCodec.ApplyBound(shape.NonVisualShapeProperties?.NonVisualDrawingProperties, semantic.Accessibility);
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties)
            drawingProperties.TextBox = semantic.Geometry == "textbox" ? true : null;
        var fillOpacity = semantic.HasFillOpacityThousandthPercent ? semantic.FillOpacityThousandthPercent : (uint?)null;
        if (!FillMatches(properties, semantic.FillRgb, fillOpacity)) ReplaceFill(properties, semantic.FillRgb, fillOpacity);
        PptxLineStyleCodec.Apply(properties, semantic);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
            nonVisual.Name = source.Name;
        ApplyShadow(properties, semantic.Shadow);
        PptxTextCodec.Apply(shape, semantic, slideContext);
    }

    private static CodecException UnsupportedPresentationEdit(int slideIndex, int elementIndex, OpenXmlPart slidePart) => new(
        "unsupported_presentation_edit",
        $"Presentation slide {slideIndex + 1} element {elementIndex + 1} is preserved but not safely editable by this codec slice.",
        PartPath(slidePart));

    private static void ValidateNativeObjectRequest(PresentationElement original, PresentationElement requested)
    {
        var allowed = original.Clone();
        allowed.Name = requested.Name;
        allowed.Opaque.LeftEmu = requested.Opaque.LeftEmu;
        allowed.Opaque.TopEmu = requested.Opaque.TopEmu;
        allowed.Opaque.WidthEmu = requested.Opaque.WidthEmu;
        allowed.Opaque.HeightEmu = requested.Opaque.HeightEmu;
        if (allowed.Opaque.OleWorkbook is not null && requested.Opaque.OleWorkbook is not null)
            allowed.Opaque.OleWorkbook.ReplacementAssetId = requested.Opaque.OleWorkbook.ReplacementAssetId;
        if (allowed.Opaque.OleOfficePackage is not null && requested.Opaque.OleOfficePackage is not null)
            allowed.Opaque.OleOfficePackage.ReplacementAssetId = requested.Opaque.OleOfficePackage.ReplacementAssetId;
        if (allowed.Opaque.DiagramText is not null && requested.Opaque.DiagramText is not null)
        {
            allowed.Opaque.DiagramText.Nodes.Clear();
            allowed.Opaque.DiagramText.Nodes.Add(requested.Opaque.DiagramText.Nodes);
        }
        // Source binding equality is checked against the actual source above;
        // reuse the caller's equivalent instance to keep protobuf equality
        // focused on the semantic payload.
        allowed.Source = requested.Source.Clone();
        if (!allowed.Equals(requested))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation native object {requested.Id} may edit only its name, outer frame, and an explicitly recognized OLE payload.");
    }

    private static bool NativePlacementChanged(PresentationElement original, PresentationElement requested) =>
        original.Name != requested.Name ||
        original.Opaque.LeftEmu != requested.Opaque.LeftEmu ||
        original.Opaque.TopEmu != requested.Opaque.TopEmu ||
        original.Opaque.WidthEmu != requested.Opaque.WidthEmu ||
        original.Opaque.HeightEmu != requested.Opaque.HeightEmu;

    private static void ApplyNativePlacement(OpenXmlElement source, PresentationElement requested)
    {
        var frame = requested.Opaque;
        if (source is P.Picture picture)
        {
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(picture.ShapeProperties?.GetFirstChild<A.Transform2D>() ??
                throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported picture placement owner."), frame);
            return;
        }
        if (source is P.ConnectionShape connector)
        {
            if (connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(connector.ShapeProperties?.GetFirstChild<A.Transform2D>() ??
                throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported connector placement owner."), frame);
            return;
        }
        if (source is P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(graphicFrame.Transform!, frame);
            if (PptxNativeObjectCatalog.Classify(source) == "oleObject")
            {
                // PowerPoint stores a second transform on the OLE preview
                // picture. Keep it derived from the outer frame so Office and
                // fallback renderers agree after a move/resize.
                var previewTransform = graphicFrame.Descendants<A.Transform2D>().FirstOrDefault();
                if (previewTransform is not null) SetFrame(previewTransform, frame);
            }
            return;
        }
        if (source is P.GroupShape group)
        {
            if (group.GetFirstChild<P.NonVisualGroupShapeProperties>()?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(group.GetFirstChild<P.GroupShapeProperties>()!.GetFirstChild<A.TransformGroup>()!, frame);
            return;
        }
        throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported placement owner.");
    }

    private static void SetFrame(P.Transform transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.Transform2D transform, PresentationOpaqueElement frame)
    {
        transform.Offset ??= new A.Offset();
        transform.Extents ??= new A.Extents();
        transform.Offset.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.TransformGroup transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static OpenXmlElement BuildFill(string rgb, uint? opacity)
    {
        if (string.IsNullOrWhiteSpace(rgb)) return new A.NoFill();
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(rgb) };
        if (opacity.HasValue) color.Append(new A.Alpha { Val = checked((int)opacity.Value) });
        return new A.SolidFill(color);
    }

    private static void ReplaceFill(OpenXmlCompositeElement parent, string rgb, uint? opacity)
    {
        foreach (var child in parent.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) child.Remove();
        var fill = BuildFill(rgb, opacity);
        var reference = parent.ChildElements.FirstOrDefault(child => child is A.Outline || child.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d");
        if (reference is null) parent.Append(fill);
        else parent.InsertBefore(fill, reference);
    }

    private static bool FillMatches(OpenXmlCompositeElement parent, string rgb, uint? opacity)
    {
        var requested = string.IsNullOrWhiteSpace(rgb) ? string.Empty : PptxColor.Normalize(rgb);
        if (parent.GetFirstChild<A.NoFill>() is not null) return requested.Length == 0 && opacity is null;
        var solidFill = parent.GetFirstChild<A.SolidFill>();
        var solid = PptxColor.SolidRgb(solidFill);
        if (solid.Length > 0)
            return requested.Equals(solid, StringComparison.OrdinalIgnoreCase) && ReadFillOpacity(solidFill) == opacity;
        return requested.Length == 0 && opacity is null && !parent.ChildElements.Any(child => child.LocalName.EndsWith("Fill", StringComparison.Ordinal));
    }

    private static void BuildPresentation(PresentationDocument package, PresentationArtifact artifact, PptxAssetCatalog assetCatalog)
    {
        if (artifact.Masters.Count > 1)
            throw new CodecException(
                "unsupported_presentation_features",
                "New PPTX authoring supports one canonical master. Multiple slide-master graphs require a validated source package.");
        var authoredMaster = artifact.Masters.FirstOrDefault();
        var canonicalMasterId = authoredMaster?.Id ?? "__officekit/default-master";
        var presentationPart = package.AddPresentationPart();
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster1");
        var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme1");

        themePart.Theme = BasicTheme();
        var sourceLayouts = artifact.Layouts.ToList();
        PresentationLayout? fallbackLayout = null;
        if (sourceLayouts.Count == 0 || artifact.Slides.Any(slide => string.IsNullOrWhiteSpace(slide.LayoutId)))
        {
            var fallbackId = "__officekit/default-layout";
            while (sourceLayouts.Any(layout => layout.Id.Equals(fallbackId, StringComparison.Ordinal))) fallbackId += "-";
            fallbackLayout = new PresentationLayout
            {
                Id = fallbackId,
                Name = "Blank",
                MasterId = canonicalMasterId,
                Type = "blank",
            };
            sourceLayouts.Insert(0, fallbackLayout);
        }

        var layoutIdList = new P.SlideLayoutIdList();
        var layoutEntries = new List<(PresentationLayout Layout, SlideLayoutPart Part)>();
        for (var layoutIndex = 0; layoutIndex < sourceLayouts.Count; layoutIndex++)
        {
            var sourceLayout = sourceLayouts[layoutIndex];
            var relationshipId = $"rIdLayout{layoutIndex + 1}";
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.AddPart(masterPart, "rIdMaster1");
            var layoutRoot = new P.SlideLayout(
                new P.CommonSlideData(BasicShapeTree()) { Name = string.IsNullOrWhiteSpace(sourceLayout.Name) ? "Blank" : sourceLayout.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()))
            { Preserve = true };
            layoutRoot.SetAttribute(new OpenXmlAttribute("type", string.Empty, SourceFreeLayoutType(sourceLayout)));
            layoutPart.SlideLayout = layoutRoot;
            layoutIdList.Append(new P.SlideLayoutId
            {
                Id = checked(2_147_483_649U + (uint)layoutIndex),
                RelationshipId = relationshipId,
            });
            layoutEntries.Add((sourceLayout, layoutPart));
        }
        var layoutPartById = layoutEntries.ToDictionary(entry => entry.Layout.Id, entry => entry.Part, StringComparer.Ordinal);
        var defaultLayoutPart = fallbackLayout is not null
            ? layoutPartById[fallbackLayout.Id]
            : layoutEntries[0].Part;
        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(BasicShapeTree()) { Name = string.IsNullOrWhiteSpace(authoredMaster?.Name) ? "Office Clean Room" : authoredMaster.Name },
            BasicColorMap(),
            layoutIdList,
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var slideIdList = new P.SlideIdList();
        var slideParts = new SlidePart[artifact.Slides.Count];
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var relationshipId = $"rIdSlide{slideIndex + 1}";
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            slideParts[slideIndex] = slidePart;
            var layoutPart = string.IsNullOrWhiteSpace(source.LayoutId)
                ? defaultLayoutPart
                : layoutPartById.GetValueOrDefault(source.LayoutId) ??
                  throw new CodecException("invalid_presentation_layout", $"Presentation slide {source.Id} references missing layout {source.LayoutId}.");
            slidePart.AddPart(layoutPart, "rIdLayout1");
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(BasicShapeTree()) { Name = source.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()));
            PptxSlideVisibilityCodec.BuildSourceFree(slidePart.Slide, source);
            slideIdList.Append(new P.SlideId { Id = checked((uint)(256 + slideIndex)), RelationshipId = relationshipId });
        }
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var slidePartById = slideParts
            .Select((part, index) => (Part: part, Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
        var customShowCatalog = PptxCustomShowCatalog.From(artifact.CustomShows);
        var masterContext = new PptxPartContext(masterPart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
        PptxBackgroundCodec.Build(masterPart.SlideMaster.CommonSlideData!, authoredMaster?.Background);
        PptxMasterTextStylesCodec.Build(masterPart.SlideMaster, authoredMaster?.TextStyles, masterContext);
        var masterShapeTree = masterPart.SlideMaster.CommonSlideData!.ShapeTree!;
        foreach (var (placeholder, index) in (authoredMaster?.Placeholders ?? []).Select((placeholder, index) => (placeholder, index)))
            masterShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), masterContext));
        foreach (var (layout, layoutPart) in layoutEntries)
        {
            var layoutContext = new PptxPartContext(layoutPart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
            var layoutCommon = layoutPart.SlideLayout!.CommonSlideData!;
            PptxBackgroundCodec.Build(layoutCommon, layout.Background);
            var layoutShapeTree = layoutCommon.ShapeTree!;
            foreach (var (placeholder, index) in layout.Placeholders.Select((placeholder, index) => (placeholder, index)))
                layoutShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), layoutContext));
        }
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var slidePart = slideParts[slideIndex];
            var slideCommon = slidePart.Slide!.CommonSlideData!;
            PptxBackgroundCodec.Build(slideCommon, source.Background);
            PptxTransitionCodec.Build(slidePart.Slide!, source.Transition);
            var shapeTree = slideCommon.ShapeTree!;
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
            var flattenedElements = FlattenPresentationElements(source.Elements).ToArray();
            var nativeIdsByElementId = flattenedElements.Select((element, index) => (element.Id, NativeId: checked((uint)(index + 2))))
                .ToDictionary(item => item.Id, item => item.NativeId, StringComparer.Ordinal);
            foreach (var element in source.Elements)
                shapeTree.Append(BuildElement(element, nativeIdsByElementId, slideContext, slidePart));
            PptxTimingCodec.ValidateMorphContext(source, slideIndex > 0 ? artifact.Slides[slideIndex - 1] : null);
            PptxTimingCodec.Build(slidePart.Slide!, source, nativeIdsByElementId);
            slidePart.Slide.Save();
        }
        var notesMasterRelationshipId = PptxSpeakerNotesCodec.BuildSourceFree(presentationPart, themePart, slideParts, artifact.Slides);
        var presentationRoot = new P.Presentation();
        presentationRoot.Append(new P.SlideMasterIdList(new P.SlideMasterId { Id = 2_147_483_648U, RelationshipId = "rIdMaster1" }));
        if (notesMasterRelationshipId is not null)
            presentationRoot.Append(new P.NotesMasterIdList(new P.NotesMasterId { Id = notesMasterRelationshipId }));
        presentationRoot.Append(
            slideIdList,
            new P.SlideSize
            {
                Cx = checked((int)(artifact.SlideWidthEmu > 0 ? artifact.SlideWidthEmu : DefaultSlideWidthEmu)),
                Cy = checked((int)(artifact.SlideHeightEmu > 0 ? artifact.SlideHeightEmu : DefaultSlideHeightEmu)),
            },
            new P.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L },
            new P.DefaultTextStyle());
        if (artifact.CustomShows.Count > 0)
        {
            var relationshipIdBySlideId = artifact.Slides
                .Select((slide, index) => (slide.Id, RelationshipId: $"rIdSlide{index + 1}"))
                .ToDictionary(item => item.Id, item => item.RelationshipId, StringComparer.Ordinal);
            PptxCustomShowCodec.BuildSourceFree(presentationRoot, artifact, relationshipIdBySlideId);
        }
        if (artifact.Sections.Count > 0)
        {
            var nativeSlideIdByPublicId = artifact.Slides
                .Select((slide, index) => (slide.Id, NativeId: checked((uint)(256 + index))))
                .ToDictionary(item => item.Id, item => item.NativeId, StringComparer.Ordinal);
            PptxSectionCodec.BuildSourceFree(presentationRoot, artifact, nativeSlideIdByPublicId);
        }
        presentationPart.Presentation = presentationRoot;
        PptxLegacyCommentsCodec.BuildSourceFree(presentationPart, slideParts, artifact.Slides);
        PptxModernCommentsCodec.BuildSourceFree(presentationPart, slideIdList.Elements<P.SlideId>().ToArray(), slideParts, artifact.Slides);
        themePart.Theme.Save();
        foreach (var (_, layoutPart) in layoutEntries) layoutPart.SlideLayout!.Save();
        masterPart.SlideMaster.Save();
        presentationPart.Presentation.Save();
    }

    private static string SourceFreeLayoutType(PresentationLayout layout)
    {
        var type = string.IsNullOrWhiteSpace(layout.Type) ? "blank" : layout.Type;
        return type switch
        {
            "blank" or "title" or "titleOnly" or "obj" => type,
            _ => throw new CodecException(
                "unsupported_presentation_features",
                $"Source-free presentation layout {layout.Id} uses unsupported type {type}. Use blank, title, titleOnly, or obj."),
        };
    }

    private static void ValidateSourceFreeTextPlaceholder(PresentationPlaceholder placeholder, string ownerId)
    {
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, placeholder.Id);
        if (placeholder.DirectFrame is null)
            throw new CodecException("invalid_presentation_placeholder", $"Source-free {ownerId} placeholder {placeholder.Id} requires a direct frame.");
    }

    private static void ValidateSourceFreeTextPlaceholderIdentity(PresentationPlaceholderIdentity placeholder, string elementId) =>
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, elementId);

    private static void ValidateSourceFreeTextPlaceholderType(string type, string ownerId)
    {
        if (type is "title" or "body" or "ctrTitle" or "subTitle") return;
        throw new CodecException(
            "unsupported_presentation_features",
            $"Source-free presentation placeholder {ownerId} uses {type}; only title, body, ctrTitle, and subTitle text placeholders are supported.");
    }

    private static IEnumerable<PresentationElement> FlattenPresentationElements(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in FlattenPresentationElements(element.Group.Children)) yield return child;
        }
    }

    private static (PresentationElement[] Retained, PresentationElement[] Authored) SplitSourceBoundElements(
        PresentationSlide slide,
        int sourceElementCount,
        int slideIndex,
        OpenXmlPart owner)
    {
        var retainedCount = 0;
        while (retainedCount < slide.Elements.Count && slide.Elements[retainedCount].Source is not null) retainedCount++;
        var retained = slide.Elements.Take(retainedCount).ToArray();
        var authored = slide.Elements.Skip(retainedCount).ToArray();
        if (authored.Any(element => element.Source is not null))
            throw new CodecException(
                "presentation_element_topology_changed",
                $"Presentation slide {slideIndex + 1} source-bound elements must remain an ordered prefix before authored overlay elements.",
                PartPath(owner));
        if (retained.Length + slide.ElementDeletions.Count != sourceElementCount)
            throw new CodecException(
                "presentation_element_topology_changed",
                $"Source-preserving PPTX export requires slide {slideIndex + 1}'s original {sourceElementCount}-element topology to be covered by retained elements plus explicit deletions; the artifact contains {retained.Length} retained elements, {slide.ElementDeletions.Count} deletions, and {authored.Length} authored overlay elements.",
                PartPath(owner));
        foreach (var element in authored) ValidateAuthoredOverlayElement(element, slideIndex, owner);
        return (retained, authored);
    }

    private static void ValidateAuthoredOverlayElement(PresentationElement element, int slideIndex, OpenXmlPart owner)
    {
        if (element.ContentCase == PresentationElement.ContentOneofCase.Image)
            return;
        if (element.ContentCase != PresentationElement.ContentOneofCase.Shape)
            throw new CodecException(
                "unsupported_presentation_authored_overlay",
                $"Presentation slide {slideIndex + 1} authored overlay {element.Id} must be a canonical textbox, basic shape, or embedded rectangular image.",
                PartPath(owner));
        var shape = element.Shape;
        if (shape.Geometry is not ("textbox" or "rect" or "roundRect" or "ellipse") ||
            shape.Placeholder is not null || shape.DirectFrame is not null || shape.HasUseBackgroundFill ||
            shape.CustomPaths.Count > 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
            shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 || shape.TextRectangle is not null)
            throw new CodecException(
                "unsupported_presentation_authored_overlay",
                $"Presentation slide {slideIndex + 1} authored overlay {element.Id} uses geometry or layout identity outside the bounded textbox/basic-shape profile.",
                PartPath(owner));
        var paragraphs = (shape.TextBody?.Paragraphs ?? []).Concat(shape.TextBody?.ListStyles ?? []);
        if (paragraphs.Any(paragraph =>
                paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet ||
                paragraph.Runs.Any(run => run.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.RunHyperlink)))
            throw new CodecException(
                "unsupported_presentation_authored_overlay",
                $"Presentation slide {slideIndex + 1} authored overlay {element.Id} cannot add picture or hyperlink relationships.",
                PartPath(owner));
    }

    private static IReadOnlyDictionary<string, uint> AuthoredOverlayNativeIds(
        IReadOnlyList<OpenXmlElement> sourceElements,
        IReadOnlyList<PresentationElement> authored,
        int slideIndex,
        OpenXmlPart owner)
    {
        var occupied = sourceElements.SelectMany(PptxElementDeletionCodec.NativeIds).ToHashSet();
        var next = occupied.Count == 0 ? 1U : occupied.Max();
        var output = new Dictionary<string, uint>(StringComparer.Ordinal);
        foreach (var element in authored)
        {
            if (next == uint.MaxValue)
                throw new CodecException(
                    "unsupported_presentation_authored_overlay",
                    $"Presentation slide {slideIndex + 1} has no remaining native drawing ID for authored overlay {element.Id}.",
                    PartPath(owner));
            next++;
            if (!output.TryAdd(element.Id, next))
                throw new CodecException(
                    "invalid_presentation_element",
                    $"Presentation slide {slideIndex + 1} contains duplicate authored overlay identity {element.Id}.",
                    PartPath(owner));
        }
        return output;
    }

    private static OpenXmlElement BuildElement(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart) => element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => BuildShape(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Image => PptxPictureCodec.Build(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Table => PptxTableCodec.Build(element, nativeIdsByElementId[element.Id]),
            PresentationElement.ContentOneofCase.Connector => PptxConnectorCodec.Build(element, nativeIdsByElementId[element.Id], nativeIdsByElementId),
            PresentationElement.ContentOneofCase.Chart => PptxChartCodec.Build(element, nativeIdsByElementId[element.Id], slidePart),
            PresentationElement.ContentOneofCase.Group => BuildGroup(element, nativeIdsByElementId, slideContext, slidePart),
            _ => throw new CodecException("unsupported_presentation_element", $"Opaque presentation element {element.Id} requires its validated source package and cannot be authored from scratch."),
        };

    private static P.GroupShape BuildGroup(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart)
    {
        var group = element.Group;
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeIdsByElementId[element.Id], Name = element.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, group.Accessibility);
        var output = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                nonVisual,
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = group.LeftEmu, Y = group.TopEmu },
                new A.Extents { Cx = group.WidthEmu, Cy = group.HeightEmu },
                new A.ChildOffset { X = group.ChildLeftEmu, Y = group.ChildTopEmu },
                new A.ChildExtents { Cx = group.ChildWidthEmu, Cy = group.ChildHeightEmu })));
        foreach (var child in group.Children)
            output.Append(BuildElement(child, nativeIdsByElementId, slideContext, slidePart));
        return output;
    }

    private static P.Shape BuildShape(PresentationElement source, uint nativeId, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        var directFrame = semantic.DirectFrame;
        var transform = new A.Transform2D(
            new A.Offset { X = directFrame?.LeftEmu ?? semantic.LeftEmu, Y = directFrame?.TopEmu ?? semantic.TopEmu },
            new A.Extents { Cx = directFrame?.WidthEmu ?? semantic.WidthEmu, Cy = directFrame?.HeightEmu ?? semantic.HeightEmu });
        if (directFrame is not null)
        {
            transform.Rotation = directFrame.HasRotationAngle60000 ? directFrame.RotationAngle60000 : null;
            transform.HorizontalFlip = directFrame.HasFlipHorizontal ? directFrame.FlipHorizontal : null;
            transform.VerticalFlip = directFrame.HasFlipVertical ? directFrame.FlipVertical : null;
        }
        else
        {
            PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        }
        var properties = new P.ShapeProperties(transform);
        PptxCustomGeometryCodec.Apply(properties, semantic, source.Id);
        properties.Append(BuildFill(
            semantic.FillRgb,
            semantic.HasFillOpacityThousandthPercent ? semantic.FillOpacityThousandthPercent : (uint?)null));
        properties.Append(PptxLineStyleCodec.Build(semantic));
        ApplyShadow(properties, semantic.Shadow);
        var applicationProperties = new P.ApplicationNonVisualDrawingProperties();
        if (semantic.Placeholder is not null)
        {
            var nativePlaceholder = new P.PlaceholderShape { Index = semantic.Placeholder.Index };
            nativePlaceholder.SetAttribute(new OpenXmlAttribute("type", string.Empty, semantic.Placeholder.Type));
            applicationProperties.Append(nativePlaceholder);
        }
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, semantic.Accessibility);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                nonVisual,
                new P.NonVisualShapeDrawingProperties { TextBox = semantic.Geometry == "textbox" ? true : null },
                applicationProperties),
            properties,
            PptxTextCodec.Build(semantic, slideContext));
    }

    private static bool ApplyGroup(
        P.GroupShape source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location)
    {
        if (original.ContentCase != PresentationElement.ContentOneofCase.Group || requested.ContentCase != PresentationElement.ContentOneofCase.Group)
            throw new CodecException("presentation_group_content_changed", $"Presentation slide {slideIndex + 1} {location} changed its group content type.", PartPath(slideContext.Owner));
        var sourceChildren = GroupElements(source);
        if (sourceChildren.Length != original.Group.Children.Count || sourceChildren.Length != requested.Group.Children.Count)
            throw new CodecException("presentation_group_topology_changed", $"Presentation slide {slideIndex + 1} {location} changed its fixed group topology.", PartPath(slideContext.Owner));

        var changed = !Equals(original.Group.Accessibility, requested.Group.Accessibility);
        PptxNonVisualAccessibilityCodec.ApplyBound(
            source.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            requested.Group.Accessibility,
            "group");
        if (requested.Name != original.Name ||
            requested.Group.LeftEmu != original.Group.LeftEmu || requested.Group.TopEmu != original.Group.TopEmu ||
            requested.Group.WidthEmu != original.Group.WidthEmu || requested.Group.HeightEmu != original.Group.HeightEmu ||
            requested.Group.ChildLeftEmu != original.Group.ChildLeftEmu || requested.Group.ChildTopEmu != original.Group.ChildTopEmu ||
            requested.Group.ChildWidthEmu != original.Group.ChildWidthEmu || requested.Group.ChildHeightEmu != original.Group.ChildHeightEmu)
        {
            source.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name = requested.Name;
            var transform = source.GroupShapeProperties!.GetFirstChild<A.TransformGroup>()!;
            transform.Offset!.X = requested.Group.LeftEmu;
            transform.Offset.Y = requested.Group.TopEmu;
            transform.Extents!.Cx = requested.Group.WidthEmu;
            transform.Extents.Cy = requested.Group.HeightEmu;
            transform.ChildOffset!.X = requested.Group.ChildLeftEmu;
            transform.ChildOffset.Y = requested.Group.ChildTopEmu;
            transform.ChildExtents!.Cx = requested.Group.ChildWidthEmu;
            transform.ChildExtents.Cy = requested.Group.ChildHeightEmu;
            changed = true;
        }

        for (var index = 0; index < sourceChildren.Length; index++)
        {
            var sourceChild = sourceChildren[index];
            var originalChild = original.Group.Children[index];
            var requestedChild = requested.Group.Children[index];
            var binding = requestedChild.Source ?? throw new CodecException(
                "missing_presentation_element_binding",
                $"Presentation slide {slideIndex + 1} {location} child {index + 1} is missing its source binding.",
                PartPath(slideContext.Owner));
            if (requestedChild.Id != originalChild.Id || binding.ShapeTreeIndex != index ||
                !binding.ElementSha256.Equals(HashElement(sourceChild), StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != originalChild.Source?.Editable ||
                binding.TextEditable != originalChild.Source?.TextEditable ||
                binding.AccessibilityEditable != originalChild.Source?.AccessibilityEditable ||
                !binding.SemanticSha256.Equals(originalChild.Source?.SemanticSha256 ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
                !SemanticHash(originalChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_element_binding_mismatch",
                    $"Presentation slide {slideIndex + 1} {location} child {index + 1} does not match its owner-local source binding.",
                    PartPath(slideContext.Owner));
            if (SemanticHash(requestedChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} child {index + 1} is read-only.", PartPath(slideContext.Owner));
            ApplyGroupChild(sourceChild, originalChild, requestedChild, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"{location} child {index + 1}");
            changed = true;
        }
        return changed;
    }

    private static void ApplyGroupChild(
        OpenXmlElement source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location)
    {
        if (source is P.Shape shape && requested.ContentCase == PresentationElement.ContentOneofCase.Shape && IsSimpleShape(shape))
            ApplyShape(shape, requested, slideContext);
        else if (source is P.Picture picture && requested.ContentCase == PresentationElement.ContentOneofCase.Image && PptxPictureCodec.TryRead(picture, slideContext, out _))
            PptxPictureCodec.Apply(picture, requested, slideContext);
        else if (source is P.GraphicFrame table && requested.ContentCase == PresentationElement.ContentOneofCase.Table && PptxTableCodec.TryRead(table, out _))
            PptxTableCodec.Apply(table, requested);
        else if (source is P.ConnectionShape connector && requested.ContentCase == PresentationElement.ContentOneofCase.Connector && PptxConnectorCodec.TryRead(connector, elementIdsByNativeId, out _))
            PptxConnectorCodec.Apply(connector, requested, nativeIdsByElementId);
        else if (source is P.GraphicFrame chart && requested.ContentCase == PresentationElement.ContentOneofCase.Chart && PptxChartCodec.TryRead(chart, slideContext, out _, out var chartEditable) && chartEditable)
        {
            var replacement = PptxChartCodec.Apply(chart, requested, slideContext);
            changedParts.Add(replacement.PartPath);
            replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
        }
        else if (source is P.GroupShape group && requested.ContentCase == PresentationElement.ContentOneofCase.Group && original.ContentCase == PresentationElement.ContentOneofCase.Group && TryReadGroup(group, original.Id, slideContext, elementIdsByNativeId, out _))
            _ = ApplyGroup(group, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, location);
        else
            throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} changed outside the bounded group-child profile.", PartPath(slideContext.Owner));
    }

    private static P.ShapeTree BasicShapeTree() => new(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L },
            new A.ChildOffset { X = 0L, Y = 0L },
            new A.ChildExtents { Cx = 0L, Cy = 0L })));

    private static P.ColorMap BasicColorMap() => new()
    {
        Background1 = A.ColorSchemeIndexValues.Light1,
        Text1 = A.ColorSchemeIndexValues.Dark1,
        Background2 = A.ColorSchemeIndexValues.Light2,
        Text2 = A.ColorSchemeIndexValues.Dark2,
        Accent1 = A.ColorSchemeIndexValues.Accent1,
        Accent2 = A.ColorSchemeIndexValues.Accent2,
        Accent3 = A.ColorSchemeIndexValues.Accent3,
        Accent4 = A.ColorSchemeIndexValues.Accent4,
        Accent5 = A.ColorSchemeIndexValues.Accent5,
        Accent6 = A.ColorSchemeIndexValues.Accent6,
        Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
    };

    private static A.Theme BasicTheme() => new(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" })) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty })) { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                new A.LineStyleList(
                    BasicThemeOutline(9_525),
                    BasicThemeOutline(25_400),
                    BasicThemeOutline(38_100)),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))) { Name = "Office" }))
    { Name = "Office Clean Room" };

    private static A.Outline BasicThemeOutline(int width) => new(
        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
        new A.PresetDash { Val = A.PresetLineDashValues.Solid })
    { Width = width, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    private static OpenXmlElement[] ShapeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static OpenXmlElement[] GroupElements(P.GroupShape group) =>
        group.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static IReadOnlyDictionary<uint, string> NativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId)
    {
        var output = new Dictionary<uint, string>();
        CollectNativeElementIds(elements, ownerId, output);
        return output;
    }

    private static void CollectNativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId, IDictionary<uint, string> output)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            var elementId = $"{ownerId}/element/{index + 1}";
            var nativeId = elements[index].Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ??
                           elements[index].Descendants<P14.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value;
            if (nativeId is not null) output[nativeId.Value] = elementId;
            if (elements[index] is P.GroupShape group)
                CollectNativeElementIds(GroupElements(group), elementId, output);
        }
    }

    private static (long Left, long Top, long Width, long Height) ReadFrame(OpenXmlElement element)
    {
        if (element is P.GraphicFrame graphicFrame && graphicFrame.Transform?.Offset is { } graphicOffset && graphicFrame.Transform.Extents is { } graphicExtents)
            return (graphicOffset.X?.Value ?? 0, graphicOffset.Y?.Value ?? 0, graphicExtents.Cx?.Value ?? 0, graphicExtents.Cy?.Value ?? 0);
        if (element is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { Offset: { } groupOffset, Extents: { } groupExtents })
            return (groupOffset.X?.Value ?? 0, groupOffset.Y?.Value ?? 0, groupExtents.Cx?.Value ?? 0, groupExtents.Cy?.Value ?? 0);
        if (element is P.ContentPart contentPart && contentPart.Transform2D is { Offset: { } contentOffset, Extents: { } contentExtents })
            return (contentOffset.X?.Value ?? 0, contentOffset.Y?.Value ?? 0, contentExtents.Cx?.Value ?? 0, contentExtents.Cy?.Value ?? 0);
        var transform = element.Descendants<A.Transform2D>().FirstOrDefault();
        if (transform?.Offset is not null && transform.Extents is not null)
            return (transform.Offset.X?.Value ?? 0, transform.Offset.Y?.Value ?? 0, transform.Extents.Cx?.Value ?? 0, transform.Extents.Cy?.Value ?? 0);
        var offset = element.Descendants<A.Offset>().FirstOrDefault();
        var extents = element.Descendants<A.Extents>().FirstOrDefault();
        return (offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, extents?.Cx?.Value ?? 0, extents?.Cy?.Value ?? 0);
    }

    private static string Geometry(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.TextBox?.Value == true) return "textbox";
        if (shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>() is not null) return "custom";
        var value = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;
        if (value is null) return "rect";
        return value.Equals(A.ShapeTypeValues.Ellipse) ? "ellipse" :
            value.Equals(A.ShapeTypeValues.RoundRectangle) ? "roundRect" :
            value.Equals(A.ShapeTypeValues.Line) ? "line" :
            value.Equals(A.ShapeTypeValues.Rectangle) ? "rect" : value.ToString() ?? "rect";
    }

    private static string ElementName(OpenXmlElement element, int index) =>
        element.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ??
        element.Descendants<P14.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ??
        $"{element.LocalName} {index + 1}";

    private static string DescendantText(OpenXmlElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants<A.Text>().Select(text => text.Text));

    private static string SemanticHash(PresentationElement element)
    {
        var semantic = element.Clone();
        var placementEditable = semantic.ContentCase == PresentationElement.ContentOneofCase.Opaque && semantic.Source?.Editable == true;
        ClearElementIdentity(semantic);
        if (semantic.ContentCase == PresentationElement.ContentOneofCase.Shape)
        {
            PptxTextCodec.NormalizeSemantics(semantic.Shape);
            PptxLineStyleCodec.NormalizeSemantics(semantic.Shape);
            semantic.Shape.FillRgb = string.IsNullOrWhiteSpace(semantic.Shape.FillRgb) ? string.Empty : PptxColor.Normalize(semantic.Shape.FillRgb);
        }
        else if (placementEditable) semantic.Opaque.RawXml = string.Empty;
        return Hash(semantic.ToByteArray());
    }

    private static void ClearElementIdentity(PresentationElement element)
    {
        element.Id = string.Empty;
        element.Source = null;
        if (element.ContentCase != PresentationElement.ContentOneofCase.Group) return;
        foreach (var child in element.Group.Children) ClearElementIdentity(child);
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string DataPartPath(DataPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
    private static byte[] PartBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
    private static byte[] DataPartBytes(DataPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
    private static void CopyPartBytes(OpenXmlPart source, OpenXmlPart target)
    {
        using var input = source.GetStream(FileMode.Open, FileAccess.Read);
        using var output = target.GetStream(FileMode.Create, FileAccess.Write);
        input.CopyTo(output);
    }
    private static void CopyDataPartBytes(DataPart source, DataPart target)
    {
        using var input = source.GetStream(FileMode.Open, FileAccess.Read);
        target.FeedData(input);
    }
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static PptxAssetCatalog ValidateEnvelope(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        if (envelope.ProtocolVersion != CodecProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Presentation || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Presentation)
            throw new CodecException("invalid_presentation_artifact", "Artifact envelope does not contain a presentation payload.");
        if (envelope.Presentation.Slides.Count == 0)
            throw new CodecException("missing_slides", "Presentation must contain at least one slide.");
        if ((uint)envelope.Presentation.Slides.Count > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"Presentation has {envelope.Presentation.Slides.Count} slides and exceeds max_sheets ({limits.MaxSheets}).");
        if (envelope.Presentation.SlideWidthEmu < 0 || envelope.Presentation.SlideHeightEmu < 0 || envelope.Presentation.SlideWidthEmu > int.MaxValue || envelope.Presentation.SlideHeightEmu > int.MaxValue)
            throw new CodecException("invalid_slide_size", "Presentation slide dimensions must fit the PresentationML signed 32-bit EMU range.");
        var assetCatalog = new PptxAssetCatalog(envelope.Assets, limits);
        var hasSourcePackage = envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false };
        PptxViewPropertiesCodec.Validate(envelope.Presentation.ViewProperties, hasSourcePackage);

        if (envelope.Presentation.Masters.Count > 64)
            throw new CodecException("presentation_master_budget_exceeded", "Presentation cannot contain more than 64 slide masters.");
        if ((uint)envelope.Presentation.Layouts.Count > limits.MaxSheets)
            throw new CodecException("presentation_layout_budget_exceeded", $"Presentation has {envelope.Presentation.Layouts.Count} layouts and exceeds max_sheets ({limits.MaxSheets}).");
        ulong items = 0;
        PptxCustomShowCodec.Validate(envelope.Presentation, hasSourcePackage, limits, ref items);
        PptxSectionCodec.Validate(envelope.Presentation, hasSourcePackage, limits, ref items);
        var masterIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var master in envelope.Presentation.Masters)
        {
            if (string.IsNullOrWhiteSpace(master.Id) || !masterIds.Add(master.Id))
                throw new CodecException("invalid_presentation_master", "Presentation master IDs must be non-empty and unique.");
            if (master.Name.Length > 1_024)
                throw new CodecException("invalid_presentation_master", $"Presentation master {master.Id} name exceeds 1024 characters.");
            PptxMasterTextStylesCodec.Validate(master.TextStyles);
            PptxBackgroundCodec.Validate(master.Background);
            ValidatePlaceholders(master.Id, master.Placeholders, assetCatalog, limits, ref items);
            foreach (var paragraph in MasterStyleParagraphs(master.TextStyles))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        var layoutIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var layout in envelope.Presentation.Layouts)
        {
            if (string.IsNullOrWhiteSpace(layout.Id) || !layoutIds.Add(layout.Id))
                throw new CodecException("invalid_presentation_layout", "Presentation layout IDs must be non-empty and unique.");
            if (!masterIds.Contains(layout.MasterId))
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} references missing master {layout.MasterId}.");
            if (layout.Name.Length > 1_024 || layout.Type.Length > 128)
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} has invalid name or type metadata.");
            PptxBackgroundCodec.Validate(layout.Background);
            ValidatePlaceholders(layout.Id, layout.Placeholders, assetCatalog, limits, ref items);
        }

        var mastersById = envelope.Presentation.Masters.ToDictionary(master => master.Id, StringComparer.Ordinal);
        var layoutsById = envelope.Presentation.Layouts.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
        if (!hasSourcePackage)
        {
            foreach (var master in envelope.Presentation.Masters)
                foreach (var placeholder in master.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, master.Id);
            foreach (var layout in envelope.Presentation.Layouts)
            {
                _ = SourceFreeLayoutType(layout);
                foreach (var placeholder in layout.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, layout.Id);
            }
        }

        for (var slideIndex = 0; slideIndex < envelope.Presentation.Slides.Count; slideIndex++)
        {
            var slide = envelope.Presentation.Slides[slideIndex];
            PptxSpeakerNotesCodec.Validate(slide.SpeakerNotes);
            PptxBackgroundCodec.Validate(slide.Background);
            PptxTransitionCodec.Validate(slide.Transition);
            PptxLegacyCommentsCodec.Validate(slide, slideIndex);
            PptxModernCommentsCodec.Validate(slide, slideIndex, hasSourcePackage);
            if (!string.IsNullOrWhiteSpace(slide.LayoutId) && !layoutIds.Contains(slide.LayoutId))
                throw new CodecException("invalid_presentation_layout", $"Presentation slide {slide.Id} references missing layout {slide.LayoutId}.");
            if (!hasSourcePackage)
            {
                var placeholderShapes = slide.Elements.Where(element =>
                    element.ContentCase == PresentationElement.ContentOneofCase.Shape && element.Shape.Placeholder is not null).ToArray();
                if (placeholderShapes.Length > 0)
                {
                    if (string.IsNullOrWhiteSpace(slide.LayoutId) || !layoutsById.TryGetValue(slide.LayoutId, out var layout))
                        throw new CodecException("invalid_presentation_layout", $"Source-free presentation slide {slide.Id} has placeholders but no explicit layout binding.");
                    var master = mastersById[layout.MasterId];
                    foreach (var element in placeholderShapes)
                    {
                        var placeholder = element.Shape.Placeholder;
                        if (placeholder.InheritsGeometry || element.Shape.DirectFrame is null)
                            throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
                        ValidateSourceFreeTextPlaceholderIdentity(placeholder, element.Id);
                        if (!master.Placeholders.Concat(layout.Placeholders).Any(candidate =>
                                candidate.Type.Equals(placeholder.Type, StringComparison.Ordinal) && candidate.Index == placeholder.Index))
                            throw new CodecException("presentation_placeholder_binding_mismatch", $"Source-free presentation slide placeholder {element.Id} has no matching master/layout placeholder.");
                    }
                }
            }
            foreach (var element in slide.Elements)
                ValidatePresentationElement(element, hasSourcePackage, assetCatalog, limits, ref items, 0);
        }
        return assetCatalog;
    }

    private static void ValidatePresentationElement(
        PresentationElement element,
        bool hasSourcePackage,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items,
        int depth)
    {
        items++;
        if (items > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        if (depth > 16)
            throw new CodecException("presentation_group_depth_exceeded", "Presentation groups cannot be nested more than 16 levels.");
        if (string.IsNullOrWhiteSpace(element.Id) || element.Id.Length > 1_024 || element.Name.Length > 1_024)
            throw new CodecException("invalid_presentation_element", "Presentation element IDs and names must be bounded non-empty metadata.");

        // The source-preserving path re-reads this exact bound element before
        // applying any request, verifies its native and semantic hashes, and
        // rejects every semantic change when Editable is false. Treat that
        // contract as one opaque branch here rather than re-implementing a
        // partial validator for every DrawingML feature we deliberately do
        // not model. The envelope-wide item/depth/id bounds above still apply.
        if (hasSourcePackage && element.Source?.Editable == false) return;

        if (element.ContentCase == PresentationElement.ContentOneofCase.Shape)
        {
            if (element.Shape.HasUseBackgroundFill && !hasSourcePackage)
                throw new CodecException("unsupported_presentation_features", $"Presentation shape {element.Id} cannot author useBgFill without a validated source package.");
            if (element.Shape.Placeholder is not null && !hasSourcePackage &&
                (element.Shape.Placeholder.InheritsGeometry || element.Shape.DirectFrame is null))
                throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
            if (element.Shape.Placeholder is not null && element.Shape.Transform is not null)
                throw new CodecException("invalid_presentation_transform", $"Presentation placeholder shape {element.Id} cannot carry an ordinary shape transform.");
            var inheritedPlaceholderGeometry = element.Shape.Placeholder?.InheritsGeometry == true &&
                element.Shape.DirectFrame is null && element.Source?.Editable == false;
            var freeLine = element.Shape.Geometry == "line";
            var invalidExtent = freeLine
                ? element.Shape.WidthEmu == 0 && element.Shape.HeightEmu == 0
                : element.Shape.WidthEmu == 0 || element.Shape.HeightEmu == 0;
            if ((!inheritedPlaceholderGeometry && (element.Shape.LeftEmu < 0 || element.Shape.TopEmu < 0 ||
                    element.Shape.WidthEmu < 0 || element.Shape.HeightEmu < 0 ||
                    invalidExtent)) ||
                element.Shape.LineWidthEmu < 0 || element.Shape.LineWidthEmu > int.MaxValue)
                throw new CodecException(
                    "invalid_presentation_frame",
                    $"Presentation shape {element.Id} has an invalid {element.Shape.Geometry} frame " +
                    $"({element.Shape.LeftEmu},{element.Shape.TopEmu},{element.Shape.WidthEmu},{element.Shape.HeightEmu}).");
            if (element.Shape.DirectFrame is not null)
            {
                if (element.Shape.Placeholder is null || element.Shape.Placeholder.InheritsGeometry)
                    throw new CodecException("invalid_presentation_placeholder", $"Presentation shape {element.Id} has inconsistent direct placeholder geometry.");
                PptxPlaceholderCodec.ValidateDirectFrame(element.Shape.DirectFrame, element.Id);
            }
            if (element.Shape.Geometry is not ("rect" or "ellipse" or "roundRect" or "textbox" or "line" or "custom"))
                throw new CodecException("unsupported_presentation_geometry", $"Presentation shape {element.Id} uses unsupported geometry {element.Shape.Geometry}.");
            if (freeLine && element.Shape.Placeholder is not null)
                throw new CodecException("unsupported_presentation_geometry", $"Presentation free line {element.Id} cannot be a placeholder.");
            PptxCustomGeometryCodec.Validate(element.Shape, element.Id);
            if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb)) PptxColor.Normalize(element.Shape.FillRgb);
            if (element.Shape.HasFillOpacityThousandthPercent &&
                (string.IsNullOrWhiteSpace(element.Shape.FillRgb) || element.Shape.FillOpacityThousandthPercent > 100_000))
                throw new CodecException("invalid_presentation_fill", $"Presentation shape {element.Id} has invalid solid-fill opacity.");
            PptxLineStyleCodec.Validate(element.Shape, element.Id);
            PptxShapeTransformCodec.Validate(element.Shape.Transform, element.Id);
            ValidateShadow(element.Shape.Shadow, element.Id);
            PptxNonVisualAccessibilityCodec.Validate(element.Shape.Accessibility, element.Id);
            PptxTextCodec.Validate(element.Shape);
            foreach (var paragraph in element.Shape.TextBody?.Paragraphs ?? [])
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Image)
            PptxPictureCodec.Validate(element.Image, element.Id, assetCatalog, sourceBound: hasSourcePackage && element.Source is not null);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Table)
        {
            // Imported DrawingML tables may use a graphic-frame scale that
            // differs from the table grid coordinate space.  Keep authored
            // tables strict, while allowing a trusted source-bound table to
            // retain that bounded scale during a no-op or local edit.
            PptxTableCodec.Validate(
                element.Table,
                element.Id,
                allowScaledFrame: hasSourcePackage && element.Source is not null);
            items += checked((ulong)element.Table.Rows.Sum(row => row.Cells.Count));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Connector)
            PptxConnectorCodec.Validate(element.Connector, element.Id, element.Name);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Chart)
        {
            PptxChartCodec.Validate(element.Chart, element.Id, element.Name);
            items += checked((ulong)(
                element.Chart.Series.Sum(series => series.Values.Count) +
                element.Chart.ComboSeries.Sum(entry => entry.Series?.Values.Count ?? 0)));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
        {
            var group = element.Group;
            PptxNonVisualAccessibilityCodec.Validate(group.Accessibility, element.Id, "group");
            if (group.LeftEmu < 0 || group.TopEmu < 0 || group.WidthEmu <= 0 || group.HeightEmu <= 0 ||
                group.ChildWidthEmu <= 0 || group.ChildHeightEmu <= 0 || group.Children.Count == 0)
                throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} requires positive outer/child extents and at least one child.");
            var childIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var child in group.Children)
            {
                if (!childIds.Add(child.Id))
                    throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} contains duplicate child ID {child.Id}.");
                if (child.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                    throw new CodecException("unsupported_presentation_features", $"Presentation group {element.Id} contains a source-free or semantically edited opaque child.");
                ValidatePresentationElement(child, hasSourcePackage, assetCatalog, limits, ref items, depth + 1);
            }
        }
        else if (element.ContentCase != PresentationElement.ContentOneofCase.Opaque)
            throw new CodecException("missing_presentation_element_content", $"Presentation element {element.Id} has no content.");
        else if (element.Source?.Editable == true)
        {
            if (element.Opaque.LeftEmu < 0 || element.Opaque.TopEmu < 0 || element.Opaque.WidthEmu <= 0 || element.Opaque.HeightEmu <= 0)
                throw new CodecException("invalid_presentation_frame", $"Presentation native object {element.Id} has an invalid frame.");
        }
    }

    private static void ValidatePlaceholders(
        string ownerId,
        IList<PresentationPlaceholder> placeholders,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items)
    {
        if (placeholders.Count > 128)
            throw new CodecException("presentation_placeholder_budget_exceeded", $"Presentation owner {ownerId} exceeds the 128-placeholder budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var placeholder in placeholders)
        {
            items++;
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
            if (!ids.Add(placeholder.Id))
                throw new CodecException("invalid_presentation_placeholder", $"Presentation owner {ownerId} contains duplicate placeholder ID {placeholder.Id}.");
            PptxPlaceholderCodec.Validate(placeholder);
            foreach (var paragraph in (placeholder.TextBody?.Paragraphs ?? []).Concat(placeholder.TextBody?.ListStyles ?? []))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
    }

    private static void AssertPackagePartsUnchangedExcept(byte[] sourceBytes, byte[] outputBytes, HashSet<string> allowedPaths)
    {
        var before = PackagePartHashes(sourceBytes);
        var after = PackagePartHashes(outputBytes);
        var inventoryChanges = before.Keys.Concat(after.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !before.ContainsKey(path) || !after.ContainsKey(path))
            .Where(path => !allowedPaths.Contains(path))
            .Take(8)
            .ToArray();
        if (inventoryChanges.Length > 0)
            throw new CodecException("presentation_package_topology_changed", $"Source-preserving PPTX export changed unowned OPC part inventory: {string.Join(", ", inventoryChanges)}.");
        var changed = before.Keys.Intersect(after.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(path => !before[path].Equals(after[path], StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var unexpected = changed.Where(path => !allowedPaths.Contains(path)).Take(8).ToArray();
        if (unexpected.Length > 0)
            throw new CodecException("presentation_unowned_part_changed", $"Source-preserving PPTX export changed unowned package parts: {string.Join(", ", unexpected)}.");
    }

    private static void AssertOpaqueGraphMatchesWithModeledAdditions(
        OpaqueOpcGraph expected,
        OpaqueOpcGraph actual,
        IReadOnlySet<string> allowedAddedRelationshipIds,
        IReadOnlySet<string> allowedAddedPartPaths,
        IReadOnlyDictionary<string, string> allowedChangedPartHashes,
        IReadOnlySet<string> removedSourcePartPaths,
        IReadOnlySet<string> removedSourceRelationshipKeys)
    {
        var guarded = actual.Clone();
        var removed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in guarded.PackageRelationships.ToArray())
        {
            var key = $"{relationship.SourcePath}\0{relationship.Id}";
            if (!allowedAddedRelationshipIds.Contains(key)) continue;
            guarded.PackageRelationships.Remove(relationship);
            removed.Add(key);
        }
        if (!removed.SetEquals(allowedAddedRelationshipIds))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX relationship additions do not match the relationships written to the package.");
        var removedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in guarded.Parts.ToArray())
        {
            if (!allowedAddedPartPaths.Contains(part.Path)) continue;
            guarded.Parts.Remove(part);
            removedParts.Add(part.Path);
        }
        if (!removedParts.SetEquals(allowedAddedPartPaths))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX additions do not match the parts written to the package.");
        foreach (var (path, requestedHash) in allowedChangedPartHashes)
        {
            var before = expected.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            var after = guarded.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (before is null || after is null ||
                !before.ContentType.Equals(after.ContentType, StringComparison.OrdinalIgnoreCase) ||
                !after.Sha256.Equals(requestedHash, StringComparison.OrdinalIgnoreCase) ||
                !before.Relationships.SequenceEqual(after.Relationships))
                throw new CodecException("opaque_content_not_preserved", $"Modeled PPTX OLE workbook replacement did not preserve the package contract for {path}.", path);
        }
        PackageGuards.AssertOpaqueGraphMatches(
            expected,
            guarded,
            "opaque_content_not_preserved",
            ignoreRelationship: relationship =>
                removedSourcePartPaths.Contains(relationship.SourcePath) ||
                removedSourceRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"),
            ignorePart: part => allowedChangedPartHashes.ContainsKey(part.Path) || removedSourcePartPaths.Contains(part.Path));
    }

    private static void AssertPlannedPartsRemoved(byte[] sourceBytes, byte[] outputBytes, IReadOnlySet<string> removedPartPaths)
    {
        if (removedPartPaths.Count == 0) return;
        var before = PackagePartHashes(sourceBytes);
        var after = PackagePartHashes(outputBytes);
        var retained = removedPartPaths
            .Where(path => before.ContainsKey(path) && after.ContainsKey(path))
            .Take(8)
            .ToArray();
        if (retained.Length > 0)
            throw new CodecException(
                "presentation_delete_incomplete",
                $"Source-preserving PPTX deletion retained planned package parts: {string.Join(", ", retained)}.");
    }

    private static bool IsNumberedSlidePath(string path)
    {
        const string prefix = "ppt/slides/slide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedMasterPath(string path)
    {
        const string prefix = "ppt/slideMasters/slideMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedLayoutPath(string path)
    {
        const string prefix = "ppt/slideLayouts/slideLayout";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlidePath(string path)
    {
        const string prefix = "ppt/notesSlides/notesSlide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlideRelationshipPath(string path)
    {
        const string prefix = "ppt/notesSlides/_rels/notesSlide";
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesMasterPath(string path)
    {
        const string prefix = "ppt/notesMasters/notesMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesMasterRelationshipPath(string path)
    {
        const string prefix = "ppt/notesMasters/_rels/notesMaster";
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedThemePath(string path)
    {
        const string suffix = ".xml";
        foreach (var prefix in new[] { "ppt/theme/theme", "ppt/slideMasters/theme/theme" })
            if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
                path[prefix.Length..^suffix.Length].Length > 0 &&
                path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit))
                return true;
        return false;
    }

    private static bool IsNumberedCommentsPath(string path)
    {
        const string prefix = "ppt/comments/comment";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsLegacyCommentAuthorsPath(string path) =>
        path.Equals("ppt/commentAuthors.xml", StringComparison.OrdinalIgnoreCase);

    private static void ValidatePreservedSlideElements(byte[] sourceBytes, byte[] outputBytes, PresentationArtifact requested, EffectiveCodecLimits limits)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var orderedSourceSlides = OrderedSlideParts(sourcePackage);
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (sourceTargets.Length != requested.Slides.Count || outputSlides.Length != requested.Slides.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide topology changed during source-preserving export.");
        var retainedSourceSlideParts = sourceTargets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            var outputSlide = outputSlides[targetIndex];
            if (!target.IsClone && !PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
            if (target.IsClone)
            {
                if (PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_postwrite_clone_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source slide copy.", PartPath(outputSlide));
                if (target.Target.ElementDeletions.Count == 0)
                {
                    if (!HashElement(target.Source.Part.Slide!).Equals(HashElement(outputSlide.Slide!), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_clone_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source slide copy.", PartPath(outputSlide));
                    PptxSlideCloneCodec.Validate(target.Source, outputSlide, retainedSourceSlideParts);
                }
                else
                {
                    ValidateCloneElementProjection(target.Source, outputSlide, target.Target);
                }
            }
            else
            {
                PptxSpeakerNotesCodec.ValidateSourceBoundOutput(
                    sourcePresentationPart,
                    outputPresentationPart,
                    target.Source.Part,
                    outputSlide,
                    target.Target,
                    targetIndex);
                PptxLegacyCommentsCodec.ValidateSourceBoundOutput(
                    sourcePresentationPart,
                    outputPresentationPart,
                    target.Source.Part,
                    outputSlide,
                    target.Target,
                    targetIndex);
            }
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits);
        var outputAssets = new PptxAssetCatalog([], limits);
        var customShowCatalog = PptxCustomShowCatalog.From(requested.CustomShows);

        for (var slideIndex = 0; slideIndex < requested.Slides.Count; slideIndex++)
        {
            if (sourceTargets[slideIndex].IsClone) continue;
            var target = sourceTargets[slideIndex];
            var sourceSlide = target.Source.Part;
            var outputSlide = outputSlides[slideIndex];
            var outputRoot = outputSlide.Slide ??
                throw new CodecException("missing_slide_root", $"PPTX output slide {slideIndex + 1} has no slide root.", PartPath(outputSlide));
            var sourceRoot = sourceSlide.Slide!;
            var outputName = outputRoot.CommonSlideData?.Name?.Value ?? string.Empty;
            if (!string.Equals(outputName, requested.Slides[slideIndex].Name, StringComparison.Ordinal))
                throw new CodecException(
                    "presentation_postwrite_slide_name_mismatch",
                    $"PPTX slide {slideIndex + 1} name does not match the requested source-bound value.",
                    PartPath(outputSlide));
            if (!PptxSlideVisibilityCodec.Matches(requested.Slides[slideIndex], sourceRoot, outputRoot))
                throw new CodecException(
                    "presentation_postwrite_slide_visibility_mismatch",
                    $"PPTX slide {slideIndex + 1} visibility does not match the requested source-bound value.",
                    PartPath(outputSlide));
            var sourceTransition = PptxTransitionCodec.Read(sourceRoot);
            var outputTransition = PptxTransitionCodec.Read(outputRoot);
            var requestedTransition = requested.Slides[slideIndex].Transition;
            var transitionChanged = !PptxTransitionCodec.SemanticHash(requestedTransition)
                .Equals(PptxTransitionCodec.SemanticHash(sourceTransition), StringComparison.OrdinalIgnoreCase);
            if (!transitionChanged && !PptxTransitionCodec.ElementHash(sourceRoot)
                    .Equals(PptxTransitionCodec.ElementHash(outputRoot), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unchanged_transition_modified",
                    $"PPTX slide {slideIndex + 1} unchanged transition was modified during export.",
                    PartPath(outputSlide));
            if (!PptxTransitionCodec.SemanticHash(outputTransition)
                    .Equals(PptxTransitionCodec.SemanticHash(requestedTransition), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_transition_semantics_mismatch",
                    $"PPTX slide {slideIndex + 1} transition does not match requested semantics after export.",
                    PartPath(outputSlide));
            var sourceTimingElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"PPTX source slide {slideIndex + 1} has no shape tree.", PartPath(sourceSlide)));
            var sourceTimingIds = NativeElementIds(sourceTimingElements, requested.Slides[slideIndex].Id);
            P.Slide? previousSourceRoot = null;
            IReadOnlyDictionary<uint, string>? previousSourceTimingIds = null;
            string? previousSourceSlideId = null;
            if (target.Source.Index > 0 && PptxTimingCodec.HasMorph(sourceRoot))
            {
                previousSourceRoot = orderedSourceSlides[target.Source.Index - 1].Slide;
                var previousSourceTarget = sourceTargets.FirstOrDefault(candidate =>
                    !candidate.IsClone && candidate.Source.Index == target.Source.Index - 1);
                previousSourceSlideId = previousSourceTarget?.Target.Id;
                var previousSourceTree = previousSourceRoot?.CommonSlideData?.ShapeTree;
                if (previousSourceTree is not null && !string.IsNullOrWhiteSpace(previousSourceSlideId))
                    previousSourceTimingIds = NativeElementIds(ShapeElements(previousSourceTree), previousSourceSlideId);
            }
            var sourceTiming = PptxTimingCodec.Read(
                sourceRoot,
                sourceTimingIds,
                previousSourceRoot,
                previousSourceTimingIds,
                previousSourceSlideId);
            P.Slide? previousOutputRoot = null;
            IReadOnlyDictionary<uint, string>? previousOutputTimingIds = null;
            string? previousOutputSlideId = null;
            if (slideIndex > 0 && PptxTimingCodec.HasMorph(outputRoot))
            {
                previousOutputRoot = outputSlides[slideIndex - 1].Slide;
                previousOutputSlideId = requested.Slides[slideIndex - 1].Id;
                var previousOutputTree = previousOutputRoot?.CommonSlideData?.ShapeTree;
                if (previousOutputTree is not null)
                    previousOutputTimingIds = NativeElementIds(ShapeElements(previousOutputTree), previousOutputSlideId);
            }
            var outputTiming = PptxTimingCodec.Read(
                outputRoot,
                sourceTimingIds,
                previousOutputRoot,
                previousOutputTimingIds,
                previousOutputSlideId);
            var requestedTimingHash = PptxTimingCodec.SemanticHash(requested.Slides[slideIndex].Animations, requested.Slides[slideIndex].Morph);
            var requestedOpaqueNoop = sourceTiming.Present && !sourceTiming.Editable && requested.Slides[slideIndex].Animations.Count == 0 && requested.Slides[slideIndex].Morph is null;
            if (!requestedOpaqueNoop && !outputTiming.SemanticSha256.Equals(requestedTimingHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_timing_semantics_mismatch",
                    $"PPTX slide {slideIndex + 1} timing does not match requested animation semantics after export.",
                    PartPath(outputSlide));
            var sourceContext = new PptxPartContext(sourceSlide, sourceIdByPartPath, assets: sourceAssets, customShows: customShowCatalog);
            var outputContext = new PptxPartContext(outputSlides[slideIndex], outputIdByPartPath, assets: outputAssets, customShows: customShowCatalog);
            var before = ShapeElements(sourceSlide.Slide!.CommonSlideData!.ShapeTree!);
            var after = ShapeElements(outputRoot.CommonSlideData!.ShapeTree!);
            var (elements, authoredElements) = SplitSourceBoundElements(requested.Slides[slideIndex], before.Length, slideIndex, outputSlide);
            var authoredNativeIds = AuthoredOverlayNativeIds(before, authoredElements, slideIndex, outputSlide);
            var outputNativeIdSet = after.SelectMany(PptxElementDeletionCodec.NativeIds).ToHashSet();
            var afterIds = sourceTimingIds
                .Where(entry => outputNativeIdSet.Contains(entry.Key))
                .ToDictionary(entry => entry.Key, entry => entry.Value);
            foreach (var (elementId, nativeId) in authoredNativeIds)
                afterIds.Add(nativeId, elementId);
            var deletions = requested.Slides[slideIndex].ElementDeletions;
            if (before.Length != elements.Length + deletions.Count || after.Length != elements.Length + authoredElements.Length)
                throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} element topology changed during source-preserving export.", PartPath(outputSlide));
            var outputNativeIds = outputNativeIdSet;
            foreach (var deletion in deletions)
            {
                var binding = deletion.Source ?? throw new CodecException(
                    "missing_presentation_element_deletion_binding",
                    $"PPTX slide {slideIndex + 1} element deletion {deletion.Id} lost its source binding after export.",
                    PartPath(outputSlide));
                var sourceElementIndex = checked((int)binding.ShapeTreeIndex);
                if (sourceElementIndex >= before.Length ||
                    !binding.ElementSha256.Equals(HashElement(before[sourceElementIndex]), StringComparison.OrdinalIgnoreCase) ||
                    PptxElementDeletionCodec.NativeIds(before[sourceElementIndex]).Overlaps(outputNativeIds))
                    throw new CodecException(
                        "presentation_postwrite_element_delete_mismatch",
                        $"PPTX slide {slideIndex + 1} deleted element {sourceElementIndex + 1} remains or no longer matches its source binding.",
                        PartPath(outputSlide));
                var plan = PptxElementDeletionCodec.Analyze(sourceSlide, before[sourceElementIndex], before);
                if (!plan.Supported)
                    throw new CodecException(
                        "presentation_postwrite_element_delete_mismatch",
                        $"PPTX slide {slideIndex + 1} deleted element {sourceElementIndex + 1} no longer satisfies its source deletion proof.",
                        PartPath(outputSlide));
            }
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                var request = elements[elementIndex];
                var binding = request.Source!;
                var sourceElementIndex = checked((int)binding.ShapeTreeIndex);
                if (sourceElementIndex >= before.Length)
                    throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} retained element {elementIndex + 1} has an invalid source index.", PartPath(outputSlide));
                var beforeElement = before[sourceElementIndex];
                var afterElement = after[elementIndex];
                var changed = !SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
                if (!changed)
                {
                    if (!HashElement(beforeElement).Equals(HashElement(afterElement), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unchanged_element_modified",
                            $"PPTX slide {slideIndex + 1} unchanged retained element {sourceElementIndex + 1} was modified during export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                {
                    if (!NativeObjectResidualHash(beforeElement).Equals(NativeObjectResidualHash(afterElement), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_native_content_changed",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputFrame = ReadFrame(afterElement);
                    if (!ElementName(afterElement, elementIndex).Equals(request.Name, StringComparison.Ordinal) ||
                        outputFrame.Left != request.Opaque.LeftEmu || outputFrame.Top != request.Opaque.TopEmu ||
                        outputFrame.Width != request.Opaque.WidthEmu || outputFrame.Height != request.Opaque.HeightEmu)
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} does not match the requested name/frame.",
                            PartPath(outputSlides[slideIndex]));
                    PptxDiagramTextCodec.ValidateSourceBoundOutput(
                        sourceSlide,
                        outputSlide,
                        beforeElement,
                        afterElement,
                        request.Opaque);
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Group)
                {
                    if (beforeElement is not P.GroupShape beforeGroup || afterElement is not P.GroupShape afterGroup)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    ValidateGroupOutput(beforeGroup, afterGroup, request, sourceContext, outputContext, afterIds, slideIndex, $"element {elementIndex + 1}");
                    var outputGroupSemantic = ReadElement(afterGroup, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputGroupSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Image)
                {
                    if (beforeElement is not P.Picture beforePicture || afterElement is not P.Picture afterPicture)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_picture_content_changed",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputPictureSemantic = ReadElement(afterPicture, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputPictureSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    if (beforeElement is not P.GraphicFrame beforeTable || afterElement is not P.GraphicFrame afterTable)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!TableResidualHash(beforeTable).Equals(TableResidualHash(afterTable), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_table_content_changed",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputTableSemantic = ReadElement(afterTable, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputTableSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Connector)
                {
                    if (beforeElement is not P.ConnectionShape beforeConnector || afterElement is not P.ConnectionShape afterConnector)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed unmodeled native content.", PartPath(outputSlides[slideIndex]));
                    var outputConnectorSemantic = ReadElement(afterConnector, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputConnectorSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Chart)
                {
                    if (beforeElement is not P.GraphicFrame beforeChart || afterElement is not P.GraphicFrame afterChart)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed unmodeled frame content.", PartPath(outputSlides[slideIndex]));
                    var outputChartSemantic = ReadElement(afterChart, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputChartSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (beforeElement is not P.Shape beforeShape || afterElement is not P.Shape afterShape)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited element {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                if (!ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_unmodeled_shape_content_changed",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} changed unmodeled native content.",
                        PartPath(outputSlides[slideIndex]));
                var outputSemantic = ReadElement(afterShape, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlides[slideIndex]));
            }
            for (var authoredIndex = 0; authoredIndex < authoredElements.Length; authoredIndex++)
            {
                var request = authoredElements[authoredIndex];
                var outputIndex = elements.Length + authoredIndex;
                var outputElement = after[outputIndex];
                var authoredOutputNativeIds = PptxElementDeletionCodec.NativeIds(outputElement).ToArray();
                var nativeTypeMatches = request.ContentCase switch
                {
                    PresentationElement.ContentOneofCase.Shape => outputElement is P.Shape,
                    PresentationElement.ContentOneofCase.Image => outputElement is P.Picture,
                    _ => false,
                };
                if (!nativeTypeMatches || authoredOutputNativeIds.Length != 1 || authoredOutputNativeIds[0] != authoredNativeIds[request.Id])
                    throw new CodecException(
                        "presentation_postwrite_authored_overlay_mismatch",
                        $"PPTX slide {slideIndex + 1} authored overlay {authoredIndex + 1} changed native type or identity during export.",
                        PartPath(outputSlide));
                var outputSemantic = ReadElement(outputElement, slideIndex, outputIndex, outputContext, elementIdsByNativeId: afterIds);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} authored overlay {authoredIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlide));
            }
        }
    }

    private static void ValidateCloneElementProjection(
        PptxSourceSlideEntry source,
        SlidePart output,
        PresentationSlide requested)
    {
        var sourceRoot = source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
        var outputRoot = output.Slide ??
            throw new CodecException("missing_slide_root", "Presentation cloned slide has no slide root.", PartPath(output));
        var sourceElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {source.Index + 1} has no shape tree.", PartPath(source.Part)));
        var outputElements = ShapeElements(outputRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", "Presentation cloned slide has no shape tree.", PartPath(output)));
        if (sourceElements.Length != requested.Elements.Count + requested.ElementDeletions.Count || outputElements.Length != requested.Elements.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX component clone output does not match its retained/deleted element projection.", PartPath(output));
        var outputNativeIdList = outputElements.SelectMany(PptxElementDeletionCodec.NativeIds).ToArray();
        if (outputNativeIdList.GroupBy(id => id).Any(group => group.Count() > 1))
            throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone retained duplicate native drawing IDs.", PartPath(output));
        var outputNativeIds = outputNativeIdList.ToHashSet();
        foreach (var deletion in requested.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", "PPTX component clone deletion lost its source binding.", PartPath(output));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length ||
                !binding.ElementSha256.Equals(HashElement(sourceElements[sourceIndex]), StringComparison.OrdinalIgnoreCase) ||
                PptxElementDeletionCodec.NativeIds(sourceElements[sourceIndex]).Overlaps(outputNativeIds))
                throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone retained a deleted source element or lost its source proof.", PartPath(output));
            var plan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[sourceIndex], sourceElements, allowDuplicateNativeIds: true);
            if (!plan.Supported)
                throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone deletion no longer satisfies its source deletion proof.", PartPath(output));
        }
        for (var retainedIndex = 0; retainedIndex < requested.Elements.Count; retainedIndex++)
        {
            var binding = requested.Elements[retainedIndex].Source ??
                throw new CodecException("missing_presentation_element_binding", "PPTX component clone retained element lost its source binding.", PartPath(output));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length ||
                !binding.ElementSha256.Equals(HashElement(sourceElements[sourceIndex]), StringComparison.OrdinalIgnoreCase) ||
                !HashElement(sourceElements[sourceIndex]).Equals(HashElement(outputElements[retainedIndex]), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_clone_mismatch", "PPTX component clone changed an element that was not selected for reuse.", PartPath(output));
        }
    }

    private static void ValidateGroupOutput(
        P.GroupShape before,
        P.GroupShape after,
        PresentationElement request,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        IReadOnlyDictionary<uint, string> afterIds,
        int slideIndex,
        string location)
    {
        if (!GroupShellResidualHash(before).Equals(GroupShellResidualHash(after), StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_unmodeled_group_content_changed", $"PPTX slide {slideIndex + 1} {location} changed unmodeled group-shell content.", PartPath(outputContext.Owner));
        var beforeChildren = GroupElements(before);
        var afterChildren = GroupElements(after);
        if (beforeChildren.Length != request.Group.Children.Count || afterChildren.Length != request.Group.Children.Count)
            throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} {location} group topology changed during export.", PartPath(outputContext.Owner));

        for (var index = 0; index < request.Group.Children.Count; index++)
        {
            var child = request.Group.Children[index];
            var binding = child.Source ?? throw new CodecException("missing_presentation_element_binding", $"PPTX slide {slideIndex + 1} {location} child {index + 1} is missing its source binding.", PartPath(outputContext.Owner));
            var changed = !SemanticHash(child).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                if (!HashElement(beforeChildren[index]).Equals(HashElement(afterChildren[index]), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_element_modified", $"PPTX slide {slideIndex + 1} {location} unchanged child {index + 1} was modified during export.", PartPath(outputContext.Owner));
                continue;
            }

            if (child.ContentCase == PresentationElement.ContentOneofCase.Group)
            {
                if (beforeChildren[index] is not P.GroupShape beforeGroup || afterChildren[index] is not P.GroupShape afterGroup)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} {location} child group {index + 1} changed native element type.", PartPath(outputContext.Owner));
                ValidateGroupOutput(beforeGroup, afterGroup, child, sourceContext, outputContext, afterIds, slideIndex, $"{location} child {index + 1}");
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Image)
            {
                if (beforeChildren[index] is not P.Picture beforePicture || afterChildren[index] is not P.Picture afterPicture ||
                    !PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_picture_content_changed", $"PPTX slide {slideIndex + 1} {location} child image {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Table)
            {
                if (beforeChildren[index] is not P.GraphicFrame beforeTable || afterChildren[index] is not P.GraphicFrame afterTable ||
                    !TableResidualHash(beforeTable).Equals(TableResidualHash(afterTable), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_table_content_changed", $"PPTX slide {slideIndex + 1} {location} child table {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Connector)
            {
                if (beforeChildren[index] is not P.ConnectionShape beforeConnector || afterChildren[index] is not P.ConnectionShape afterConnector ||
                    !ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} {location} child connector {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Chart)
            {
                if (beforeChildren[index] is not P.GraphicFrame beforeChart || afterChildren[index] is not P.GraphicFrame afterChart ||
                    !ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} {location} child chart {index + 1} changed unmodeled frame content.", PartPath(outputContext.Owner));
            }
            else
            {
                if (beforeChildren[index] is not P.Shape beforeShape || afterChildren[index] is not P.Shape afterShape ||
                    !ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_shape_content_changed", $"PPTX slide {slideIndex + 1} {location} child shape {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }

            var outputSemantic = ReadElement(afterChildren[index], request.Id, index, outputContext, elementIdsByNativeId: afterIds);
            if (!SemanticHash(outputSemantic).Equals(SemanticHash(child), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} {location} child {index + 1} does not match requested semantics after export.", PartPath(outputContext.Owner));
        }
    }

    private static string GroupShellResidualHash(P.GroupShape source)
    {
        var clone = (P.GroupShape)source.CloneNode(true);
        foreach (var child in GroupElements(clone)) child.Remove();
        if (clone.NonVisualGroupShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (clone.GroupShapeProperties?.GetFirstChild<A.TransformGroup>() is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
            transform.ChildOffset!.X = 0L;
            transform.ChildOffset.Y = 0L;
            transform.ChildExtents!.Cx = 1L;
            transform.ChildExtents.Cy = 1L;
        }
        return HashElement(clone);
    }

    private static void ValidatePreservedMasterAndLayoutContent(byte[] sourceBytes, byte[] outputBytes, PresentationArtifact requested, EffectiveCodecLimits limits)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var sourceGraph = ReadMasterGraph(sourcePresentationPart);
        var outputGraph = ReadMasterGraph(outputPresentationPart);
        if (sourceGraph.Length != requested.Masters.Count || outputGraph.Length != requested.Masters.Count)
            throw new CodecException("presentation_postwrite_master_topology_changed", "PPTX master topology changed during source-preserving export.");
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (outputSlides.Length != sourceTargets.Length)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            if (!target.IsClone && !PartPath(outputSlides[targetIndex]).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits);
        var outputAssets = new PptxAssetCatalog([], limits);
        var customShowCatalog = PptxCustomShowCatalog.From(requested.CustomShows);
        for (var masterIndex = 0; masterIndex < requested.Masters.Count; masterIndex++)
        {
            var before = sourceGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX source master {masterIndex + 1} has no root.");
            var after = outputGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX output master {masterIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceGraph[masterIndex].Part, sourceSlideMap, assets: sourceAssets, customShows: customShowCatalog);
            var outputContext = new PptxPartContext(outputGraph[masterIndex].Part, outputSlideMap, assets: outputAssets, customShows: customShowCatalog);
            if (!MasterResidualHash(before, sourceContext).Equals(MasterResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_master_content_changed",
                    $"PPTX master {masterIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputGraph[masterIndex].Part));
            var outputStyles = PptxMasterTextStylesCodec.Read(after, outputContext);
            if (!MasterTextStylesSemanticHash(outputStyles).Equals(MasterTextStylesSemanticHash(requested.Masters[masterIndex].TextStyles), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_semantics_mismatch",
                    $"PPTX master {masterIndex + 1} text styles do not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData)).Equals(BackgroundSemanticHash(requested.Masters[masterIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_background_mismatch",
                    $"PPTX master {masterIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Masters[masterIndex].Placeholders,
                requested.Masters[masterIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputGraph[masterIndex].Part));
        }
        var sourceLayouts = sourceGraph.SelectMany(master => master.Layouts).ToArray();
        var outputLayouts = outputGraph.SelectMany(master => master.Layouts).ToArray();
        if (sourceLayouts.Length != requested.Layouts.Count || outputLayouts.Length != requested.Layouts.Count)
            throw new CodecException("presentation_postwrite_layout_topology_changed", "PPTX layout topology changed during source-preserving export.");
        for (var layoutIndex = 0; layoutIndex < requested.Layouts.Count; layoutIndex++)
        {
            var before = sourceLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX source layout {layoutIndex + 1} has no root.");
            var after = outputLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX output layout {layoutIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceLayouts[layoutIndex].Part, sourceSlideMap, assets: sourceAssets, customShows: customShowCatalog);
            var outputContext = new PptxPartContext(outputLayouts[layoutIndex].Part, outputSlideMap, assets: outputAssets, customShows: customShowCatalog);
            if (!LayoutResidualHash(before, sourceContext).Equals(LayoutResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_layout_content_changed",
                    $"PPTX layout {layoutIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputLayouts[layoutIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData)).Equals(BackgroundSemanticHash(requested.Layouts[layoutIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_layout_background_mismatch",
                    $"PPTX layout {layoutIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputLayouts[layoutIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Layouts[layoutIndex].Placeholders,
                requested.Layouts[layoutIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputLayouts[layoutIndex].Part));
        }
    }

    private static void ValidatePlaceholderOutput(
        P.ShapeTree? sourceTree,
        P.ShapeTree? outputTree,
        IList<PresentationPlaceholder> requested,
        string ownerId,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        string partPath)
    {
        if (sourceTree is null || outputTree is null)
            throw new CodecException("missing_shape_tree", $"Presentation owner {ownerId} has no shape tree.", partPath);
        var before = PptxPlaceholderCodec.Read(sourceTree, ownerId, sourceContext);
        var after = PptxPlaceholderCodec.Read(outputTree, ownerId, outputContext);
        if (before.Count != requested.Count || after.Count != requested.Count)
            throw new CodecException("presentation_postwrite_placeholder_topology_changed", $"Presentation owner {ownerId} placeholder topology changed during export.", partPath);
        for (var index = 0; index < requested.Count; index++)
        {
            var request = requested[index];
            var binding = request.Source ?? throw new CodecException("missing_presentation_placeholder_binding", $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.", partPath);
            var changed = !PptxPlaceholderCodec.SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                var sourceShape = PptxPlaceholderCodec.BoundShape(sourceTree, before[index]);
                var outputShape = PptxPlaceholderCodec.BoundShape(outputTree, after[index]);
                if (sourceShape is null || outputShape is null ||
                    !PptxPlaceholderCodec.ElementHash(sourceShape).Equals(PptxPlaceholderCodec.ElementHash(outputShape), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_placeholder_modified", $"Presentation placeholder {index + 1} under {ownerId} was modified during export.", partPath);
            }
            if (after[index].Id != request.Id ||
                !PptxPlaceholderCodec.SemanticHash(after[index]).Equals(PptxPlaceholderCodec.SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_placeholder_semantics_mismatch", $"Presentation placeholder {index + 1} under {ownerId} does not match requested semantics after export.", partPath);
        }
    }

    private static SlidePart[] OrderedSlideParts(PresentationDocument package)
    {
        var presentationPart = package.PresentationPart ?? throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        return ResolveSlideParts(presentationPart, presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? []);
    }

    private static PptxMasterGraphEntry[] ReadMasterGraph(PresentationPart presentationPart)
    {
        var masterIds = presentationPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>().ToArray() ?? [];
        return masterIds.Select((masterId, masterIndex) =>
        {
            var relationshipId = masterId.RelationshipId?.Value ?? string.Empty;
            var masterPart = presentationPart.GetPartById(relationshipId) as SlideMasterPart ??
                throw new CodecException("missing_slide_master_part", $"Presentation master {masterIndex + 1} has an unresolved relationship.", "ppt/presentation.xml");
            var layoutIds = masterPart.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>().ToArray() ?? [];
            var layouts = layoutIds.Select((layoutId, layoutIndex) =>
            {
                var layoutRelationshipId = layoutId.RelationshipId?.Value ?? string.Empty;
                var layoutPart = masterPart.GetPartById(layoutRelationshipId) as SlideLayoutPart ??
                    throw new CodecException("missing_slide_layout_part", $"Presentation layout {layoutIndex + 1} under master {masterIndex + 1} has an unresolved relationship.", PartPath(masterPart));
                return new PptxLayoutGraphEntry(
                    layoutIndex,
                    $"presentation/master/{masterIndex + 1}/layout/{layoutIndex + 1}",
                    layoutRelationshipId,
                    layoutPart);
            }).ToArray();
            return new PptxMasterGraphEntry(
                masterIndex,
                $"presentation/master/{masterIndex + 1}",
                relationshipId,
                masterPart,
                layouts);
        }).ToArray();
    }

    private static string LayoutTypeName(P.SlideLayout source)
    {
        // `type` is a typed PresentationML attribute. Reading it through the
        // generic attribute accessor with an empty namespace is rejected by
        // Open XML SDK 3.x for typed elements, even though it is a valid
        // unqualified attribute in the serialized package. Use the generated
        // property so imported layouts from Office-authored templates remain
        // readable.
        var value = source.Type?.InnerText;
        return string.IsNullOrWhiteSpace(value) ? "custom" : value;
    }

    private static SlidePart[] ResolveSlideParts(PresentationPart presentationPart, IEnumerable<P.SlideId> slideIds) =>
        slideIds.Select(slideId => presentationPart.GetPartById(slideId.RelationshipId?.Value ?? string.Empty) as SlidePart ??
            throw new CodecException("missing_slide_part", "PPTX presentation contains an unresolved slide relationship.", "ppt/presentation.xml"))
        .ToArray();

    private static IReadOnlyDictionary<string, string> BuildCustomShowSlideIdMap(
        IEnumerable<(string RelationshipId, string PublicId)> entries)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var ambiguous = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (relationshipId, publicId) in entries)
        {
            if (relationshipId.Length == 0 || ambiguous.Contains(relationshipId)) continue;
            if (!result.TryAdd(relationshipId, publicId))
            {
                result.Remove(relationshipId);
                ambiguous.Add(relationshipId);
            }
        }
        return result;
    }

    private static IReadOnlyDictionary<uint, string> BuildSectionSlideIdMap(
        IEnumerable<(uint? NativeId, string PublicId)> entries)
    {
        var result = new Dictionary<uint, string>();
        var ambiguous = new HashSet<uint>();
        foreach (var (nativeId, publicId) in entries)
        {
            if (nativeId is not { } id || publicId.Length == 0 || ambiguous.Contains(id)) continue;
            if (!result.TryAdd(id, publicId))
            {
                result.Remove(id);
                ambiguous.Add(id);
            }
        }
        return result;
    }

    // Imported source topology keeps ordinary bindings and clone requests
    // separate. A clone is a new SlidePart with a verified origin, never a
    // second presentation reference to the same source part.
    private static PptxTargetSlideEntry[] BindSourcePreservingSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PresentationSlide> requested)
    {
        if (requested.Count == 0)
            throw new CodecException(
                "presentation_topology_changed",
                "Source-preserving PPTX export must retain at least one source slide.",
                "ppt/presentation.xml");
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var sourceSlides = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .ToArray();
        var targets = new PptxTargetSlideEntry[requested.Count];
        var seenSourceParts = new HashSet<SlidePart>();
        var seenCloneSourceParts = new HashSet<SlidePart>();
        for (var targetIndex = 0; targetIndex < requested.Count; targetIndex++)
        {
            var target = requested[targetIndex];
            var isClone = target.CloneSource is not null;
            if ((target.Source is null) == !isClone)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} must carry exactly one of source or clone_source.",
                    "ppt/presentation.xml");
            var binding = target.Source ?? target.CloneSource!;
            if (binding.SlideIndex >= sourceSlides.Length)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} references source slide {binding.SlideIndex + 1}, which does not exist.",
                    "ppt/presentation.xml");
            var source = sourceSlides[binding.SlideIndex];
            var sourceRoot = source.Part.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
            if (!binding.PartPath.Equals(PartPath(source.Part), StringComparison.OrdinalIgnoreCase) ||
                !binding.RelationshipId.Equals(source.RelationshipId, StringComparison.Ordinal) ||
                !binding.SlideXmlSha256.Equals(HashElement(sourceRoot), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} does not match its hash-bound source slide.",
                    PartPath(source.Part));
            if (!isClone && !seenSourceParts.Add(source.Part))
                throw new CodecException(
                    "presentation_topology_changed",
                    "Source-preserving PPTX export cannot bind more than one ordinary target to a source SlidePart.",
                    "ppt/presentation.xml");
            if (isClone && !seenCloneSourceParts.Add(source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    "The bounded source-preserving PPTX clone profile permits only one pending clone per source SlidePart.",
                    PartPath(source.Part));
            targets[targetIndex] = new PptxTargetSlideEntry(targetIndex, target, source, isClone);
        }
        return targets;
    }

    private static void DeleteUnrequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        OpaqueOpcGraph opaque,
        ISet<string> changedParts,
        ISet<string> removedSourcePartPaths)
    {
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var retainedParts = targets.Where(target => !target.IsClone).Select(target => target.Source.Part).ToHashSet();
        var removed = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .Where(source => !retainedParts.Contains(source.Part))
            .ToArray();
        if (removed.Length == 0) return;

        var plans = removed
            .Select(source => (Source: source, Plan: PptxSlideDeletionCodec.Analyze(presentationPart, source, opaque)))
            .ToArray();
        foreach (var (source, plan) in plans)
        {
            if (!plan.Supported) throw UnsupportedSourceSlideDelete(source, plan.BlockedReason);
        }
        var transactionPlan = PptxSlideDeletionCodec.AnalyzeTransaction(presentationPart, removed, opaque);
        if (!transactionPlan.Supported)
            throw UnsupportedSourceSlideDelete(removed[0], transactionPlan.BlockedReason);

        foreach (var (source, _) in plans)
        {
            presentationPart.DeletePart(source.Part);
        }
        changedParts.UnionWith(transactionPlan.RemovedPackagePartPaths);
        removedSourcePartPaths.UnionWith(transactionPlan.RemovedPackagePartPaths);
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
    }

    private static CodecException UnsupportedSourceSlideDelete(PptxSourceSlideEntry source, string reason) =>
        new(
            "unsupported_presentation_slide_delete",
            $"Source-preserving PPTX deletion requires an exclusively owned OPC descendant closure; slide {source.Index + 1} cannot be deleted because {reason}.",
            PartPath(source.Part));

    private static void CloneRequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        PptxAssetCatalog assetCatalog,
        PptxCustomShowCatalog customShowCatalog,
        PptxNativeObjectCatalog nativeObjects,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths,
        ISet<string> clonedPackageEntryPaths,
        IDictionary<string, string> clonedPartSourcePaths)
    {
        var cloneTargets = targets.Where(target => target.IsClone).ToArray();
        if (cloneTargets.Length == 0) return;
        var retainedSlideParts = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var slideIdList = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var nextSlideId = slideIdList.Elements<P.SlideId>()
            .Select(slideId => slideId.Id?.Value ?? 255U)
            .DefaultIfEmpty(255U)
            .Max();

        foreach (var target in cloneTargets)
        {
            AssertSourceSlideRequestUnchanged(presentationPart, target, layoutIdByPartPath, slideIdByPartPath, assetCatalog, customShowCatalog, nativeObjects);
            var sourcePart = target.Source.Part;
            var omittedShapeTreeIndices = target.Target.ElementDeletions
                .Select(deletion => deletion.Source?.ShapeTreeIndex)
                .Where(index => index is not null)
                .Select(index => checked((int)index!.Value))
                .ToHashSet();
            var result = PptxSlideCloneCodec.Clone(presentationPart, target.Source, retainedSlideParts, omittedShapeTreeIndices);
            var clonePart = result.Part;
            // Validate the complete graph before applying any authorized
            // component projection. The clone codec proves that every
            // relationship and descendant is an exact source copy; the
            // bounded deletion pass below is the only permitted difference.
            PptxSlideCloneCodec.Validate(target.Source, clonePart, retainedSlideParts);
            changedParts.UnionWith(result.ChangedPackagePaths);
            clonedPackageEntryPaths.UnionWith(result.ChangedPackagePaths);
            addedPartPaths.UnionWith(result.AddedOpaquePartPaths);
            addedRelationshipIds.UnionWith(result.AddedOpaqueRelationshipKeys);
            foreach (var (clonePath, sourcePath) in result.CopiedPartSourcePaths)
                clonedPartSourcePaths.Add(clonePath, sourcePath);
            if (nextSlideId == uint.MaxValue)
                throw new CodecException("presentation_slide_id_exhausted", "PPTX cannot allocate another 32-bit slide identifier.", "ppt/presentation.xml");
            nextSlideId++;
            target.OutputPart = clonePart;
            target.OutputSlideId = new P.SlideId
            {
                Id = nextSlideId,
                RelationshipId = presentationPart.GetIdOfPart(clonePart),
            };
            ApplyCloneElementDeletions(
                target,
                clonePart,
                changedParts,
                addedRelationshipIds,
                addedPartPaths,
                clonedPackageEntryPaths,
                clonedPartSourcePaths);
        }
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
        clonedPackageEntryPaths.Add("[Content_Types].xml");
    }

    private static void ApplyCloneElementDeletions(
        PptxTargetSlideEntry target,
        SlidePart clonePart,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths,
        ISet<string> clonedPackageEntryPaths,
        IDictionary<string, string> clonedPartSourcePaths)
    {
        if (target.Target.ElementDeletions.Count == 0) return;
        var sourceRoot = target.Source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {target.Source.Index + 1} has no slide root.", PartPath(target.Source.Part));
        var cloneRoot = clonePart.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation cloned slide {target.TargetIndex + 1} has no slide root.", PartPath(clonePart));
        var sourceElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {target.Source.Index + 1} has no shape tree.", PartPath(target.Source.Part)));
        var cloneElements = ShapeElements(cloneRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation cloned slide {target.TargetIndex + 1} has no shape tree.", PartPath(clonePart)));
        if (sourceElements.Length != cloneElements.Length)
            throw PptxSlideCloneCodec.Unsupported(target.Source, "the cloned shape tree changed before the authorized component projection");

        var pending = new List<(int Index, OpenXmlElement Source, OpenXmlElement Clone, PptxElementDeletionPlan Plan)>();
        foreach (var deletion in target.Target.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", $"Presentation cloned slide {target.TargetIndex + 1} deletion {deletion.Id} is missing its source binding.", PartPath(target.Source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length)
                throw PptxSlideCloneCodec.Unsupported(target.Source, "an authorized component deletion identifies an invalid source shape-tree index");
            var sourceElement = sourceElements[sourceIndex];
            var cloneElement = cloneElements[sourceIndex];
            var sourcePlan = PptxElementDeletionCodec.Analyze(target.Source.Part, sourceElement, sourceElements, allowDuplicateNativeIds: true);
            var clonePlan = PptxElementDeletionCodec.Analyze(clonePart, cloneElement, cloneElements, allowDuplicateNativeIds: true);
            if (!sourcePlan.Supported || !clonePlan.Supported)
                throw new CodecException("unsupported_presentation_element_delete", $"Presentation cloned slide {target.TargetIndex + 1} component deletion {deletion.Id} is not supported by the source or clone deletion proof.", PartPath(clonePart));
            pending.Add((sourceIndex, sourceElement, cloneElement, clonePlan));
        }
        foreach (var deletion in pending.OrderByDescending(item => item.Index))
        {
            PptxElementDeletionCodec.Apply(clonePart, deletion.Clone, deletion.Plan);
            changedParts.Add(PartPath(clonePart));
            if (deletion.Plan.RelationshipIds.Count > 0)
            {
                changedParts.Add(RelationshipPartPath(clonePart));
                foreach (var relationshipId in deletion.Plan.RelationshipIds)
                    addedRelationshipIds.Remove($"{PartPath(clonePart)}\0{relationshipId}");
            }
            if (deletion.Plan.RemovedPackagePartPaths.Count > 0)
            {
                changedParts.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                addedPartPaths.ExceptWith(deletion.Plan.RemovedPackagePartPaths);
                clonedPackageEntryPaths.ExceptWith(deletion.Plan.RemovedPackagePartPaths);
                foreach (var path in deletion.Plan.RemovedPackagePartPaths)
                    clonedPartSourcePaths.Remove(path);
                changedParts.Add("[Content_Types].xml");
            }
        }
        cloneRoot.Save();
    }

    // The first clone profile preserves the origin SlidePart byte-for-byte.
    // Allowing its removal in the same transaction would turn this leaf clone
    // into a topology-replacement primitive before the new graph has crossed
    // an export/reimport boundary, so reject it before any package mutation.
    private static void AssertCloneOriginsRetained(IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var retainedSources = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        foreach (var clone in targets.Where(target => target.IsClone))
        {
            if (!retainedSources.Contains(clone.Source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    $"Presentation clone {clone.TargetIndex + 1} cannot remove its origin slide in the same source-preserving export. Export and import the unchanged clone before changing source topology.",
                    PartPath(clone.Source.Part));
        }
    }

    private static void AssertSourceSlideRequestUnchanged(
        PresentationPart presentationPart,
        PptxTargetSlideEntry target,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        PptxAssetCatalog assetCatalog,
        PptxCustomShowCatalog customShowCatalog,
        PptxNativeObjectCatalog nativeObjects)
    {
        var source = target.Source;
        var root = source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
        var common = root.CommonSlideData ??
            throw new CodecException("missing_common_slide_data", $"Presentation source slide {source.Index + 1} has no common slide data.", PartPath(source.Part));
        var tree = common.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {source.Index + 1} has no shape tree.", PartPath(source.Part));
        if (target.Target.Name != (common.Name?.Value ?? string.Empty))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source name");
        var layoutPart = source.Part.SlideLayoutPart ??
            throw PptxSlideCloneCodec.Unsupported(source, "it does not have a resolvable layout relationship");
        var expectedLayoutId = layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart));
        if (string.IsNullOrWhiteSpace(expectedLayoutId) || target.Target.LayoutId != expectedLayoutId)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source layout binding");
        var sourceVisibility = PptxSlideVisibilityCodec.Read(root);
        if (sourceVisibility.Editable != target.Target.HasHidden ||
            target.Target.HasHidden && sourceVisibility.Hidden != target.Target.Hidden)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes or invents its source visibility");
        var sourceBinding = target.Target.CloneSource ??
            throw new CodecException("missing_presentation_slide_clone_binding", $"Presentation clone {target.TargetIndex + 1} is missing clone_source.", PartPath(source.Part));
        if (sourceBinding.LayoutRelationshipId != source.Part.GetIdOfPart(layoutPart) ||
            !sourceBinding.BackgroundSemanticSha256.Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common)), StringComparison.OrdinalIgnoreCase) ||
            !sourceBinding.TransitionSemanticSha256.Equals(PptxTransitionCodec.SemanticHash(PptxTransitionCodec.Read(root)), StringComparison.OrdinalIgnoreCase) ||
            sourceBinding.TransitionEditable != PptxTransitionCodec.Supports(root) ||
            sourceBinding.TransitionPresent != PptxTransitionCodec.HasTransition(root) ||
            sourceBinding.TransitionAddable != PptxTransitionCodec.CanAdd(root) ||
            sourceBinding.VisibilityEditable != sourceVisibility.Editable ||
            !sourceBinding.VisibilitySemanticSha256.Equals(sourceVisibility.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_slide_clone_binding_mismatch", $"Presentation clone {target.TargetIndex + 1} does not match its source layout/background/transition binding.", PartPath(source.Part));
        if (!BackgroundSemanticHash(target.Target.Background).Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common)), StringComparison.OrdinalIgnoreCase))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source background");
        if (!PptxTransitionCodec.SemanticHash(target.Target.Transition).Equals(PptxTransitionCodec.SemanticHash(PptxTransitionCodec.Read(root)), StringComparison.OrdinalIgnoreCase))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source transition");
        if (!PptxSpeakerNotesCodec.Equivalent(target.Target.SpeakerNotes, PptxSpeakerNotesCodec.Read(source.Part)))
            throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} speaker notes are not unchanged source notes.", PartPath(source.Part));
        var legacyProfile = PptxLegacyCommentsCodec.Profile(presentationPart, source.Part, source.Index);
        if (legacyProfile.Supported && !PptxLegacyCommentsCodec.Equivalent(legacyProfile.Comments, target.Target.LegacyComments) ||
            !legacyProfile.Supported && target.Target.LegacyComments.Count > 0)
            throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} legacy comments are not unchanged source comments.", PartPath(source.Part));

        var sourceElements = ShapeElements(tree);
        var zOrderPlan = AnalyzeElementZOrder(sourceElements);
        if (sourceElements.Length != target.Target.Elements.Count + target.Target.ElementDeletions.Count)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone does not account for every source element");
        var context = new PptxPartContext(source.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
        var elementIdsByNativeId = NativeElementIds(sourceElements, $"presentation/slide/{source.Index + 1}");
        var requestedBySourceIndex = new Dictionary<int, PresentationElement>();
        var previousSourceIndex = -1;
        foreach (var requested in target.Target.Elements)
        {
            var binding = requested.Source ??
                throw new CodecException("missing_presentation_element_binding", $"Presentation clone {target.TargetIndex + 1} element {requested.Id} is missing its source binding.", PartPath(source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length || sourceIndex <= previousSourceIndex || !requestedBySourceIndex.TryAdd(sourceIndex, requested))
                throw PptxSlideCloneCodec.Unsupported(source, "retained clone elements do not preserve unique source shape-tree order");
            previousSourceIndex = sourceIndex;
        }
        var deletionsBySourceIndex = new Dictionary<int, PresentationElementDeletion>();
        foreach (var deletion in target.Target.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", $"Presentation clone {target.TargetIndex + 1} deletion {deletion.Id} is missing its source binding.", PartPath(source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length || requestedBySourceIndex.ContainsKey(sourceIndex) || !deletionsBySourceIndex.TryAdd(sourceIndex, deletion))
                throw PptxSlideCloneCodec.Unsupported(source, "component deletions do not identify unique omitted source elements");
        }
        var retainedNativeIds = sourceElements
            .Select((element, index) => (element, index))
            .Where(item => !deletionsBySourceIndex.ContainsKey(item.index))
            .SelectMany(item => PptxElementDeletionCodec.NativeIds(item.element))
            .ToArray();
        if (retainedNativeIds.GroupBy(id => id).Any(group => group.Count() > 1))
            throw PptxSlideCloneCodec.Unsupported(source, "retained component elements contain duplicate native drawing IDs");
        var retainedNativeIdSet = retainedNativeIds.ToHashSet();
        for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
        {
            var original = ReadElement(sourceElements[elementIndex], source.Index, elementIndex, context, nativeObjects, elementIdsByNativeId);
            var strictDeletionPlan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[elementIndex], sourceElements);
            SetElementDeletionCapability(original, strictDeletionPlan);
            SetElementZOrderCapability(original, zOrderPlan);
            if (requestedBySourceIndex.TryGetValue(elementIndex, out var requested))
            {
                var binding = requested.Source!;
                AssertElementBinding(requested.Id, binding, sourceElements[elementIndex], original, strictDeletionPlan, zOrderPlan, target.TargetIndex, elementIndex, source.Part);
                if (!SemanticHash(requested).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} is not an unchanged source element.", PartPath(source.Part));
                continue;
            }
            if (!deletionsBySourceIndex.TryGetValue(elementIndex, out var deletion))
                throw PptxSlideCloneCodec.Unsupported(source, "a source element is neither retained nor explicitly deleted");
            AssertElementBinding(deletion.Id, deletion.Source!, sourceElements[elementIndex], original, strictDeletionPlan, zOrderPlan, target.TargetIndex, elementIndex, source.Part);
            if (!deletion.Id.Equals(original.Id, StringComparison.Ordinal))
                throw new CodecException("presentation_element_deletion_binding_mismatch", $"Presentation clone {target.TargetIndex + 1} deletion {elementIndex + 1} changed its source element identity.", PartPath(source.Part));
            var projectionDeletionPlan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[elementIndex], sourceElements, allowDuplicateNativeIds: true);
            if (!projectionDeletionPlan.Supported)
                throw new CodecException("unsupported_presentation_element_delete", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} cannot be safely deleted: {projectionDeletionPlan.BlockedReason}.", PartPath(source.Part));
            if (PptxElementDeletionCodec.NativeIds(sourceElements[elementIndex]).Overlaps(retainedNativeIdSet))
                throw PptxSlideCloneCodec.Unsupported(source, "a deleted element shares a native drawing ID with a retained component element");
        }
    }

    private static bool ReorderSourceSlideIdList(PresentationPart presentationPart, IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var list = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var sourceIds = list.Elements<P.SlideId>().ToArray();
        var requestedRelationshipIds = targets.Select(target => target.OutputSlideId.RelationshipId?.Value ?? string.Empty).ToArray();
        if (sourceIds.Select(slideId => slideId.RelationshipId?.Value ?? string.Empty).SequenceEqual(requestedRelationshipIds, StringComparer.Ordinal)) return false;

        // Keep any future extension child in its original relative position;
        // only replace the ordered p:sldId entries with exact source clones.
        var firstNonSlideId = list.ChildElements.FirstOrDefault(item => item is not P.SlideId);
        var reordered = targets.Select(target => (P.SlideId)target.OutputSlideId.CloneNode(true)).ToArray();
        foreach (var sourceId in sourceIds) sourceId.Remove();
        foreach (var slideId in reordered)
        {
            if (firstNonSlideId is null) list.Append(slideId);
            else list.InsertBefore(slideId, firstNonSlideId);
        }
        root.Save();
        return true;
    }

    private static string MasterTextStylesSemanticHash(PresentationMasterTextStyles? source)
    {
        var semantic = source?.Clone() ?? new PresentationMasterTextStyles();
        PptxMasterTextStylesCodec.NormalizeSemantics(semantic);
        return Hash(semantic.ToByteArray());
    }

    private static string BackgroundSemanticHash(PresentationBackground? source) =>
        Hash((source ?? new PresentationBackground()).ToByteArray());

    private static bool ApplyPlaceholders(
        P.ShapeTree shapeTree,
        string ownerId,
        IList<PresentationPlaceholder> requested,
        PptxPartContext partContext,
        string partPath)
    {
        var originals = PptxPlaceholderCodec.Read(shapeTree, ownerId, partContext);
        if (originals.Count != requested.Count)
            throw new CodecException(
                "presentation_placeholder_topology_changed",
                $"Source-preserving PPTX export requires {ownerId}'s original {originals.Count}-placeholder topology; the artifact contains {requested.Count} placeholders.",
                partPath);
        var changed = false;
        for (var index = 0; index < originals.Count; index++)
        {
            var original = originals[index];
            var target = requested[index];
            var sourceBinding = original.Source!;
            var binding = target.Source ?? throw new CodecException(
                "missing_presentation_placeholder_binding",
                $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.",
                partPath);
            var sourceShape = PptxPlaceholderCodec.BoundShape(shapeTree, original);
            if (sourceShape is null || target.Id != original.Id ||
                binding.ShapeTreeIndex != sourceBinding.ShapeTreeIndex ||
                !binding.ElementSha256.Equals(sourceBinding.ElementSha256, StringComparison.OrdinalIgnoreCase) ||
                !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != sourceBinding.Editable ||
                binding.DirectFramePresenceEditable != sourceBinding.DirectFramePresenceEditable ||
                binding.TextEditable != sourceBinding.TextEditable ||
                binding.AccessibilityEditable != sourceBinding.AccessibilityEditable ||
                !binding.ElementSha256.Equals(PptxPlaceholderCodec.ElementHash(sourceShape), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_binding_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its hash-bound source element.",
                    partPath);
            if (!PptxPlaceholderCodec.SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_source_semantics_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its source semantic binding.",
                    partPath);
            if (PptxPlaceholderCodec.SemanticHash(target).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation placeholder {index + 1} under {ownerId} has no safely editable semantic component in this codec slice.",
                    partPath);
            PptxPlaceholderCodec.Apply(sourceShape, original, target, partContext);
            changed = true;
        }
        return changed;
    }

    private static IEnumerable<PresentationTextParagraph> MasterStyleParagraphs(PresentationMasterTextStyles? source) =>
        source is null
            ? []
            : source.TitleLevels.Concat(source.BodyLevels).Concat(source.OtherLevels);

    private static void TrackContextChanges(
        OpenXmlPart owner,
        PptxPartContext context,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths)
    {
        if (context.RelationshipsChanged)
        {
            changedParts.Add(RelationshipPartPath(owner));
            foreach (var id in context.AddedRelationshipIds)
                addedRelationshipIds.Add($"{PartPath(owner)}\0{id}");
        }
        foreach (var path in context.AddedPartPaths)
        {
            changedParts.Add(path);
            addedPartPaths.Add(path);
            changedParts.Add("[Content_Types].xml");
        }
    }

    private static string ShapeResidualHash(P.Shape source, PptxPartContext slideContext)
    {
        var shape = (P.Shape)source.CloneNode(true);
        PptxNonVisualAccessibilityCodec.ScrubModeledContent(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties) drawingProperties.TextBox = null;
        if (shape.ShapeProperties is { } properties)
        {
            if (properties.Transform2D is { } transform)
            {
                if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
                if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
                PptxShapeTransformCodec.Scrub(transform);
            }
            properties.GetFirstChild<A.CustomGeometry>()?.Remove();
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry) geometry.Preset = A.ShapeTypeValues.Rectangle;
            else properties.InsertAfter(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }, properties.GetFirstChild<A.Transform2D>());
            properties.GetFirstChild<A.EffectList>()?.Remove();
            foreach (var fill in properties.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray()) fill.Remove();
            if (properties.GetFirstChild<A.Outline>() is { } outline)
                PptxLineStyleCodec.ScrubModeledContent(outline);
        }
        PptxTextCodec.ScrubModeledContent(shape.TextBody, slideContext);
        return HashElement(shape);
    }

    private static string ConnectorResidualHash(P.ConnectionShape source)
    {
        var connector = (P.ConnectionShape)source.CloneNode(true);
        if (connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties is { } drawingProperties)
            drawingProperties.RemoveAllChildren();
        connector.ShapeProperties?.RemoveAllChildren();
        return HashElement(connector);
    }

    private static string ChartFrameResidualHash(P.GraphicFrame source)
    {
        var chart = (P.GraphicFrame)source.CloneNode(true);
        PptxChartCodec.ScrubFrame(chart);
        return HashElement(chart);
    }

    private static string PictureResidualHash(P.Picture source)
    {
        var picture = (P.Picture)source.CloneNode(true);
        PptxPictureCodec.ScrubModeledContent(picture);
        return HashElement(picture);
    }

    private static string TableResidualHash(P.GraphicFrame source)
    {
        var table = (P.GraphicFrame)source.CloneNode(true);
        PptxTableCodec.ScrubModeledContent(table);
        return HashElement(table);
    }

    private static string NativeObjectResidualHash(OpenXmlElement source)
    {
        var clone = source.CloneNode(true);
        if (clone.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault() is { } nonVisual)
            nonVisual.Name = string.Empty;
        if (clone is P.Picture picture && picture.ShapeProperties?.GetFirstChild<A.Transform2D>() is { } pictureTransform)
        {
            ScrubFrame(pictureTransform);
        }
        else if (clone is P.ConnectionShape connector && connector.ShapeProperties?.GetFirstChild<A.Transform2D>() is { } connectorTransform)
        {
            ScrubFrame(connectorTransform);
        }
        else if (clone is P.GraphicFrame graphicFrame && graphicFrame.Transform is { } transform)
        {
            ScrubFrame(transform);
            if (PptxNativeObjectCatalog.Classify(clone) == "oleObject" && graphicFrame.Descendants<A.Transform2D>().FirstOrDefault() is { } preview)
                ScrubFrame(preview);
        }
        else if (clone is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { } groupTransform)
        {
            ScrubFrame(groupTransform);
        }
        return HashElement(clone);
    }

    private static void ScrubFrame(P.Transform transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
    }

    private static void ScrubFrame(A.Transform2D transform)
    {
        if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
        if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
    }

    private static void ScrubFrame(A.TransformGroup transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
    }

    private static string MasterResidualHash(P.SlideMaster source, PptxPartContext partContext)
    {
        var master = (P.SlideMaster)source.CloneNode(true);
        PptxMasterTextStylesCodec.ScrubModeledContent(master, partContext);
        PptxBackgroundCodec.ScrubModeledContent(master.CommonSlideData);
        PptxPlaceholderCodec.ScrubModeledContent(master.CommonSlideData?.ShapeTree, partContext);
        return HashElement(master);
    }

    private static string LayoutResidualHash(P.SlideLayout source, PptxPartContext partContext)
    {
        var layout = (P.SlideLayout)source.CloneNode(true);
        PptxBackgroundCodec.ScrubModeledContent(layout.CommonSlideData);
        PptxPlaceholderCodec.ScrubModeledContent(layout.CommonSlideData?.ShapeTree, partContext);
        return HashElement(layout);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }

    private static Dictionary<string, string> PackagePartHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.Where(entry => !entry.FullName.EndsWith('/')).ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var source = entry.Open();
                using var copy = new MemoryStream();
                source.CopyTo(copy);
                return Hash(copy.ToArray());
            },
            StringComparer.OrdinalIgnoreCase);
    }

    private static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated PPTX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    private const int MaxOffice2021ValidationErrors = 256;

    private static void ValidateOffice2021(byte[] bytes)
    {
        var errors = Office2021ValidationErrors(bytes);
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Take(8).Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Generated PPTX is not valid Office 2021 Open XML: {detail}");
    }

    // An imported package may use valid-in-host extension markup that the
    // Open XML SDK's Office 2021 validator does not recognize. Source-bound
    // export may retain such diagnostics, but must never introduce a new one.
    // This comparison is intentionally unavailable to source-free export.
    private static int ValidateOffice2021AgainstSource(
        byte[] sourceBytes,
        byte[] outputBytes,
        IReadOnlyDictionary<string, string>? clonedPartSourcePaths = null)
    {
        var sourceErrors = Office2021ValidationErrors(sourceBytes);
        var sourceSignatures = sourceErrors
            .Select(error => error.Signature)
            .ToHashSet(StringComparer.Ordinal);
        var exactCopiedPartSources = ExactCopiedPartSources(sourceBytes, outputBytes, clonedPartSourcePaths);
        var introduced = Office2021ValidationErrors(outputBytes)
            .Where(error =>
                !sourceSignatures.Contains(error.Signature) &&
                (!exactCopiedPartSources.TryGetValue(error.Part, out var sourcePart) ||
                 !sourceSignatures.Contains(error.SignatureForPart(sourcePart))))
            .Take(8)
            .ToArray();
        if (introduced.Length == 0) return sourceErrors.Length;
        var detail = string.Join("; ", introduced.Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Source-preserving PPTX export introduced Office 2021 Open XML validation error(s): {detail}");
    }

    private sealed record Office2021ValidationError(
        string Part,
        string Path,
        string Id,
        string Description)
    {
        internal string Signature => SignatureForPart(Part);
        internal string Detail => $"{(Path.Length == 0 ? Part : Path)}: {Description}";
        internal string SignatureForPart(string part) => string.Join("\u001f", part, Path, Id, Description);
    }

    private static Office2021ValidationError[] Office2021ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021)
            .Validate(package)
            .Take(MaxOffice2021ValidationErrors + 1)
            .Select(error =>
            {
                var part = error.Part?.Uri.ToString().TrimStart('/') ?? "package";
                var path = error.Path?.XPath ?? string.Empty;
                return new Office2021ValidationError(part, path, error.Id ?? string.Empty, error.Description ?? string.Empty);
            })
            .ToArray();
        if (errors.Length <= MaxOffice2021ValidationErrors) return errors;
        throw new CodecException(
            "openxml_validation_budget_exceeded",
            $"PPTX validation produced more than {MaxOffice2021ValidationErrors} errors; refusing an unbounded validation result.");
    }

    private static IReadOnlyDictionary<string, string> ExactCopiedPartSources(
        byte[] sourceBytes,
        byte[] outputBytes,
        IReadOnlyDictionary<string, string>? clonedPartSourcePaths)
    {
        if (clonedPartSourcePaths is null || clonedPartSourcePaths.Count == 0)
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sourceHashes = PackagePartHashes(sourceBytes);
        var outputHashes = PackagePartHashes(outputBytes);
        return clonedPartSourcePaths
            .Where(pair =>
                sourceHashes.TryGetValue(pair.Value, out var sourceHash) &&
                outputHashes.TryGetValue(pair.Key, out var outputHash) &&
                sourceHash.Equals(outputHash, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
    }

}
