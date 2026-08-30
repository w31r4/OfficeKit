using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

/// <summary>
/// Selects the authored or source-bound PPJ lowering path after one language
/// validation. Source-bound programs always compile against the exact PPTX
/// bytes named by their source descriptor; PPJ never turns a stale descriptor
/// into write authority.
/// </summary>
internal static class PpjPresentationCompiler
{
    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        byte[] sourceBytes,
        EffectiveCodecLimits limits)
    {
        var validation = PpjProgramValidator.Validate(request.ProgramJson.Span);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        if (validation.Program!.Source is null)
        {
            if (sourceBytes.Length != 0)
                throw new CodecException(
                    "ppj.unexpectedSource",
                    "A source-free PPJ compile cannot attach a PPTX source package.",
                    "$.source");
            if (request.ValidationOnly)
                return PpjAuthoredPresentationCompiler.ValidateOnly(request, validation);
            return PpjAuthoredPresentationCompiler.Compile(request, limits);
        }

        return PpjSourceBoundPresentationCompiler.Compile(request, sourceBytes, limits, validation);
    }
}

/// <summary>
/// Reprojects an exact source package, compares it with the requested PPJ by
/// stable IDs, and maps only capability-issued semantic differences onto the
/// existing source-preserving Presentation IR. The mature PPTX writer remains
/// the independent package-level oracle; this class never writes OOXML.
/// </summary>
internal static class PpjSourceBoundPresentationCompiler
{
    private const double EmuPerPoint = 12_700d;
    private sealed class MutationState
    {
        internal bool SemanticChanges { get; set; }
        internal List<NativeLeafMutation> NativeLeaves { get; } = [];
    }

    private sealed record NativeLeafMutation(
        string ProgramElementId,
        PresentationSlide Slide,
        PresentationElement Element,
        IReadOnlyList<uint> ShapeTreePath,
        uint NativeLeafIndex,
        uint TextLeafIndex,
        string Before,
        string After,
        string LeafKind);

    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        byte[] sourceBytes,
        EffectiveCodecLimits limits,
        PpjValidationResult validation)
    {
        if (sourceBytes.Length == 0)
            throw new CodecException(
                "ppj.source.missing",
                "A source-bound PPJ compile requires the exact source PPTX bytes.",
                "$.source");

        var requested = validation.Program!;
        var source = requested.Source!;
        var sourceSha256 = Sha256(sourceBytes);
        if (!source.Kind.Equals("pptx", StringComparison.Ordinal) ||
            !sourceSha256.Equals(source.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.source.hashMismatch",
                "The source-bound PPJ does not match the exact PPTX input bytes.",
                "$.source.sha256");

        var projected = PpjPresentationProjector.Project(
            sourceBytes,
            new PresentationProgramRequest
            {
                SourceUri = source.Uri,
                AssetRootUri = AssetRoot(requested),
                IncludeNodeMap = true,
            },
            limits);
        var baselineValidation = PpjProgramValidator.Validate(projected.Program.ProgramJson.Span);
        if (!baselineValidation.IsValid)
            throw new CodecException(
                "ppj.projection.invalid",
                "The exact PPTX source could not be reprojected into a valid PPJ baseline.",
                "$.source");
        var baseline = baselineValidation.Program!;

        RequireExactSourceDescriptor(baseline, requested);
        RequireBaselineAssets(baseline, requested);
        RequireRootTopology(baseline, requested);

        var artifact = projected.SourceArtifact ??
            throw new CodecException("ppj.source.projection", "Source-bound projection did not return its native source artifact.", "$.source");
        var presentation = artifact.Presentation ??
            throw new CodecException("ppj.source.presentation", "The exact source did not import as a Presentation artifact.", "$.source");
        var assetIds = BuildAssetCatalog(baseline, requested, projected, request.Assets, artifact);
        var changedNodeIds = new HashSet<string>(StringComparer.Ordinal);
        var mutations = new MutationState();
        var physicalChanges = ApplyPages(
            baseline,
            requested,
            presentation,
            assetIds,
            projected.NativeLeafBindings,
            changedNodeIds,
            mutations);

        if (request.ValidationOnly)
        {
            var validationReceipt = new PresentationProgramResult
            {
                ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
                ProgramSha256 = validation.ProgramSha256,
                NodeMapJson = request.IncludeNodeMap ? projected.Program.NodeMapJson : ByteString.Empty,
                SourceSha256 = sourceSha256,
                SourceBound = true,
                ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
            };
            validationReceipt.Assets.Add(projected.Program.Assets.Select(asset => asset.Clone()));
            validationReceipt.ChangedNodeIds.Add(changedNodeIds.OrderBy(id => id, StringComparer.Ordinal));
            return new([], validationReceipt, projected.Diagnostics);
        }

        byte[] output;
        IReadOnlyList<Diagnostic> diagnostics;
        if (!physicalChanges)
        {
            // A physical no-op is the original byte array, not a reserialized
            // equivalent package. PPJ-only intent/design metadata may still
            // advance outside the third-party PPTX.
            output = sourceBytes;
            diagnostics = projected.Diagnostics;
        }
        else if (mutations.NativeLeaves.Count > 0)
        {
            if (mutations.SemanticChanges)
                throw Unsupported("$.pages", "mixing precise native-leaf edits with other semantic edits in one source-bound build");
            var edited = PptxEditPlanCodec.Apply(sourceBytes, NativeLeafEditPlan(sourceSha256, mutations.NativeLeaves), limits);
            output = edited.File;
            diagnostics = projected.Diagnostics.Concat(edited.Diagnostics).ToArray();
        }
        else
        {
            var exported = PptxCodec.Export(artifact, limits);
            output = exported.File;
            diagnostics = projected.Diagnostics.Concat(exported.Diagnostics).ToArray();
        }

        var receipt = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? projected.Program.NodeMapJson : ByteString.Empty,
            SourceSha256 = sourceSha256,
            OutputSha256 = Sha256(output),
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        receipt.Assets.Add(projected.Program.Assets.Select(asset => asset.Clone()));
        receipt.ChangedNodeIds.Add(changedNodeIds.OrderBy(id => id, StringComparer.Ordinal));
        receipt.ChangedParts.Add(ChangedParts(sourceBytes, output));
        return new(output, receipt, diagnostics);
    }

    private static void RequireExactSourceDescriptor(PpjProgramModel baseline, PpjProgramModel requested)
    {
        if (baseline.Source is null || requested.Source is null ||
            !JsonEqual(baseline.Root.GetProperty("source"), requested.Root.GetProperty("source")))
            throw new CodecException(
                "ppj.source.staleProjection",
                "The PPJ source descriptor no longer matches a fresh projection of the exact PPTX.",
                "$.source");
    }

    private static void RequireBaselineAssets(PpjProgramModel baseline, PpjProgramModel requested)
    {
        var requestedById = requested.Assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        foreach (var source in baseline.Assets)
        {
            if (!requestedById.TryGetValue(source.Id, out var current) ||
                !source.Uri.Equals(current.Uri, StringComparison.Ordinal) ||
                !source.MimeType.Equals(current.MimeType, StringComparison.OrdinalIgnoreCase) ||
                !source.Sha256.Equals(current.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "ppj.asset.sourceBindingChanged",
                    $"Source-projected PPJ asset {source.Id} changed its URI, MIME type, or content hash.",
                    "$.assets");
        }
    }

    private static void RequireRootTopology(PpjProgramModel baseline, PpjProgramModel requested)
    {
        if (requested.Components.Count != 0)
            throw Unsupported("$", "source-bound PPJ cannot introduce component definitions in this compiler slice");
        RequirePropertyEqual(baseline.Root, requested.Root, "sections", "$.sections");
        RequirePropertyEqual(baseline.Root, requested.Root, "customShows", "$.customShows");
        RequirePropertyEqual(baseline.Root, requested.Root, "comments", "$.comments");
    }

    private static Dictionary<string, string> BuildAssetCatalog(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PpjProjectionResult projected,
        IEnumerable<Asset> supplied,
        ArtifactEnvelope artifact)
    {
        var declared = requested.Assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        var suppliedById = supplied.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        foreach (var id in suppliedById.Keys.Where(id => !declared.ContainsKey(id)))
            throw new CodecException("ppj.asset.undeclared", $"PPJ compile received undeclared asset {id}.", "$.assets");

        var output = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var declaration in baseline.Assets)
        {
            var native = artifact.Assets.FirstOrDefault(asset =>
                AssetHash(asset).Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) &&
                asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase));
            if (native is null)
                throw new CodecException("ppj.asset.sourceMissing", $"Source-projected asset {declaration.Id} no longer resolves in the exact PPTX.", "$.assets");
            output[declaration.Id] = native.Id;
        }

        var baselineIds = baseline.Assets.Select(asset => asset.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var declaration in requested.Assets.Where(asset => !baselineIds.Contains(asset.Id)))
        {
            if (!suppliedById.TryGetValue(declaration.Id, out var asset) || asset.Data.IsEmpty)
                throw new CodecException("ppj.asset.missing", $"PPJ asset {declaration.Id} has no supplied bytes.", "$.assets");
            var hash = Sha256(asset.Data.Span);
            if (!hash.Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) ||
                !hash.Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.hashMismatch", $"PPJ asset {declaration.Id} does not match its declared SHA-256.", "$.assets");
            if (!asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.mimeMismatch", $"PPJ asset {declaration.Id} does not match its declared MIME type.", "$.assets");
            var nativeId = $"asset/ppj/{hash}";
            if (!artifact.Assets.Any(item => item.Id.Equals(nativeId, StringComparison.Ordinal)))
            {
                var copy = asset.Clone();
                copy.Id = nativeId;
                copy.FileName = declaration.Uri;
                copy.Sha256 = hash;
                artifact.Assets.Add(copy);
            }
            output[declaration.Id] = nativeId;
        }
        return output;
    }

    private static IReadOnlyDictionary<string, (double Width, double Height)> AssetDimensions(JsonElement root)
    {
        var output = new Dictionary<string, (double Width, double Height)>(StringComparer.Ordinal);
        if (!root.TryGetProperty("assets", out var assets)) return output;
        foreach (var asset in assets.EnumerateArray())
        {
            if (!asset.TryGetProperty("widthPx", out var width) || !asset.TryGetProperty("heightPx", out var height)) continue;
            output[asset.GetProperty("id").GetString()!] = (width.GetDouble(), height.GetDouble());
        }
        return output;
    }

    private static string ResolveAsset(IReadOnlyDictionary<string, string> assets, string id, string path) =>
        assets.TryGetValue(id, out var nativeId)
            ? nativeId
            : throw new CodecException("ppj.asset.missing", $"PPJ image asset {id} has no validated bytes.", path);

    private static PresentationImagePaint BuildImagePaint(
        JsonElement fill,
        double frameWidth,
        double frameHeight,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path) => PpjImagePaintLowering.Build(
            fill,
            frameWidth,
            frameHeight,
            id => ResolveAsset(assets, id, path + ".asset"),
            id => assetDimensions.TryGetValue(id, out var dimensions) ? dimensions : null,
            path);

    private static PresentationBackground? BuildBackground(
        JsonElement fill,
        double canvasWidth,
        double canvasHeight,
        Func<string, string> resolveAsset,
        Func<string, (double Width, double Height)?> assetDimensions,
        string path)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return null;
        if (type == "image") return new PresentationBackground
        {
            ImagePaint = PpjImagePaintLowering.Build(fill, canvasWidth, canvasHeight, resolveAsset, assetDimensions, path),
        };
        if (type == "gradient") return new PresentationBackground { GradientFill = BuildGradientFill(fill, path) };
        if (type == "solid")
        {
            if (fill.TryGetProperty("opacity", out var opacity) && opacity.GetDouble() != 1)
                throw Unsupported(path, "translucent solid background");
            return new PresentationBackground { Solid = true, ColorRgb = Rgb(fill.GetProperty("color"), path + ".color") };
        }
        throw Unsupported(path, $"background fill {type}");
    }

    private static bool ApplyPages(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> nativeLeafBindings,
        ISet<string> changedNodeIds,
        MutationState mutations)
    {
        if (baseline.Pages.Count != presentation.Slides.Count || requested.Pages.Count > baseline.Pages.Count)
            throw Unsupported("$.pages", "source-bound page insertion or inconsistent source topology");
        var sourcePages = baseline.Pages.Select((page, index) => new
        {
            Program = page,
            Wire = presentation.Slides[index],
        }).ToDictionary(item => item.Program.Id, StringComparer.Ordinal);
        var requestedPageIds = requested.Pages.Select(page => page.Id).ToArray();
        if (requestedPageIds.Distinct(StringComparer.Ordinal).Count() != requestedPageIds.Length ||
            requestedPageIds.Any(id => !sourcePages.ContainsKey(id)))
            throw Unsupported("$.pages", "source-bound page insertion or identity change");
        var retainedPageSourceOrder = baseline.Pages
            .Where(page => requestedPageIds.Contains(page.Id, StringComparer.Ordinal))
            .Select(page => page.Id);
        if (!retainedPageSourceOrder.SequenceEqual(requestedPageIds, StringComparer.Ordinal))
            throw Unsupported("$.pages", "source-bound page reorder requires a dedicated capability");

        var changed = false;
        var assetDimensions = AssetDimensions(requested.Root);
        var requestedSlides = new List<PresentationSlide>(requested.Pages.Count);
        for (var index = 0; index < requested.Pages.Count; index++)
        {
            var after = requested.Pages[index];
            var sourcePage = sourcePages[after.Id];
            var before = sourcePage.Program;
            var slide = sourcePage.Wire;
            var path = $"$.pages[{index}]";
            RequireNativeRef(before.Raw, after.Raw, path);
            RequireEqualExcept(before.Raw, after.Raw, path, "role", "claim", "background", "elements");
            if (PropertyChanged(before.Raw, after.Raw, "background"))
            {
                RequireCapability(after.NativeRef, "setBackground", path + ".background");
                slide.Background = after.Raw.TryGetProperty("background", out var background)
                    ? BuildBackground(
                        background,
                        requested.Design.Width,
                        requested.Design.Height,
                        id => ResolveAsset(assets, id, path + ".background.asset"),
                        id => assetDimensions.TryGetValue(id, out var dimensions) ? dimensions : null,
                        path + ".background")
                    : null;
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (before.Elements.Count != slide.Elements.Count || after.Elements.Count > before.Elements.Count)
                throw Unsupported(path + ".elements", "source-bound element insertion or inconsistent source topology");
            var sourceElements = before.Elements.Select((element, elementIndex) => new
            {
                Program = element,
                Wire = slide.Elements[elementIndex],
            }).ToDictionary(item => item.Program.Id, StringComparer.Ordinal);
            var requestedIds = after.Elements.Select(element => element.Id).ToArray();
            if (requestedIds.Distinct(StringComparer.Ordinal).Count() != requestedIds.Length ||
                requestedIds.Any(id => !sourceElements.ContainsKey(id)))
                throw Unsupported(path + ".elements", "source-bound element insertion or identity change");

            var requestedWire = new List<PresentationElement>(after.Elements.Count);
            for (var elementIndex = 0; elementIndex < after.Elements.Count; elementIndex++)
            {
                var elementPath = $"{path}.elements[{elementIndex}]";
                var beforeElement = sourceElements[after.Elements[elementIndex].Id];
                var wireElement = beforeElement.Wire;
                var shapeTreePath = new[] { wireElement.Source?.ShapeTreeIndex ?? checked((uint)elementIndex) };
                if (ApplyElement(beforeElement.Program, after.Elements[elementIndex], wireElement, slide, shapeTreePath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, elementPath))
                {
                    changed = true;
                    changedNodeIds.Add(after.Id);
                }
                requestedWire.Add(wireElement);
            }

            foreach (var deleted in before.Elements.Where(element => !requestedIds.Contains(element.Id, StringComparer.Ordinal)))
            {
                RequireCapability(deleted, "delete", path + ".elements");
                var wire = sourceElements[deleted.Id].Wire;
                if (wire.Source?.DeletionCapability?.Supported != true)
                    throw Unsupported(path + ".elements", $"deleting {deleted.Id} without a re-proven native deletion profile");
                slide.ElementDeletions.Add(new PresentationElementDeletion
                {
                    Id = wire.Id,
                    Source = wire.Source.Clone(),
                });
                changedNodeIds.Add(deleted.Id);
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }

            var retainedSourceOrder = before.Elements.Where(element => requestedIds.Contains(element.Id, StringComparer.Ordinal)).Select(element => element.Id).ToArray();
            if (!retainedSourceOrder.SequenceEqual(requestedIds, StringComparer.Ordinal))
            {
                for (var elementIndex = 0; elementIndex < after.Elements.Count; elementIndex++)
                {
                    if (retainedSourceOrder[elementIndex] == requestedIds[elementIndex]) continue;
                    RequireCapability(after.Elements[elementIndex], "reorder", $"{path}.elements[{elementIndex}]");
                    changedNodeIds.Add(after.Elements[elementIndex].Id);
                }
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            slide.Elements.Clear();
            slide.Elements.Add(requestedWire);
            requestedSlides.Add(slide);
        }

        foreach (var deleted in baseline.Pages.Where(page => !requestedPageIds.Contains(page.Id, StringComparer.Ordinal)))
        {
            RequireCapability(deleted.NativeRef, "delete", "$.pages");
            var slide = sourcePages[deleted.Id].Wire;
            if (slide.Source?.DeletionCapability?.Supported != true)
                throw Unsupported("$.pages", $"deleting {deleted.Id} without a re-proven native deletion profile");
            changedNodeIds.Add(deleted.Id);
            mutations.SemanticChanges = true;
            changed = true;
        }
        presentation.Slides.Clear();
        presentation.Slides.Add(requestedSlides);
        return changed;
    }

    private static bool ApplyElement(
        PpjElementModel before,
        PpjElementModel after,
        PresentationElement target,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> nativeLeafBindings,
        ISet<string> changedNodeIds,
        MutationState mutations,
        string path)
    {
        if (!before.Id.Equals(after.Id, StringComparison.Ordinal) || !before.Type.Equals(after.Type, StringComparison.Ordinal))
            throw Unsupported(path, "source-bound element reorder, identity, or type change");
        RequireNativeRef(before.Raw, after.Raw, path);
        var nativeLeafChanged = CollectIssuedNativeLeafMutations(
            before.NativeRef,
            after.NativeRef,
            nativeLeafBindings,
            slide,
            target,
            shapeTreePath,
            after.Id,
            mutations,
            path + ".nativeRef.leaves");

        bool changed;
        switch (before)
        {
            case PpjTextElementModel beforeText when after is PpjTextElementModel afterText && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyTextElement(beforeText, afterText, target, slide, shapeTreePath, assets, assetDimensions, mutations, path);
                break;
            case PpjShapeElementModel beforeShape when after is PpjShapeElementModel afterShape && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyShapeElement(beforeShape, afterShape, target, slide, shapeTreePath, assets, assetDimensions, mutations, path);
                break;
            case PpjPlaceholderElementModel beforePlaceholder when after is PpjPlaceholderElementModel afterPlaceholder && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyPlaceholderElement(beforePlaceholder, afterPlaceholder, target, slide, shapeTreePath, mutations, path);
                break;
            case PpjImageElementModel beforeImage when after is PpjImageElementModel afterImage && target.ContentCase == PresentationElement.ContentOneofCase.Image:
                changed = ApplyImageElement(beforeImage, afterImage, target.Image, assets, assetDimensions, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjChartElementModel beforeChart when after is PpjChartElementModel afterChart && target.ContentCase == PresentationElement.ContentOneofCase.Chart:
                changed = ApplyChartElement(beforeChart, afterChart, target.Chart, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjTableElementModel beforeTable when after is PpjTableElementModel afterTable && target.ContentCase == PresentationElement.ContentOneofCase.Table:
                changed = ApplyTableElement(beforeTable, afterTable, target.Table, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjConnectorElementModel beforeConnector when after is PpjConnectorElementModel afterConnector && target.ContentCase == PresentationElement.ContentOneofCase.Connector:
                changed = ApplyConnectorElement(beforeConnector, afterConnector, target.Connector, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjGroupElementModel beforeGroup when after is PpjGroupElementModel afterGroup && target.ContentCase == PresentationElement.ContentOneofCase.Group:
                changed = ApplyGroupElement(beforeGroup, afterGroup, target.Group, slide, shapeTreePath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, path);
                break;
            case PpjOpaqueElementModel beforeOpaque when after is PpjOpaqueElementModel afterOpaque:
                changed = ApplyOpaqueElement(beforeOpaque, afterOpaque, target, slide, shapeTreePath, mutations, path);
                break;
            default:
                throw Unsupported(path, "the exact source object no longer matches its PPJ projection type");
        }
        changed |= nativeLeafChanged;
        if (changed) changedNodeIds.Add(after.Id);
        return changed;
    }

    private static bool ApplyTextElement(
        PpjTextElementModel before,
        PpjTextElementModel after,
        PresentationElement element,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        MutationState mutations,
        string path)
    {
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "text", "fill", "stroke");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, slide, element, shapeTreePath, mutations, path + ".text");
        }
        semanticChanged |= ApplyFillProperty(before, after, target, "fill", assets, assetDimensions, path);
        semanticChanged |= ApplyStrokeProperty(before, after, target, "stroke", path);
        mutations.SemanticChanges |= semanticChanged;
        changed |= semanticChanged;
        return changed;
    }

    private static bool ApplyShapeElement(
        PpjShapeElementModel before,
        PpjShapeElementModel after,
        PresentationElement element,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        MutationState mutations,
        string path)
    {
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "geometry", "text", "style");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source text body");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, slide, element, shapeTreePath, mutations, path + ".text");
        }
        if (PropertyChanged(before.Raw, after.Raw, "geometry"))
        {
            RequireCapability(after, "setGeometry", path + ".geometry.adjustments");
            var oldGeometry = before.Raw.GetProperty("geometry");
            var newGeometry = after.Raw.GetProperty("geometry");
            RequireEqualExcept(oldGeometry, newGeometry, path + ".geometry", "adjustments");
            target.PresetAdjustments.Clear();
            target.PresetAdjustments.Add(after.GeometryAdjustments);
            semanticChanged = true;
        }
        semanticChanged |= ApplyShapeStyle(before, after, target, assets, assetDimensions, path);
        mutations.SemanticChanges |= semanticChanged;
        changed |= semanticChanged;
        return changed;
    }

    private static bool ApplyPlaceholderElement(
        PpjPlaceholderElementModel before,
        PpjPlaceholderElementModel after,
        PresentationElement element,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        MutationState mutations,
        string path)
    {
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "text");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source placeholder text body");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, slide, element, shapeTreePath, mutations, path + ".text");
        }
        mutations.SemanticChanges |= semanticChanged;
        return changed;
    }

    private static bool ApplyImageElement(
        PpjImageElementModel before,
        PpjImageElementModel after,
        PresentationImage target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "asset", "fit", "crop", "opacity", "mask");
        var changed = ApplyFrame(before, after, target, path);
        if (!before.AssetId.Equals(after.AssetId, StringComparison.Ordinal))
        {
            RequireCapability(after, "replaceImage", path + ".asset");
            if (!assets.TryGetValue(after.AssetId, out var nativeAssetId))
                throw new CodecException("ppj.asset.missing", $"PPJ image asset {after.AssetId} has no validated bytes.", path + ".asset");
            target.AssetId = nativeAssetId;
            changed = true;
        }
        var cropChanged = PropertyChanged(before.Raw, after.Raw, "crop");
        var fitChanged = PropertyChanged(before.Raw, after.Raw, "fit");
        if (cropChanged) RequireCapability(after, "setImageCrop", path + ".crop");
        if (fitChanged) RequireCapability(after, "setImageFit", path + ".fit");
        if (cropChanged || fitChanged)
        {
            var paint = BuildImagePaint(after.Raw, after.Frame.Width, after.Frame.Height, assets, assetDimensions, path);
            target.Crop = paint.Crop;
            target.Tiled = paint.Mode == PresentationImagePaint.Types.Mode.Tile;
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "opacity"))
        {
            RequireCapability(after, "setOpacity", path + ".opacity");
            if (after.Raw.TryGetProperty("opacity", out var opacity)) target.OpacityThousandthPercent = Unit(opacity.GetDouble());
            else target.ClearOpacityThousandthPercent();
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "mask"))
        {
            RequireCapability(after, "setImageMask", path + ".mask.adjustments");
            if (before.MaskKind != "preset" || after.MaskKind != "preset" ||
                before.MaskPreset is null || after.MaskPreset is null ||
                !before.MaskPreset.Equals(after.MaskPreset, StringComparison.Ordinal))
                throw Unsupported(path + ".mask", "source-bound picture mask topology or preset identity change");
            var beforeMask = before.Raw.GetProperty("mask");
            var afterMask = after.Raw.GetProperty("mask");
            RequireEqualExcept(beforeMask, afterMask, path + ".mask", "adjustments");
            target.MaskPresetAdjustments.Clear();
            target.MaskPresetAdjustments.Add(after.MaskAdjustments);
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartElement(PpjChartElementModel before, PpjChartElementModel after, PresentationChart target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path,
            "role", "tags", "frame", "title", "data", "style",
            "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "title"))
        {
            RequireCapability(after, "setChartTitle", path + ".title");
            if (after.Title is not { PlainText: not null })
                throw Unsupported(path + ".title", "rich source chart-title authoring");
            target.Title = after.Title.PlainText;
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "data"))
        {
            ApplyChartData(before, after, target, path + ".data");
            changed = true;
        }
        changed |= ApplyChartStyles(before, after, target, path);
        return changed;
    }

    private static bool ApplyChartStyles(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        string path)
    {
        var changed = ApplyChartStyleTextStyles(before, after, target, path);
        changed |= ApplyChartAxisTextStyle(before, after, target.XAxis, "xAxis", path);
        changed |= ApplyChartAxisTextStyle(before, after, target.YAxis, "yAxis", path);
        changed |= ApplyChartAxisTextStyle(before, after, target.SecondaryXAxis, "secondaryXAxis", path);
        changed |= ApplyChartAxisTextStyle(before, after, target.SecondaryYAxis, "secondaryYAxis", path);
        return changed;
    }

    private static bool ApplyChartStyleTextStyles(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        string path)
    {
        var oldStyle = OptionalProperty(before.Raw, "style");
        var newStyle = OptionalProperty(after.Raw, "style");
        if (JsonEqual(oldStyle, newStyle)) return false;
        RequireOnlyBoundedProperties(
            oldStyle,
            newStyle,
            path + ".style",
            "titleTextStyle",
            "legendTextStyle",
            "dataLabels",
            "chartAreaFill",
            "plotAreaFill");

        var changed = false;
        if (PropertyChanged(oldStyle, newStyle, "titleTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + ".style.titleTextStyle");
            var titleTextStyle = SourceBoundChartTextStyle(newStyle, "titleTextStyle", path + ".style.titleTextStyle");
            if (titleTextStyle is not null && target.Title.Length == 0)
                throw Unsupported(path + ".style.titleTextStyle", "chart-title style without an existing title");
            target.TitleTextStyle = titleTextStyle;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "legendTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + ".style.legendTextStyle");
            var legendTextStyle = SourceBoundChartTextStyle(newStyle, "legendTextStyle", path + ".style.legendTextStyle");
            if (legendTextStyle is not null && !target.HasLegend)
                throw Unsupported(path + ".style.legendTextStyle", "chart-legend style without an existing legend");
            target.LegendTextStyle = legendTextStyle;
            changed = true;
        }

        var oldLabels = oldStyle is { } oldStyleValue ? OptionalProperty(oldStyleValue, "dataLabels") : null;
        var newLabels = newStyle is { } newStyleValue ? OptionalProperty(newStyleValue, "dataLabels") : null;
        if (!JsonEqual(oldLabels, newLabels))
        {
            if (oldLabels is null || newLabels is null || target.DataLabels is null)
                throw Unsupported(path + ".style.dataLabels", "source-bound chart-data-label topology change");
            RequireEqualExcept(oldLabels.Value, newLabels.Value, path + ".style.dataLabels", "textStyle");
            if (PropertyChanged(oldLabels, newLabels, "textStyle"))
            {
                RequireCapability(after, "setChartTextStyle", path + ".style.dataLabels.textStyle");
                target.DataLabels.TextStyle = SourceBoundChartTextStyle(
                    newLabels,
                    "textStyle",
                    path + ".style.dataLabels.textStyle");
                changed = true;
            }
        }
        if (PropertyChanged(oldStyle, newStyle, "chartAreaFill"))
        {
            RequireCapability(after, "setChartFill", path + ".style.chartAreaFill");
            target.ChartAreaFill = SourceBoundChartFill(newStyle, "chartAreaFill", path + ".style.chartAreaFill");
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "plotAreaFill"))
        {
            RequireCapability(after, "setChartFill", path + ".style.plotAreaFill");
            target.PlotAreaFill = SourceBoundChartFill(newStyle, "plotAreaFill", path + ".style.plotAreaFill");
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartAxisTextStyle(
        PpjChartElementModel before,
        PpjChartElementModel after,
        SpreadsheetChartAxisArtifact? target,
        string axisName,
        string path)
    {
        var oldAxis = OptionalProperty(before.Raw, axisName);
        var newAxis = OptionalProperty(after.Raw, axisName);
        if (JsonEqual(oldAxis, newAxis)) return false;
        if (oldAxis is null || newAxis is null || target is null)
            throw Unsupported(path + "." + axisName, "source-bound chart-axis topology change");
        RequireEqualExcept(oldAxis.Value, newAxis.Value, path + "." + axisName, "textStyle", "titleTextStyle");
        var changed = false;
        if (PropertyChanged(oldAxis, newAxis, "textStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + "." + axisName + ".textStyle");
            target.TextStyle = SourceBoundChartTextStyle(newAxis, "textStyle", path + "." + axisName + ".textStyle");
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "titleTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + "." + axisName + ".titleTextStyle");
            var titleTextStyle = SourceBoundChartTextStyle(
                newAxis,
                "titleTextStyle",
                path + "." + axisName + ".titleTextStyle");
            if (titleTextStyle is not null && target.Title.Length == 0)
                throw Unsupported(path + "." + axisName + ".titleTextStyle", "axis-title style without an existing title");
            target.TitleTextStyle = titleTextStyle;
            changed = true;
        }
        return changed;
    }

    private static void RequireOnlyBoundedProperties(
        JsonElement? before,
        JsonElement? after,
        string path,
        params string[] properties)
    {
        if (before is { } oldValue && after is { } newValue)
        {
            RequireEqualExcept(oldValue, newValue, path, properties);
            return;
        }
        var present = before ?? after!.Value;
        foreach (var item in present.EnumerateObject())
            if (!properties.Contains(item.Name, StringComparer.Ordinal))
                throw Unsupported(path + "." + item.Name, $"changing source-owned {item.Name}");
    }

    private static SpreadsheetChartTextStyleArtifact? SourceBoundChartTextStyle(
        JsonElement? owner,
        string property,
        string path)
    {
        if (owner is null || !owner.Value.TryGetProperty(property, out var source)) return null;
        var output = new SpreadsheetChartTextStyleArtifact();
        if (source.TryGetProperty("fontSize", out var fontSize)) output.FontSizePoints = fontSize.GetDouble();
        if (source.TryGetProperty("fontFamily", out var fontFamily)) output.FontFamily = fontFamily.GetString()!;
        if (source.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) output.FontFamilyEastAsia = eastAsia.GetString()!;
        if (source.TryGetProperty("bold", out var bold)) output.Bold = bold.GetBoolean();
        if (source.TryGetProperty("italic", out var italic)) output.Italic = italic.GetBoolean();
        if (source.TryGetProperty("color", out var color))
        {
            if (color.ValueKind != JsonValueKind.String)
                throw Unsupported(path + ".color", "theme-token chart color in a source-bound edit");
            var value = color.GetString()!.TrimStart('#');
            if (value.Length is not (6 or 8) || !value.All(Uri.IsHexDigit))
                throw Unsupported(path + ".color", "non-RGB chart text color");
            output.ColorRgb = value[..6].ToUpperInvariant();
            if (value.Length == 8)
                output.OpacityThousandthPercent = Unit(Convert.ToByte(value[6..], 16) / 255d);
        }
        return output;
    }

    private static SpreadsheetChartSurfaceFill? SourceBoundChartFill(
        JsonElement? owner,
        string property,
        string path)
    {
        if (owner is null || !owner.Value.TryGetProperty(property, out var fill)) return null;
        return SourceBoundChartFill(fill, path);
    }

    private static SpreadsheetChartSurfaceFill SourceBoundChartFill(JsonElement fill, string path)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return new SpreadsheetChartSurfaceFill { NoFill = true };
        if (type == "gradient") return new SpreadsheetChartSurfaceFill { GradientFill = BuildGradientFill(fill, path) };
        if (type != "solid") throw Unsupported(path, "source-bound chart paint outside none, solid, or bounded gradient fill");
        var output = new SpreadsheetChartSurfaceFill
        {
            SolidRgb = Rgb(fill.GetProperty("color"), path + ".color"),
        };
        if (fill.TryGetProperty("opacity", out var opacity)) output.OpacityThousandthPercent = Unit(opacity.GetDouble());
        return output;
    }

    private static bool ApplyTableElement(PpjTableElementModel before, PpjTableElementModel after, PresentationTable target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "rows");
        var changed = ApplyFrame(before, after, target, path);
        if (!PropertyChanged(before.Raw, after.Raw, "rows")) return changed;
        RequireCapability(after, "replaceText", path + ".rows");
        if (before.Rows.Count != after.Rows.Count || before.Rows.Count != target.Rows.Count)
            throw Unsupported(path + ".rows", "table row topology change");
        for (var row = 0; row < before.Rows.Count; row++)
        {
            var oldRow = before.Rows[row];
            var newRow = after.Rows[row];
            if (oldRow.Id != newRow.Id || oldRow.Height != newRow.Height || oldRow.Cells.Count != newRow.Cells.Count)
                throw Unsupported($"{path}.rows[{row}]", "table row topology or geometry change");
            for (var cell = 0; cell < oldRow.Cells.Count; cell++)
            {
                var oldCell = oldRow.Cells[cell];
                var newCell = newRow.Cells[cell];
                if (oldCell.Id != newCell.Id || oldCell.RowSpan != newCell.RowSpan || oldCell.ColumnSpan != newCell.ColumnSpan)
                    throw Unsupported($"{path}.rows[{row}].cells[{cell}]", "table cell topology change");
                if (TextEqual(oldCell.Text, newCell.Text)) continue;
                if (newCell.Text.PlainText is null || !TryProjectedCellCoordinates(oldCell.Id, out var physicalRow, out var physicalColumn) ||
                    physicalRow != row || physicalColumn >= target.Rows[row].Cells.Count)
                    throw Unsupported($"{path}.rows[{row}].cells[{cell}].text", "rich or noncanonical imported table-cell text");
                target.Rows[row].Cells[physicalColumn].Text = newCell.Text.PlainText;
                changed = true;
            }
        }
        return changed;
    }

    private static bool ApplyConnectorElement(PpjConnectorElementModel before, PpjConnectorElementModel after, PresentationConnector target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "stroke");
        var oldFrame = before.Frame;
        var changed = ApplyConnectorFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".stroke");
            ApplyConnectorStroke(after.Raw.GetProperty("stroke"), target, path + ".stroke");
            changed = true;
        }
        _ = oldFrame;
        return changed;
    }

    private static bool ApplyGroupElement(
        PpjGroupElementModel before,
        PpjGroupElementModel after,
        PresentationGroup target,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> nativeLeafBindings,
        ISet<string> changedNodeIds,
        MutationState mutations,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "elements");
        var changed = ApplyFrame(before, after, target, path);
        if (changed) mutations.SemanticChanges = true;
        if (before.Elements.Count != after.Elements.Count || before.Elements.Count != target.Children.Count)
            throw Unsupported(path + ".elements", "source-bound group child insertion or deletion");
        var sourceChildren = before.Elements.Select((element, index) => new
        {
            Program = element,
            Wire = target.Children[index],
        }).ToDictionary(item => item.Program.Id, StringComparer.Ordinal);
        var requestedIds = after.Elements.Select(element => element.Id).ToArray();
        if (requestedIds.Distinct(StringComparer.Ordinal).Count() != requestedIds.Length ||
            requestedIds.Any(id => !sourceChildren.ContainsKey(id)))
            throw Unsupported(path + ".elements", "source-bound group child identity change");
        var requestedChildren = new List<PresentationElement>(after.Elements.Count);
        for (var index = 0; index < after.Elements.Count; index++)
        {
            var sourceChild = sourceChildren[after.Elements[index].Id];
            var child = sourceChild.Wire;
            var childPath = shapeTreePath.Concat([child.Source?.ShapeTreeIndex ?? checked((uint)index)]).ToArray();
            changed |= ApplyElement(sourceChild.Program, after.Elements[index], child, slide, childPath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, $"{path}.elements[{index}]");
            requestedChildren.Add(child);
        }
        var sourceOrder = before.Elements.Select(element => element.Id).ToArray();
        if (!sourceOrder.SequenceEqual(requestedIds, StringComparer.Ordinal))
        {
            for (var index = 0; index < after.Elements.Count; index++)
            {
                if (sourceOrder[index] == requestedIds[index]) continue;
                RequireCapability(after.Elements[index], "reorder", $"{path}.elements[{index}]");
                changedNodeIds.Add(after.Elements[index].Id);
            }
            mutations.SemanticChanges = true;
            changed = true;
        }
        target.Children.Clear();
        target.Children.Add(requestedChildren);
        return changed;
    }

    private static bool ApplyOpaqueElement(
        PpjOpaqueElementModel before,
        PpjOpaqueElementModel after,
        PresentationElement target,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        MutationState mutations,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "visibleText");
        var changed = false;
        if (FrameChanged(before, after))
        {
            changed = target.ContentCase switch
            {
                PresentationElement.ContentOneofCase.Shape => ApplyFrame(before, after, target.Shape, path),
                PresentationElement.ContentOneofCase.Image => ApplyFrame(before, after, target.Image, path),
                PresentationElement.ContentOneofCase.Chart => ApplyFrame(before, after, target.Chart, path),
                PresentationElement.ContentOneofCase.Table => ApplyFrame(before, after, target.Table, path),
                PresentationElement.ContentOneofCase.Connector => ApplyConnectorFrame(before, after, target.Connector, path),
                PresentationElement.ContentOneofCase.Group => ApplyFrame(before, after, target.Group, path),
                PresentationElement.ContentOneofCase.Opaque => ApplyFrame(before, after, target.Opaque, path),
                _ => throw Unsupported(path + ".frame", "placement of an unrecognized source object"),
            };
            mutations.SemanticChanges = true;
        }
        if (!before.VisibleText.SequenceEqual(after.VisibleText, StringComparer.Ordinal))
        {
            RequireCapability(after, "replaceText", path + ".visibleText");
            if (target.ContentCase != PresentationElement.ContentOneofCase.Opaque ||
                !PpjNativeTextProjection.TryRead(target.Opaque.RawXml, out var sourceLeaves) ||
                !sourceLeaves.SequenceEqual(before.VisibleText, StringComparer.Ordinal) ||
                before.VisibleText.Count != after.VisibleText.Count)
                throw Unsupported(path + ".visibleText", "stale or topologically changed opaque text leaves");
            for (var index = 0; index < before.VisibleText.Count; index++)
            {
                if (before.VisibleText[index] == after.VisibleText[index]) continue;
                mutations.NativeLeaves.Add(new NativeLeafMutation(
                    after.Id,
                    slide,
                    target,
                    shapeTreePath,
                    0,
                    checked((uint)index),
                    before.VisibleText[index],
                    after.VisibleText[index],
                    "nativeText"));
                changed = true;
            }
        }
        return changed;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationShape target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        var allowTransform = target.Placeholder is null;
        RequireFrameChange(before, after, path, allowTransform);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        if (allowTransform) target.Transform = ShapeTransform(after.Frame);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationImage target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: true);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        target.Transform = ImageTransform(after.Frame);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationChart target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: true);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        target.FrameTransform = FrameTransform(after.Frame);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationTable target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: true);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        target.FrameTransform = FrameTransform(after.Frame);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationGroup target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: true);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        target.FrameTransform = FrameTransform(after.Frame);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationOpaqueElement target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: false);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyConnectorFrame(PpjElementModel before, PpjElementModel after, PresentationConnector target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path, allowTransform: false);
        var old = before.Frame;
        var next = after.Frame;
        target.StartXEmu = TransformCoordinate(target.StartXEmu, old.X, old.Width, next.X, next.Width);
        target.EndXEmu = TransformCoordinate(target.EndXEmu, old.X, old.Width, next.X, next.Width);
        target.StartYEmu = TransformCoordinate(target.StartYEmu, old.Y, old.Height, next.Y, next.Height);
        target.EndYEmu = TransformCoordinate(target.EndYEmu, old.Y, old.Height, next.Y, next.Height);
        return true;
    }

    private static void RequireFrameChange(
        PpjElementModel before,
        PpjElementModel after,
        string path,
        bool allowTransform)
    {
        RequireCapability(after, "setFrame", path + ".frame");
        var oldFrame = before.Raw.GetProperty("frame");
        var newFrame = after.Raw.GetProperty("frame");
        if (allowTransform)
            RequireEqualExcept(oldFrame, newFrame, path + ".frame", "x", "y", "width", "height", "rotation", "flipH", "flipV");
        else
            RequireEqualExcept(oldFrame, newFrame, path + ".frame", "x", "y", "width", "height");
    }

    private static PresentationShapeTransform? ShapeTransform(PpjFrameModel frame)
    {
        if (frame.Rotation == 0 && !frame.FlipH && !frame.FlipV) return null;
        var transform = new PresentationShapeTransform();
        if (frame.Rotation != 0) transform.RotationAngle60000 = RotationAngle(frame.Rotation);
        if (frame.FlipH) transform.FlipHorizontal = true;
        if (frame.FlipV) transform.FlipVertical = true;
        return transform;
    }

    private static PresentationImageTransform? ImageTransform(PpjFrameModel frame)
    {
        if (frame.Rotation == 0 && !frame.FlipH && !frame.FlipV) return null;
        var transform = new PresentationImageTransform();
        if (frame.Rotation != 0) transform.RotationAngle60000 = RotationAngle(frame.Rotation);
        if (frame.FlipH) transform.FlipHorizontal = true;
        if (frame.FlipV) transform.FlipVertical = true;
        return transform;
    }

    private static PresentationFrameTransform? FrameTransform(PpjFrameModel frame)
    {
        if (frame.Rotation == 0 && !frame.FlipH && !frame.FlipV) return null;
        var transform = new PresentationFrameTransform();
        if (frame.Rotation != 0) transform.RotationAngle60000 = RotationAngle(frame.Rotation);
        if (frame.FlipH) transform.FlipHorizontal = true;
        if (frame.FlipV) transform.FlipVertical = true;
        return transform;
    }

    private static bool CollectTextLeafMutations(
        PpjTextContentModel before,
        PpjTextContentModel after,
        JsonElement beforeRaw,
        JsonElement afterRaw,
        PresentationShape target,
        string programElementId,
        PresentationSlide slide,
        PresentationElement element,
        IReadOnlyList<uint> shapeTreePath,
        MutationState mutations,
        string path)
    {
        if (JsonEqual(beforeRaw, afterRaw)) return false;
        if (!JsonEqual(MaskTextValues(beforeRaw), MaskTextValues(afterRaw)))
            throw Unsupported(path, "rich-text topology or styling change");
        if (target.TextBody is null)
            throw Unsupported(path, "text edit without one imported bounded text body");

        if (before.PlainText is not null || after.PlainText is not null)
        {
            if (before.PlainText is null || after.PlainText is null ||
                target.TextBody.Paragraphs.Count != 1 || target.TextBody.Paragraphs[0].Runs.Count != 1 ||
                target.TextBody.Paragraphs[0].Runs[0].ContentCase != PresentationTextRun.ContentOneofCase.Text)
                throw Unsupported(path, "plain/rich text conversion or multi-leaf plain replacement");
            if (before.PlainText != after.PlainText)
                mutations.NativeLeaves.Add(new NativeLeafMutation(
                    programElementId,
                    slide,
                    element,
                    shapeTreePath,
                    0,
                    0,
                    before.PlainText,
                    after.PlainText,
                    "text"));
        }
        else
        {
            if (before.Paragraphs.Count != after.Paragraphs.Count || before.Paragraphs.Count != target.TextBody.Paragraphs.Count)
                throw Unsupported(path, "paragraph topology change");
            uint leafIndex = 0;
            for (var paragraph = 0; paragraph < before.Paragraphs.Count; paragraph++)
            {
                if (before.Paragraphs[paragraph].Runs.Count != after.Paragraphs[paragraph].Runs.Count ||
                    before.Paragraphs[paragraph].Runs.Count != target.TextBody.Paragraphs[paragraph].Runs.Count)
                    throw Unsupported(path, "run topology change");
                for (var run = 0; run < before.Paragraphs[paragraph].Runs.Count; run++)
                {
                    var targetRun = target.TextBody.Paragraphs[paragraph].Runs[run];
                    if (targetRun.ContentCase != PresentationTextRun.ContentOneofCase.Text)
                        throw Unsupported(path, "non-text imported run mutation");
                    var oldText = before.Paragraphs[paragraph].Runs[run].Text;
                    var newText = after.Paragraphs[paragraph].Runs[run].Text;
                    if (oldText != newText)
                        mutations.NativeLeaves.Add(new NativeLeafMutation(
                            programElementId,
                            slide,
                            element,
                            shapeTreePath,
                            0,
                            leafIndex,
                            oldText,
                            newText,
                            "text"));
                    leafIndex++;
                }
            }
        }
        return true;
    }

    private static bool ApplyShapeStyle(
        PpjElementModel before,
        PpjElementModel after,
        PresentationShape target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        var oldStyle = OptionalProperty(before.Raw, "style");
        var newStyle = OptionalProperty(after.Raw, "style");
        if (JsonEqual(oldStyle, newStyle)) return false;
        if (oldStyle is { } oldValue && newStyle is { } newValue)
            RequireEqualExcept(oldValue, newValue, path + ".style", "fill", "stroke");
        else
        {
            var present = oldStyle ?? newStyle!.Value;
            foreach (var property in present.EnumerateObject())
                if (property.Name is not ("fill" or "stroke"))
                    throw Unsupported(path + ".style", $"changing {property.Name}");
        }
        var changed = false;
        if (PropertyChanged(oldStyle, newStyle, "fill"))
        {
            RequireCapability(after, "setFill", path + ".style.fill");
            ApplyFill(newStyle is { } style && style.TryGetProperty("fill", out var fill) ? fill : (JsonElement?)null,
                target, assets, assetDimensions, path + ".style.fill");
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".style.stroke");
            ApplyStroke(newStyle is { } style && style.TryGetProperty("stroke", out var stroke) ? stroke : (JsonElement?)null, target, path + ".style.stroke");
            changed = true;
        }
        return changed;
    }

    private static bool ApplyFillProperty(
        PpjElementModel before,
        PpjElementModel after,
        PresentationShape target,
        string name,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setFill", $"{path}.{name}");
        ApplyFill(after.Raw.TryGetProperty(name, out var fill) ? fill : (JsonElement?)null,
            target, assets, assetDimensions, $"{path}.{name}");
        return true;
    }

    private static bool ApplyStrokeProperty(PpjElementModel before, PpjElementModel after, PresentationShape target, string name, string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setStroke", $"{path}.{name}");
        ApplyStroke(after.Raw.TryGetProperty(name, out var stroke) ? stroke : (JsonElement?)null, target, $"{path}.{name}");
        return true;
    }

    private static void ApplyFill(
        JsonElement? fill,
        PresentationShape target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        if (fill is null || fill.Value.GetProperty("type").GetString() == "none")
        {
            target.FillRgb = string.Empty;
            target.ClearFillOpacityThousandthPercent();
            target.GradientFill = null;
            target.ImageFill = null;
            return;
        }
        if (fill.Value.GetProperty("type").GetString() == "gradient")
        {
            target.FillRgb = string.Empty;
            target.ClearFillOpacityThousandthPercent();
            target.GradientFill = BuildGradientFill(fill.Value, path);
            target.ImageFill = null;
            return;
        }
        if (fill.Value.GetProperty("type").GetString() == "image")
        {
            target.FillRgb = string.Empty;
            target.ClearFillOpacityThousandthPercent();
            target.GradientFill = null;
            target.ImageFill = BuildImagePaint(fill.Value, target.WidthEmu / 12_700d, target.HeightEmu / 12_700d,
                assets, assetDimensions, path);
            return;
        }
        if (fill.Value.GetProperty("type").GetString() != "solid")
            throw Unsupported(path, "unsupported fill");
        target.GradientFill = null;
        target.ImageFill = null;
        target.FillRgb = Rgb(fill.Value.GetProperty("color"), path + ".color");
        if (fill.Value.TryGetProperty("opacity", out var opacity)) target.FillOpacityThousandthPercent = Unit(opacity.GetDouble());
        else target.ClearFillOpacityThousandthPercent();
    }

    private static void ApplyStroke(JsonElement? stroke, PresentationShape target, string path)
    {
        if (stroke is null)
        {
            target.LineRgb = string.Empty;
            target.LineStyle = "none";
            target.LineWidthEmu = 0;
            return;
        }
        target.LineRgb = Rgb(stroke.Value.GetProperty("color"), path + ".color");
        target.LineWidthEmu = Emu(stroke.Value.GetProperty("width").GetDouble());
        target.LineStyle = NativeDash(OptionalString(stroke.Value, "dash"));
        target.LineCap = OptionalString(stroke.Value, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke.Value, "join") ?? string.Empty;
        if (stroke.Value.TryGetProperty("opacity", out var opacity)) target.LineOpacityThousandthPercent = Unit(opacity.GetDouble());
        else target.ClearLineOpacityThousandthPercent();
    }

    private static void ApplyConnectorStroke(JsonElement stroke, PresentationConnector target, string path)
    {
        target.LineRgb = Rgb(stroke.GetProperty("color"), path + ".color");
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = NativeDash(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        if (stroke.TryGetProperty("opacity", out var opacity)) target.LineOpacityThousandthPercent = Unit(opacity.GetDouble());
        else target.ClearLineOpacityThousandthPercent();
    }

    private static PresentationGradientFill BuildGradientFill(JsonElement fill, string path)
    {
        var kind = OptionalString(fill, "kind") ?? "linear";
        var output = new PresentationGradientFill
        {
            Kind = kind switch
            {
                "linear" => PresentationGradientFill.Types.Kind.Linear,
                "radial" => PresentationGradientFill.Types.Kind.Radial,
                _ => throw Unsupported(path, $"gradient kind {kind}"),
            },
        };
        if (output.Kind == PresentationGradientFill.Types.Kind.Linear)
        {
            var degrees = fill.TryGetProperty("angle", out var angle) ? angle.GetDouble() : 0;
            var normalized = ((degrees % 360) + 360) % 360;
            output.Angle60000 = checked((int)Math.Round(normalized * 60_000, MidpointRounding.AwayFromZero));
        }
        else if (fill.TryGetProperty("angle", out _))
        {
            throw Unsupported(path, "radial gradient with a linear angle");
        }
        var index = 0;
        foreach (var item in fill.GetProperty("stops").EnumerateArray())
        {
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = Unit(item.GetProperty("offset").GetDouble()),
                ColorRgb = Rgb(item.GetProperty("color"), $"{path}.stops[{index}].color"),
            };
            if (item.TryGetProperty("opacity", out var opacity))
                stop.OpacityThousandthPercent = Unit(opacity.GetDouble());
            output.Stops.Add(stop);
            index++;
        }
        PptxGradientFillCodec.Validate(output, path);
        return output;
    }

    private static void ApplyChartData(PpjChartElementModel before, PpjChartElementModel after, PresentationChart target, string path)
    {
        if (before.Data.Categories.Count != after.Data.Categories.Count ||
            before.Data.Series.Count != after.Data.Series.Count)
            throw Unsupported(path, "chart series or category topology change");
        var valueChanged = !before.Data.Categories.Zip(after.Data.Categories).All(pair => JsonEqual(pair.First, pair.Second)) ||
            before.Data.Series.Zip(after.Data.Series).Any(pair =>
                !pair.First.Name.Equals(pair.Second.Name, StringComparison.Ordinal) ||
                !pair.First.Values.SequenceEqual(pair.Second.Values));
        var fillChanged = before.Data.Series.Zip(after.Data.Series)
            .Any(pair => PropertyChanged(pair.First.Raw, pair.Second.Raw, "fill"));
        if (valueChanged) RequireCapability(after, "setChartData", path);
        if (fillChanged) RequireCapability(after, "setChartFill", path + ".series[].fill");
        var categories = after.Data.Categories.Select((item, index) => item.ValueKind == JsonValueKind.String
            ? item.GetString()!
            : throw Unsupported($"{path}.categories[{index}]", "non-string imported category")).ToArray();
        target.Categories.Clear();
        target.Categories.Add(categories);

        if (target.Type == SpreadsheetChartType.Combo)
        {
            if (target.ComboSeries.Count != after.Data.Series.Count)
                throw Unsupported(path, "combo-chart series topology change");
            for (var index = 0; index < after.Data.Series.Count; index++)
                ApplyChartSeries(before.Data.Series[index], after.Data.Series[index], target.ComboSeries[index].Series, $"{path}.series[{index}]");
        }
        else
        {
            if (target.Series.Count != after.Data.Series.Count)
                throw Unsupported(path, "chart series topology change");
            for (var index = 0; index < after.Data.Series.Count; index++)
                ApplyChartSeries(before.Data.Series[index], after.Data.Series[index], target.Series[index], $"{path}.series[{index}]");
        }
    }

    private static void ApplyChartSeries(PpjChartSeriesModel before, PpjChartSeriesModel after, SpreadsheetChartSeriesArtifact target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "name", "values", "fill");
        if (before.Id != after.Id || before.ChartType != after.ChartType || before.Axis != after.Axis || before.Values.Count != after.Values.Count)
            throw Unsupported(path, "chart-series identity or topology change");
        if (!before.Values.Select(value => value is null).SequenceEqual(after.Values.Select(value => value is null)))
            throw Unsupported(path, "source-bound chart missing-point topology change");
        target.Name = after.Name;
        target.Values.Clear();
        target.MissingValueIndexes.Clear();
        for (var index = 0; index < after.Values.Count; index++)
        {
            var value = after.Values[index];
            if (value is null)
            {
                target.Values.Add(0d);
                target.MissingValueIndexes.Add(checked((uint)index));
            }
            else target.Values.Add(value.Value);
        }
        if (PropertyChanged(before.Raw, after.Raw, "fill"))
        {
            target.Fill = null;
            target.SeriesFill = after.Raw.TryGetProperty("fill", out var fill)
                ? SourceBoundChartFill(fill, path + ".fill")
                : null;
        }
    }

    private static bool FrameChanged(PpjElementModel before, PpjElementModel after) =>
        !JsonEqual(before.Raw.GetProperty("frame"), after.Raw.GetProperty("frame"));

    private static void RequireNativeRef(JsonElement before, JsonElement after, string path)
    {
        if (!before.TryGetProperty("nativeRef", out var oldReference) ||
            !after.TryGetProperty("nativeRef", out var newReference) ||
            !NativeRefEqualExceptLeafValues(oldReference, newReference))
            throw new CodecException(
                "ppj.nativeRef.stale",
                "The source-bound PPJ nativeRef is missing, changed outside an issued leaf value, or no longer matches the exact source projection.",
                path + ".nativeRef");
    }

    private static bool CollectIssuedNativeLeafMutations(
        PpjNativeRefModel? before,
        PpjNativeRefModel? after,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> bindings,
        PresentationSlide slide,
        PresentationElement element,
        IReadOnlyList<uint> shapeTreePath,
        string programElementId,
        MutationState mutations,
        string path)
    {
        if (before is null || after is null || before.Leaves.Count == 0) return false;
        if (before.Leaves.Count != after.Leaves.Count)
            throw new CodecException("ppj.nativeRef.stale", "The issued native-leaf set changed.", path);

        var changed = false;
        for (var index = 0; index < before.Leaves.Count; index++)
        {
            var oldLeaf = before.Leaves[index];
            var newLeaf = after.Leaves[index];
            var leafPath = $"{path}[{index}]";
            if (!oldLeaf.Id.Equals(newLeaf.Id, StringComparison.Ordinal) ||
                !oldLeaf.Kind.Equals(newLeaf.Kind, StringComparison.Ordinal) ||
                !oldLeaf.ExpectedHash.Equals(newLeaf.ExpectedHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.nativeRef.stale", "A native leaf changed its issued identity, kind, or expected hash.", leafPath);
            if (JsonEqual(oldLeaf.Value, newLeaf.Value)) continue;
            if (!bindings.TryGetValue(oldLeaf.Id, out var binding) ||
                !binding.ElementId.Equals(programElementId, StringComparison.Ordinal) ||
                !binding.Kind.Equals(oldLeaf.Kind, StringComparison.Ordinal))
                throw new CodecException("ppj.nativeRef.stale", "The native leaf is not part of the fresh source projection.", leafPath);

            var expectedHash = Sha256(Encoding.UTF8.GetBytes(binding.ExpectedValue));
            if (!oldLeaf.ExpectedHash.Equals(expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.nativeRef.stale", "The native leaf expected hash no longer matches the exact source value.", leafPath + ".expectedHash");
            var oldValue = PpjNativeLeafProjection.NormalizeValue(oldLeaf.Kind, oldLeaf.Value, leafPath + ".value");
            if (!oldValue.Equals(binding.ExpectedValue, StringComparison.Ordinal))
                throw new CodecException("ppj.nativeRef.stale", "The native leaf baseline value no longer matches the exact source projection.", leafPath + ".value");
            var newValue = PpjNativeLeafProjection.NormalizeValue(newLeaf.Kind, newLeaf.Value, leafPath + ".value");
            if (newValue.Equals(oldValue, StringComparison.Ordinal)) continue;

            mutations.NativeLeaves.Add(new NativeLeafMutation(
                programElementId,
                slide,
                element,
                shapeTreePath,
                binding.NativeLeafIndex,
                binding.TextLeafIndex,
                oldValue,
                newValue,
                binding.Kind));
            changed = true;
        }
        return changed;
    }

    private static bool NativeRefEqualExceptLeafValues(JsonElement before, JsonElement after)
    {
        var oldBytes = PpjCanonicalJson.Write(before);
        var newBytes = PpjCanonicalJson.Write(after);
        var oldNode = System.Text.Json.Nodes.JsonNode.Parse(oldBytes)!.AsObject();
        var newNode = System.Text.Json.Nodes.JsonNode.Parse(newBytes)!.AsObject();
        MaskLeafValues(oldNode);
        MaskLeafValues(newNode);
        using var oldDocument = JsonDocument.Parse(oldNode.ToJsonString());
        using var newDocument = JsonDocument.Parse(newNode.ToJsonString());
        return JsonEqual(oldDocument.RootElement, newDocument.RootElement);
    }

    private static void MaskLeafValues(System.Text.Json.Nodes.JsonObject nativeRef)
    {
        if (nativeRef["leaves"] is not System.Text.Json.Nodes.JsonArray leaves) return;
        foreach (var leaf in leaves.OfType<System.Text.Json.Nodes.JsonObject>()) leaf["value"] = null;
    }

    private static void RequireCapability(PpjElementModel element, string operation, string path)
    {
        RequireCapability(element.NativeRef, operation, path);
    }

    private static void RequireCapability(PpjNativeRefModel? nativeRef, string operation, string path)
    {
        var reference = nativeRef ?? throw new CodecException("ppj.nativeRef.missing", "Source-bound edits require a nativeRef.", path);
        var capability = reference.Capabilities.FirstOrDefault(item => item.Operation.Equals(operation, StringComparison.Ordinal));
        if (capability is null || !capability.ExpectedHash.Equals(reference.ObjectHash, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.nativeRef.capabilityMissing",
                $"The exact source object did not issue the {operation} capability required by this edit.",
                path);
    }

    private static void RequireEqualExcept(JsonElement before, JsonElement after, string path, params string[] allowed)
    {
        var allowedNames = allowed.ToHashSet(StringComparer.Ordinal);
        var names = before.EnumerateObject().Select(property => property.Name)
            .Concat(after.EnumerateObject().Select(property => property.Name))
            .Distinct(StringComparer.Ordinal);
        foreach (var name in names)
        {
            if (allowedNames.Contains(name)) continue;
            var oldPresent = before.TryGetProperty(name, out var oldValue);
            var newPresent = after.TryGetProperty(name, out var newValue);
            var equal = oldPresent && newPresent &&
                (name == "nativeRef"
                    ? NativeRefEqualExceptLeafValues(oldValue, newValue)
                    : JsonEqual(oldValue, newValue));
            if (oldPresent != newPresent || oldPresent && !equal)
                throw Unsupported(path + "." + name, $"changing source-owned {name}");
        }
    }

    private static void RequirePropertyEqual(JsonElement before, JsonElement after, string name, string path)
    {
        var oldPresent = before.TryGetProperty(name, out var oldValue);
        var newPresent = after.TryGetProperty(name, out var newValue);
        if (oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue))
            throw Unsupported(path, "changing source-owned presentation topology");
    }

    private static bool PropertyChanged(JsonElement before, JsonElement after, string name)
    {
        var oldPresent = before.TryGetProperty(name, out var oldValue);
        var newPresent = after.TryGetProperty(name, out var newValue);
        return oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue);
    }

    private static bool PropertyChanged(JsonElement? before, JsonElement? after, string name)
    {
        var oldValue = default(JsonElement);
        var newValue = default(JsonElement);
        var oldPresent = before is { } oldObject && oldObject.TryGetProperty(name, out oldValue);
        var newPresent = after is { } newObject && newObject.TryGetProperty(name, out newValue);
        return oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue);
    }

    private static JsonElement? OptionalProperty(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value : null;

    private static JsonElement MaskTextValues(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.String)
        {
            using var plain = JsonDocument.Parse("\"\"");
            return plain.RootElement.Clone();
        }
        var bytes = PpjCanonicalJson.Write(value);
        var node = System.Text.Json.Nodes.JsonNode.Parse(bytes)!.AsObject();
        if (node["paragraphs"] is System.Text.Json.Nodes.JsonArray paragraphs)
            foreach (var paragraph in paragraphs.OfType<System.Text.Json.Nodes.JsonObject>())
                if (paragraph["runs"] is System.Text.Json.Nodes.JsonArray runs)
                    foreach (var run in runs.OfType<System.Text.Json.Nodes.JsonObject>()) run["text"] = string.Empty;
        using var document = JsonDocument.Parse(node.ToJsonString());
        return document.RootElement.Clone();
    }

    private static bool TextEqual(PpjTextContentModel left, PpjTextContentModel right)
    {
        if (left.PlainText is not null || right.PlainText is not null) return left.PlainText == right.PlainText;
        if (left.Paragraphs.Count != right.Paragraphs.Count) return false;
        for (var paragraph = 0; paragraph < left.Paragraphs.Count; paragraph++)
        {
            if (left.Paragraphs[paragraph].Runs.Count != right.Paragraphs[paragraph].Runs.Count) return false;
            for (var run = 0; run < left.Paragraphs[paragraph].Runs.Count; run++)
                if (left.Paragraphs[paragraph].Runs[run].Text != right.Paragraphs[paragraph].Runs[run].Text) return false;
        }
        return true;
    }

    private static bool TryProjectedCellCoordinates(string? id, out int row, out int column)
    {
        row = column = -1;
        if (id is null || !id.StartsWith("cell-", StringComparison.Ordinal)) return false;
        var values = id[5..].Split('-');
        return values.Length == 2 && int.TryParse(values[0], out var oneBasedRow) && int.TryParse(values[1], out var oneBasedColumn) &&
               (row = oneBasedRow - 1) >= 0 && (column = oneBasedColumn - 1) >= 0;
    }

    private static PresentationEditPlanRequest NativeLeafEditPlan(
        string sourceSha256,
        IReadOnlyList<NativeLeafMutation> mutations)
    {
        var plan = new PresentationEditPlanRequest { ExpectedSourceSha256 = sourceSha256 };
        foreach (var mutation in mutations)
        {
            var slideSource = mutation.Slide.Source ??
                throw new CodecException("ppj.nativeRef.stale", "Text edit lost its source slide binding.", mutation.ProgramElementId);
            var elementSource = mutation.Element.Source ??
                throw new CodecException("ppj.nativeRef.stale", "Text edit lost its source element binding.", mutation.ProgramElementId);
            if (mutation.ShapeTreePath.Count == 0 || string.IsNullOrEmpty(slideSource.PartPath) ||
                string.IsNullOrEmpty(slideSource.SlideXmlSha256) || string.IsNullOrEmpty(elementSource.ElementSha256) ||
                string.IsNullOrEmpty(elementSource.SemanticSha256))
                throw new CodecException("ppj.nativeRef.stale", "Text edit has an incomplete source-bound compiler binding.", mutation.ProgramElementId);
            var seed = string.Join("\0", sourceSha256, mutation.ProgramElementId, mutation.LeafKind,
                mutation.NativeLeafIndex, mutation.TextLeafIndex, mutation.Before, mutation.After);
            var operation = new PresentationEditOperation
            {
                OperationId = $"ppj-{mutation.LeafKind}-{Sha256(Encoding.UTF8.GetBytes(seed))[..20]}",
                SlideId = mutation.Slide.Id,
                SlidePartPath = slideSource.PartPath,
                ExpectedSlideSha256 = slideSource.SlideXmlSha256,
                TargetId = mutation.Element.Id,
                ShapeTreeIndex = mutation.ShapeTreePath[0],
                ExpectedElementSha256 = elementSource.ElementSha256,
                ExpectedSemanticSha256 = elementSource.SemanticSha256,
                NativeLeafIndex = mutation.NativeLeafIndex,
                TextLeafIndex = mutation.TextLeafIndex,
                ExpectedTextSha256 = Sha256(Encoding.UTF8.GetBytes(mutation.Before)),
                ExpectedValue = mutation.Before,
                Value = mutation.After,
                LeafKind = mutation.LeafKind,
            };
            operation.ShapeTreePath.Add(mutation.ShapeTreePath);
            plan.Operations.Add(operation);
        }
        return plan;
    }

    private static string AssetRoot(PpjProgramModel program)
    {
        var uri = program.Assets.FirstOrDefault()?.Uri;
        if (string.IsNullOrEmpty(uri)) return "deck.assets/media";
        var separator = uri.LastIndexOf('/');
        return separator > 0 ? uri[..separator] : "deck.assets/media";
    }

    private static IReadOnlyList<string> ChangedParts(byte[] before, byte[] after)
    {
        if (before.AsSpan().SequenceEqual(after)) return [];
        var oldParts = ZipHashes(before);
        var newParts = ZipHashes(after);
        return oldParts.Keys.Concat(newParts.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !oldParts.TryGetValue(path, out var oldHash) || !newParts.TryGetValue(path, out var newHash) || oldHash != newHash)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static Dictionary<string, string> ZipHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var output = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(entry => !string.IsNullOrEmpty(entry.Name)))
        {
            using var input = entry.Open();
            using var memory = new MemoryStream();
            input.CopyTo(memory);
            output[entry.FullName] = Sha256(memory.ToArray());
        }
        return output;
    }

    private static string AssetHash(Asset asset) =>
        asset.Sha256 is { Length: 64 } ? asset.Sha256.ToLowerInvariant() : Sha256(asset.Data.Span);

    private static bool JsonEqual(JsonElement left, JsonElement right) =>
        PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right));

    private static bool JsonEqual(JsonElement? left, JsonElement? right) =>
        left is null || right is null ? left is null && right is null : JsonEqual(left.Value, right.Value);

    private static int RotationAngle(double degrees) =>
        checked((int)Math.Round(degrees * 60_000, MidpointRounding.AwayFromZero));

    private static long Emu(double points) => checked((long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero));
    private static uint Unit(double value) => checked((uint)Math.Round(Math.Clamp(value, 0, 1) * 100_000, MidpointRounding.AwayFromZero));
    private static int Crop(JsonElement owner, string name) => owner.TryGetProperty(name, out var value)
        ? checked((int)Math.Round(Math.Clamp(value.GetDouble(), 0, 1) * 100_000, MidpointRounding.AwayFromZero))
        : 0;

    private static long TransformCoordinate(long source, double oldStart, double oldSize, double newStart, double newSize)
    {
        var sourcePoints = source / EmuPerPoint;
        var ratio = oldSize == 0 ? 0 : (sourcePoints - oldStart) / oldSize;
        return Emu(newStart + ratio * newSize);
    }

    private static string Rgb(JsonElement color, string path)
    {
        if (color.ValueKind != JsonValueKind.String) throw Unsupported(path, "theme or computed source-bound color");
        var value = color.GetString()!.TrimStart('#');
        if (value.Length != 6 || !value.All(Uri.IsHexDigit)) throw Unsupported(path, "non-RGB source-bound color");
        return value.ToUpperInvariant();
    }

    private static string NativeDash(string? value) => value switch
    {
        null or "" or "solid" => "solid",
        "dash" => "dashed",
        "dot" => "dotted",
        "dash-dot" => "dash-dot",
        "long-dash" => "long-dash",
        _ => throw Unsupported("stroke.dash", $"unsupported dash style {value}"),
    };

    private static string? OptionalString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value.GetString() : null;

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Unsupported(string path, string detail) =>
        new("ppj.source.unsupportedMutation", $"Source-bound PPJ cannot safely compile {detail}.", path);
}
