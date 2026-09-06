using System.IO.Compression;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
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
        using var validationStage = PpjBuildProfiler.Measure("validation");
        using var validation = PpjProgramValidator.Validate(request.ProgramJson.Memory);
        validationStage.Dispose();
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        return CompileValidated(request, sourceBytes, limits, validation);
    }

    internal static PpjCompileResult CompileValidated(
        PresentationProgramRequest request,
        byte[] sourceBytes,
        EffectiveCodecLimits limits,
        PpjValidationResult validation,
        Func<PpjAssetModel, Asset?>? loadAsset = null,
        bool retainSourceAssetData = true) => CompileValidated(
            request,
            new PptxPackageSource(sourceBytes),
            limits,
            validation,
            loadAsset,
            retainSourceAssetData);

    internal static PpjCompileResult CompileValidated(
        PresentationProgramRequest request,
        PptxPackageSource sourcePackage,
        EffectiveCodecLimits limits,
        PpjValidationResult validation,
        Func<PpjAssetModel, Asset?>? loadAsset = null,
        bool retainSourceAssetData = true)
    {
        if (!validation.IsValid)
            throw new InvalidOperationException("PPJ compilation requires a successful validation result.");

        if (validation.Program!.Source is null)
        {
            if (sourcePackage.Length != 0)
                throw new CodecException(
                    "ppj.unexpectedSource",
                    "A source-free PPJ compile cannot attach a PPTX source package.",
                    "$.source");
            if (request.ValidationOnly)
                return PpjAuthoredPresentationCompiler.ValidateOnly(request, validation);
            return PpjAuthoredPresentationCompiler.Compile(request, limits, validation);
        }

        return PpjSourceBoundPresentationCompiler.Compile(
            request,
            sourcePackage,
            limits,
            validation,
            loadAsset,
            retainSourceAssetData);
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
        internal HashSet<string>? AuthoredOverlayNodeIds { get; set; }
        internal List<Diagnostic> Warnings { get; } = [];
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
        string LeafKind,
        PpjNativeChartDataBinding? ChartData = null,
        bool FrameFastPath = false);

    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        PptxPackageSource sourcePackage,
        EffectiveCodecLimits limits,
        PpjValidationResult validation,
        Func<PpjAssetModel, Asset?>? loadAsset,
        bool retainSourceAssetData)
    {
        if (sourcePackage.Length == 0)
            throw new CodecException(
                "ppj.source.missing",
                "A source-bound PPJ compile requires the exact source PPTX bytes.",
                "$.source");

        var requested = validation.Program!;
        var source = requested.Source!;
        string sourceSha256;
        using (PpjBuildProfiler.Measure("source.hash"))
            sourceSha256 = sourcePackage.Sha256();
        if (!source.Kind.Equals("pptx", StringComparison.Ordinal) ||
            !sourceSha256.Equals(source.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.source.hashMismatch",
                "The source-bound PPJ does not match the exact PPTX input bytes.",
                "$.source.sha256");

        using var projectionStage = PpjBuildProfiler.Measure("pptx.projection");
        using var projected = PpjPresentationProjector.Project(
            sourcePackage,
            new PresentationProgramRequest
            {
                SourceUri = source.Uri,
                AssetRootUri = AssetRoot(requested),
                IncludeNodeMap = true,
            },
            limits,
            retainSourceAssetData,
            sourceSha256);
        projectionStage.Dispose();
        using var reparsedBaselineValidation = projected.Validation is null
            ? PpjProgramValidator.Validate(projected.Program.ProgramJson.Memory)
            : null;
        var baselineValidation = projected.Validation ?? reparsedBaselineValidation!;
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
        IReadOnlyDictionary<string, string> assetIds;
        var changedNodeIds = new HashSet<string>(StringComparer.Ordinal);
        var mutations = new MutationState();
        bool physicalChanges;
        using (PpjBuildProfiler.Measure("ppj.mutation"))
        {
            assetIds = BuildAssetCatalog(baseline, requested, projected, request.Assets, artifact, loadAsset);
            if (ApplyDesign(baseline, requested, artifact, mutations))
                physicalChanges = true;
            else
                physicalChanges = false;
            if (ApplyCanvas(baseline, requested, presentation, changedNodeIds, mutations))
                physicalChanges = true;
            if (ApplyPages(
                baseline,
                requested,
                presentation,
                assetIds,
                projected.NativeLeafBindings,
                changedNodeIds,
                mutations))
                physicalChanges = true;
            if (ApplySections(baseline, requested, presentation, changedNodeIds))
            {
                RejectMixedAuthoredOverlay(mutations, "$.sections");
                mutations.SemanticChanges = true;
                physicalChanges = true;
            }
            if (ApplyCustomShows(baseline, requested, presentation, changedNodeIds))
            {
                RejectMixedAuthoredOverlay(mutations, "$.customShows");
                mutations.SemanticChanges = true;
                physicalChanges = true;
            }
            if (ApplyComments(baseline, requested, presentation, changedNodeIds))
            {
                RejectMixedAuthoredOverlay(mutations, "$.comments");
                mutations.SemanticChanges = true;
                physicalChanges = true;
            }
        }

        if (request.ValidationOnly)
        {
            var validationReceipt = new PresentationProgramResult
            {
                ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
                ProgramSha256 = validation.ProgramSha256,
                NodeMapJson = request.IncludeNodeMap ? projected.Program.NodeMapJson : ByteString.Empty,
                SourceSha256 = sourceSha256,
                SourceBound = true,
                ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
            };
            validationReceipt.Assets.Add(projected.Program.Assets.Select(asset => asset.Clone()));
            validationReceipt.ChangedNodeIds.Add(changedNodeIds.OrderBy(id => id, StringComparer.Ordinal));
            return new([], validationReceipt, projected.Diagnostics.Concat(mutations.Warnings).ToArray());
        }

        byte[] output;
        byte[]? materializedSource = null;
        var reuseSourceFile = false;
        var outputUsesSource = false;
        IReadOnlyList<string>? exportedChangedParts = null;
        IReadOnlyList<Diagnostic> diagnostics;
        if (!physicalChanges)
        {
            // A physical no-op is the original file, not a reserialized
            // equivalent package. The direct CLI keeps it file-backed through
            // the final exclusive copy; wire callers continue to own bytes.
            outputUsesSource = true;
            if (!sourcePackage.TryGetMaterialized(out output))
            {
                output = [];
                reuseSourceFile = true;
            }
            diagnostics = projected.Diagnostics.Concat(mutations.Warnings).ToArray();
        }
        else if (mutations.NativeLeaves.Count > 0 && !mutations.SemanticChanges)
        {
            materializedSource = sourcePackage.Materialize();
            PptxEditPlanOutput edited;
            using (PpjBuildProfiler.Measure("edit-plan"))
            {
                edited = PptxEditPlanCodec.Apply(
                    materializedSource,
                    NativeLeafEditPlan(sourceSha256, mutations.NativeLeaves),
                    limits,
                    presentation);
            }
            output = edited.File;
            diagnostics = projected.Diagnostics.Concat(edited.Diagnostics).Concat(mutations.Warnings).ToArray();
        }
        else
        {
            if (mutations.NativeLeaves.Count > 0 &&
                !mutations.NativeLeaves.All(mutation => mutation.FrameFastPath))
                throw Unsupported("$.pages", "mixing precise native-leaf edits with other semantic edits in one source-bound build");
            // Frame-only leaves are an internal acceleration of the ordinary
            // semantic writer. If another operation in the same transaction
            // requires that writer, discard the speculative leaf plan and keep
            // the pre-existing full-export behavior.
            mutations.NativeLeaves.Clear();
            materializedSource = sourcePackage.Materialize();
            artifact.OpaqueOpc!.SourcePackage = new SourcePackageSnapshot
            {
                Data = UnsafeByteOperations.UnsafeWrap(materializedSource),
                Sha256 = sourceSha256,
            };
            PptxExportResult exported;
            using (PpjBuildProfiler.Measure("writer"))
                exported = PptxCodec.Export(artifact, limits);
            output = exported.File;
            exportedChangedParts = exported.ChangedParts;
            diagnostics = projected.Diagnostics.Concat(exported.Diagnostics).Concat(mutations.Warnings).ToArray();
        }

        var receipt = new PresentationProgramResult
        {
            ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? projected.Program.NodeMapJson : ByteString.Empty,
            SourceSha256 = sourceSha256,
            OutputSha256 = outputUsesSource ? sourceSha256 : Sha256(output),
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        receipt.Assets.Add(projected.Program.Assets.Select(asset => asset.Clone()));
        receipt.ChangedNodeIds.Add(changedNodeIds.OrderBy(id => id, StringComparer.Ordinal));
        if (!outputUsesSource)
            receipt.ChangedParts.Add(exportedChangedParts ?? ChangedParts(materializedSource ?? sourcePackage.Materialize(), output));
        return new(output, receipt, diagnostics, reuseSourceFile);
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
        if (baseline.Design.Masters.Count != requested.Design.Masters.Count ||
            baseline.Design.Layouts.Count != requested.Design.Layouts.Count)
            throw Unsupported("$.design", "changing source-bound master or layout topology");
        for (var index = 0; index < baseline.Design.Masters.Count; index++)
        {
            var before = baseline.Design.Masters[index];
            var after = requested.Design.Masters[index];
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal) ||
                !before.Name.Equals(after.Name, StringComparison.Ordinal))
                throw Unsupported($"$.design.masters[{index}]", "changing source-bound master identity");
        }
        for (var index = 0; index < baseline.Design.Layouts.Count; index++)
        {
            var before = baseline.Design.Layouts[index];
            var after = requested.Design.Layouts[index];
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal) ||
                !before.Name.Equals(after.Name, StringComparison.Ordinal) ||
                !before.MasterId.Equals(after.MasterId, StringComparison.Ordinal) ||
                !before.LayoutType.Equals(after.LayoutType, StringComparison.Ordinal))
                throw Unsupported($"$.design.layouts[{index}]", "changing source-bound layout identity or binding");
        }
    }

    private static Dictionary<string, string> BuildAssetCatalog(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PpjProjectionResult projected,
        IEnumerable<Asset> supplied,
        ArtifactEnvelope artifact,
        Func<PpjAssetModel, Asset?>? loadAsset)
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
            suppliedById.TryGetValue(declaration.Id, out var asset);
            if (asset is null || asset.Data.IsEmpty)
                asset = loadAsset?.Invoke(declaration);
            if (asset is null || asset.Data.IsEmpty)
                throw new CodecException("ppj.asset.missing", $"PPJ asset {declaration.Id} has no supplied bytes.", "$.assets");
            var hash = Sha256(asset.Data.Span);
            if (!hash.Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) ||
                !hash.Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.hashMismatch", $"PPJ asset {declaration.Id} does not match its declared SHA-256.", "$.assets");
            if (!asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.mimeMismatch", $"PPJ asset {declaration.Id} does not match its declared MIME type.", "$.assets");
            var nativeId = PptxAssetCatalog.NativeAssetIdFor(declaration.MimeType, hash);
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

    private static bool ApplyDesign(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        ArtifactEnvelope artifact,
        MutationState mutations)
    {
        if (baseline.Design.Masters.Count != artifact.Presentation.Masters.Count ||
            baseline.Design.Layouts.Count != artifact.Presentation.Layouts.Count)
            throw Unsupported("$.design", "inconsistent source master or layout topology");

        // Reuse the authored lowering's bounded fill parser. It resolves only
        // declared PPJ colors/assets and produces the same PresentationML
        // value that a source-free background would receive; PptxCodec still
        // owns the source-part writeback and residual XML proof.
        var catalog = new PpjAuthoredPresentationCompiler.Catalog(requested.Root, artifact.Assets);
        var changed = false;
        for (var index = 0; index < baseline.Design.Masters.Count; index++)
        {
            var before = baseline.Design.Masters[index];
            var after = requested.Design.Masters[index];
            var path = $"$.design.masters[{index}]";
            RequireEqualExcept(before.Raw, after.Raw, path, "background", "textStyles", "placeholders", "nativeRef");
            if (before.NativeRef is null || after.NativeRef is null)
                throw new CodecException(
                    "ppj.nativeRef.stale",
                    "A source-bound master projection must retain its hash-bound nativeRef.",
                    path + ".nativeRef");
            RequireNativeRef(before.Raw, after.Raw, path);
            var backgroundChanged = PropertyChanged(before.Raw, after.Raw, "background");
            var textStylesChanged = PropertyChanged(before.Raw, after.Raw, "textStyles");
            if (backgroundChanged)
            {
                RequireCapability(after.NativeRef, "setBackground", path + ".background");
                artifact.Presentation.Masters[index].Background = after.Raw.TryGetProperty("background", out var fill)
                    ? PpjAuthoredPresentationCompiler.BuildBackground(
                        fill,
                        catalog,
                        requested.Design.Width,
                        requested.Design.Height)
                    : null;
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (textStylesChanged)
            {
                RequireCapability(after.NativeRef, "setTextParagraphStyle", path + ".textStyles");
                RequireStableMasterTextLevels(before, after, path + ".textStyles");
                artifact.Presentation.Masters[index].TextStyles = BuildMasterTextStyles(after, catalog);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (ApplyDesignPlaceholders(
                    before.Placeholders,
                    after.Placeholders,
                    artifact.Presentation.Masters[index].Placeholders,
                    path + ".placeholders",
                    requested.Root))
            {
                mutations.SemanticChanges = true;
                changed = true;
            }
        }

        for (var index = 0; index < baseline.Design.Layouts.Count; index++)
        {
            var before = baseline.Design.Layouts[index];
            var after = requested.Design.Layouts[index];
            var path = $"$.design.layouts[{index}]";
            RequireEqualExcept(before.Raw, after.Raw, path, "background", "placeholders", "nativeRef");
            if (before.NativeRef is null || after.NativeRef is null)
                throw new CodecException(
                    "ppj.nativeRef.stale",
                    "A source-bound layout projection must retain its hash-bound nativeRef.",
                    path + ".nativeRef");
            RequireNativeRef(before.Raw, after.Raw, path);
            if (PropertyChanged(before.Raw, after.Raw, "background"))
            {
                RequireCapability(after.NativeRef, "setBackground", path + ".background");
                artifact.Presentation.Layouts[index].Background = after.Raw.TryGetProperty("background", out var fill)
                    ? PpjAuthoredPresentationCompiler.BuildBackground(
                        fill,
                        catalog,
                        requested.Design.Width,
                        requested.Design.Height)
                    : null;
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (ApplyDesignPlaceholders(
                    before.Placeholders,
                    after.Placeholders,
                    artifact.Presentation.Layouts[index].Placeholders,
                    path + ".placeholders",
                    requested.Root))
            {
                mutations.SemanticChanges = true;
                changed = true;
            }
        }
        return changed;
    }

    private static void RequireStableMasterTextLevels(
        PpjMasterModel before,
        PpjMasterModel after,
        string path)
    {
        RequireStableMasterTextLevels(before.TitleTextLevels, after.TitleTextLevels, path + ".title");
        RequireStableMasterTextLevels(before.BodyTextLevels, after.BodyTextLevels, path + ".body");
        RequireStableMasterTextLevels(before.OtherTextLevels, after.OtherTextLevels, path + ".other");
    }

    private static void RequireStableMasterTextLevels(
        IReadOnlyList<JsonElement> before,
        IReadOnlyList<JsonElement> after,
        string path)
    {
        if (before.Count != after.Count ||
            !before.Select(level => level.GetProperty("level").GetInt32())
                .SequenceEqual(after.Select(level => level.GetProperty("level").GetInt32())))
            throw Unsupported(path, "changing source-bound master text-style level topology");
    }

    private static PresentationMasterTextStyles BuildMasterTextStyles(
        PpjMasterModel source,
        PpjAuthoredPresentationCompiler.Catalog catalog)
    {
        var output = new PresentationMasterTextStyles();
        output.TitleLevels.Add(source.TitleTextLevels.Select(level =>
            PpjAuthoredPresentationCompiler.BuildMasterTextLevel(level, catalog)));
        output.BodyLevels.Add(source.BodyTextLevels.Select(level =>
            PpjAuthoredPresentationCompiler.BuildMasterTextLevel(level, catalog)));
        output.OtherLevels.Add(source.OtherTextLevels.Select(level =>
            PpjAuthoredPresentationCompiler.BuildMasterTextLevel(level, catalog)));
        return output;
    }

    private static bool ApplyDesignPlaceholders(
        IReadOnlyList<PpjLayoutPlaceholderModel> before,
        IReadOnlyList<PpjLayoutPlaceholderModel> after,
        IList<PresentationPlaceholder> target,
        string path,
        JsonElement programRoot)
    {
        if (before.Count != after.Count)
            throw Unsupported(path, "changing source-bound master/layout placeholder topology");

        var changed = false;
        for (var index = 0; index < before.Count; index++)
        {
            var oldPlaceholder = before[index];
            var newPlaceholder = after[index];
            var placeholderPath = $"{path}[{index}]";
            if (!oldPlaceholder.Id.Equals(newPlaceholder.Id, StringComparison.Ordinal) ||
                !oldPlaceholder.Name.Equals(newPlaceholder.Name, StringComparison.Ordinal) ||
                !oldPlaceholder.PlaceholderType.Equals(newPlaceholder.PlaceholderType, StringComparison.Ordinal) ||
                oldPlaceholder.Index != newPlaceholder.Index)
                throw Unsupported(placeholderPath, "changing source-bound placeholder identity");

            RequireEqualExcept(oldPlaceholder.Raw, newPlaceholder.Raw, placeholderPath, "frame", "text", "style", "nativeRef");
            RequireNativeRef(oldPlaceholder.Raw, newPlaceholder.Raw, placeholderPath);
            var frameChanged = !JsonEqual(oldPlaceholder.Raw.GetProperty("frame"), newPlaceholder.Raw.GetProperty("frame"));
            var textChanged = PropertyChanged(oldPlaceholder.Raw, newPlaceholder.Raw, "text");
            var styleChanged = PropertyChanged(oldPlaceholder.Raw, newPlaceholder.Raw, "style");
            if (!frameChanged && !textChanged && !styleChanged) continue;

            if (newPlaceholder.NativeRef is null)
                throw new CodecException(
                    "ppj.nativeRef.stale",
                    "A source-bound placeholder edit requires its hash-bound nativeRef.",
                    placeholderPath + ".nativeRef");
            if (frameChanged) RequireCapability(newPlaceholder.NativeRef, "setFrame", placeholderPath + ".frame");
            if (textChanged) RequireCapability(newPlaceholder.NativeRef, "replaceText", placeholderPath + ".text");
            if (styleChanged) RequireCapabilityField(newPlaceholder.NativeRef, "setTextBodyStyle", "text.style", placeholderPath + ".style");

            var nativeType = newPlaceholder.PlaceholderType switch
            {
                "centered-title" => "ctrTitle",
                "subtitle" => "subTitle",
                "content" => "obj",
                "picture" => "pic",
                "table" => "tbl",
                "date" => "dt",
                "footer" => "ftr",
                "slide-number" => "sldNum",
                _ => newPlaceholder.PlaceholderType,
            };
            var targetIndex = -1;
            for (var targetOrdinal = 0; targetOrdinal < target.Count; targetOrdinal++)
            {
                var candidate = target[targetOrdinal];
                if (candidate.Type.Equals(nativeType, StringComparison.Ordinal) && candidate.Index == newPlaceholder.Index)
                {
                    if (targetIndex >= 0)
                        throw Unsupported(placeholderPath, "source placeholder type/index is ambiguous");
                    targetIndex = targetOrdinal;
                }
            }
            if (targetIndex < 0)
                throw Unsupported(placeholderPath, "source placeholder no longer resolves to its owner-local native shape");

            var requestedWire = target[targetIndex].Clone();
            if (frameChanged)
            {
                var frame = newPlaceholder.Raw.GetProperty("frame");
                requestedWire.DirectFrame = new PresentationPlaceholderFrame
                {
                    LeftEmu = Emu(newPlaceholder.Frame.X),
                    TopEmu = Emu(newPlaceholder.Frame.Y),
                    WidthEmu = Emu(newPlaceholder.Frame.Width),
                    HeightEmu = Emu(newPlaceholder.Frame.Height),
                };
                // Preserve optional transform presence independently from its
                // value.  Omitting rotation/flip in the requested PPJ frame
                // is an explicit removal of that native attribute; keeping a
                // present zero/false value remains a present native leaf.
                if (frame.TryGetProperty("rotation", out _))
                    requestedWire.DirectFrame.RotationAngle60000 = RotationAngle(newPlaceholder.Frame.Rotation);
                if (frame.TryGetProperty("flipH", out var flipH))
                    requestedWire.DirectFrame.FlipHorizontal = flipH.GetBoolean();
                if (frame.TryGetProperty("flipV", out var flipV))
                    requestedWire.DirectFrame.FlipVertical = flipV.GetBoolean();
            }
            if (textChanged)
            {
                if (!oldPlaceholder.Raw.TryGetProperty("text", out _) ||
                    !newPlaceholder.Raw.TryGetProperty("text", out var text))
                    throw Unsupported(placeholderPath + ".text", "source-bound placeholder text must retain an existing text owner");
                requestedWire.TextBody = PpjAuthoredPresentationCompiler.BuildSourceBoundTextBody(text, programRoot);
            }
            if (styleChanged)
            {
                if (!newPlaceholder.Raw.TryGetProperty("style", out var style))
                    throw Unsupported(placeholderPath + ".style", "removing source-bound placeholder text body style is not an explicit bounded operation");
                PpjAuthoredPresentationCompiler.MergeSourceBoundTextBodyStyle(
                    requestedWire.TextBody ?? throw Unsupported(placeholderPath + ".style", "source-bound placeholder text body style requires an existing text body"),
                    style,
                    placeholderPath + ".style");
            }
            target[targetIndex] = requestedWire;
            changed = true;
        }
        return changed;
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
        string path,
        Func<JsonElement, double>? resolveOpacity = null,
        Func<JsonElement, string>? resolveFit = null) => PpjImagePaintLowering.Build(
        fill,
        frameWidth,
        frameHeight,
        id => ResolveAsset(assets, id, path + ".asset"),
        id => assetDimensions.TryGetValue(id, out var dimensions) ? dimensions : null,
        path,
        resolveOpacity: resolveOpacity,
        resolveFit: resolveFit);

    private static JsonElement? EffectiveImageProperty(
        JsonElement raw,
        JsonElement? inlineStyle,
        JsonElement? namedStyle,
        PpjAuthoredPresentationCompiler.Catalog catalog,
        string field) => catalog.PropertyByPrecedence(
            $"image.{field}",
            raw,
            inlineStyle,
            namedStyle,
            includeElementWhenUndeclared: true);

    private static string EffectiveImageFit(
        JsonElement raw,
        JsonElement? inlineStyle,
        JsonElement? namedStyle,
        PpjAuthoredPresentationCompiler.Catalog catalog,
        string path)
    {
        var value = EffectiveImageProperty(raw, inlineStyle, namedStyle, catalog, "fit");
        if (value is null) return "stretch";
        return value.Value.ValueKind == JsonValueKind.String
            ? value.Value.GetString()!
            : catalog.StringToken(value.Value, "string", path);
    }

    private static double? EffectiveImageOpacity(
        JsonElement raw,
        JsonElement? inlineStyle,
        JsonElement? namedStyle,
        PpjAuthoredPresentationCompiler.Catalog catalog,
        string path)
    {
        var value = EffectiveImageProperty(raw, inlineStyle, namedStyle, catalog, "opacity");
        if (value is null) return null;
        var opacity = catalog.NumberToken(value.Value, "opacity", path);
        if (opacity is < 0 or > 1)
            throw new CodecException("ppj.opacity", $"PPJ {path} must be between 0 and 1.", path);
        return opacity;
    }

    private static JsonElement EffectiveImagePaintSource(
        PpjImageElementModel element,
        JsonElement raw,
        JsonElement? inlineStyle,
        JsonElement? namedStyle,
        PpjAuthoredPresentationCompiler.Catalog catalog)
    {
        var output = new JsonObject
        {
            ["asset"] = JsonValue.Create(element.AssetId),
        };
        foreach (var field in new[] { "fit", "crop", "focus", "opacity" })
        {
            if (EffectiveImageProperty(raw, inlineStyle, namedStyle, catalog, field) is not { } value)
                continue;
            output[field] = JsonNode.Parse(value.GetRawText());
        }
        using var document = JsonDocument.Parse(output.ToJsonString());
        return document.RootElement.Clone();
    }

    private static string ResolveGrammarStringToken(JsonElement root, JsonElement value, string path)
    {
        if (value.ValueKind == JsonValueKind.String)
            return value.GetString()!;
        if (value.ValueKind != JsonValueKind.Object ||
            !value.TryGetProperty("token", out var tokenValue) ||
            tokenValue.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} must be a string or a string grammar token.", path);
        var token = tokenValue.GetString()!;
        var design = root.GetProperty("design");
        if (!design.TryGetProperty("grammar", out var grammar) ||
            !grammar.TryGetProperty("tokens", out var tokens) ||
            !tokens.TryGetProperty(token, out var definition) ||
            definition.ValueKind != JsonValueKind.Object)
            throw new CodecException("ppj.grammar.tokenUnknown", $"PPJ grammar token {token} for {path} is not declared.", path);
        if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "string")
            throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {token} for {path} must declare kind string.", path);
        if (!definition.TryGetProperty("value", out var tokenText) ||
            tokenText.ValueKind != JsonValueKind.String ||
            string.IsNullOrEmpty(tokenText.GetString()))
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token {token} for {path} must resolve to a non-empty string.", path);
        return tokenText.GetString()!;
    }

    private static string ResolveGrammarEnumToken(
        JsonElement root,
        JsonElement value,
        string path,
        params string[] allowed)
    {
        var resolved = ResolveGrammarStringToken(root, value, path);
        if (!allowed.Contains(resolved, StringComparer.Ordinal))
            throw new CodecException(
                "ppj.grammar.tokenValue",
                $"PPJ {path} resolved to unsupported value {resolved}.",
                path);
        return resolved;
    }

    private static double ResolveGrammarOpacityToken(JsonElement root, JsonElement value, string path)
    {
        return ValidateOpacity(ResolveGrammarNumberToken(root, value, "opacity", path), path);
    }

    private static double ResolveGrammarSizeToken(JsonElement root, JsonElement value, string path)
    {
        var resolved = ResolveGrammarNumberToken(root, value, "size", path);
        if (!double.IsFinite(resolved) || resolved < 0 || resolved > 1_000)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} size must be finite and between 0 and 1000.", path);
        return resolved;
    }

    private static double ResolveGrammarPositiveSizeToken(JsonElement root, JsonElement value, string path)
    {
        var resolved = ResolveGrammarNumberToken(root, value, "size", path);
        if (!double.IsFinite(resolved) || resolved <= 0 || resolved > 1_000)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} size must be finite and greater than 0 and at most 1000.", path);
        return resolved;
    }

    private static double ResolveGrammarPositiveNumberToken(JsonElement root, JsonElement value, string path)
    {
        var resolved = ResolveGrammarNumberToken(root, value, "size", path);
        if (!double.IsFinite(resolved) || resolved <= 0)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} must resolve to a finite number greater than 0.", path);
        return resolved;
    }

    private static uint ResolveGrammarIntegerToken(
        JsonElement root,
        JsonElement value,
        string path,
        uint minimum,
        uint maximum)
    {
        var resolved = ResolveGrammarNumberToken(root, value, "size", path);
        if (!double.IsFinite(resolved) || resolved < minimum || resolved > maximum ||
            Math.Truncate(resolved) != resolved)
            throw new CodecException(
                "ppj.grammar.tokenValue",
                $"PPJ {path} must resolve to an integer between {minimum} and {maximum}.",
                path);
        return checked((uint)resolved);
    }

    private static int ResolveGrammarSignedIntegerToken(
        JsonElement root,
        JsonElement value,
        string path,
        int minimum,
        int maximum)
    {
        var resolved = ResolveGrammarNumberToken(root, value, "size", path);
        if (!double.IsFinite(resolved) || resolved < minimum || resolved > maximum || Math.Truncate(resolved) != resolved)
            throw new CodecException(
                "ppj.grammar.tokenValue",
                $"PPJ {path} must resolve to an integer between {minimum} and {maximum}.",
                path);
        return checked((int)resolved);
    }

    private static bool ResolveGrammarBooleanToken(JsonElement root, JsonElement value, string path)
    {
        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        if (value.ValueKind != JsonValueKind.Object ||
            !value.TryGetProperty("token", out var tokenValue) ||
            tokenValue.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} must be a boolean or a boolean grammar token.", path);
        var token = tokenValue.GetString()!;
        var design = root.GetProperty("design");
        if (!design.TryGetProperty("grammar", out var grammar) ||
            !grammar.TryGetProperty("tokens", out var tokens) ||
            !tokens.TryGetProperty(token, out var definition) ||
            definition.ValueKind != JsonValueKind.Object)
            throw new CodecException("ppj.grammar.tokenUnknown", $"PPJ grammar token {token} for {path} is not declared.", path);
        if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "boolean")
            throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {token} for {path} must declare kind boolean.", path);
        if (!definition.TryGetProperty("value", out var tokenBoolean) ||
            tokenBoolean.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token {token} for {path} must resolve to a boolean.", path);
        return tokenBoolean.GetBoolean();
    }

    private static (string Rgb, double Alpha) ResolveGrammarColorValue(JsonElement root, JsonElement value, string path)
    {
        var (rgb, alpha) = value.ValueKind == JsonValueKind.String
            ? ParseSourceBoundColor(value, path)
            : ResolveGrammarColorToken(root, value, path);
        if (value.ValueKind == JsonValueKind.Object)
        {
            if (value.TryGetProperty("tint", out var tint))
            {
                var amount = tint.GetDouble();
                if (!double.IsFinite(amount) || amount < 0 || amount > 1)
                    throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path}.tint must be finite and between 0 and 1.", path + ".tint");
                rgb = MixRgb(rgb, "FFFFFF", amount);
            }
            if (value.TryGetProperty("shade", out var shade))
            {
                var amount = shade.GetDouble();
                if (!double.IsFinite(amount) || amount < 0 || amount > 1)
                    throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path}.shade must be finite and between 0 and 1.", path + ".shade");
                rgb = MixRgb(rgb, "000000", amount);
            }
            if (value.TryGetProperty("alpha", out var explicitAlpha))
                alpha = ValidateOpacity(explicitAlpha.GetDouble(), path + ".alpha");
        }
        return (rgb, alpha);
    }

    private static (string Rgb, double Alpha) ResolveGrammarColorToken(JsonElement root, JsonElement value, string path)
    {
        if (value.ValueKind != JsonValueKind.Object ||
            !value.TryGetProperty("token", out var tokenValue) ||
            tokenValue.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} must be an RGB color or a color grammar token.", path);
        var token = tokenValue.GetString()!;
        var design = root.GetProperty("design");
        if (!design.TryGetProperty("grammar", out var grammar) ||
            !grammar.TryGetProperty("tokens", out var tokens) ||
            !tokens.TryGetProperty(token, out var definition) ||
            definition.ValueKind != JsonValueKind.Object)
            throw new CodecException("ppj.grammar.tokenUnknown", $"PPJ grammar token {token} for {path} is not declared.", path);
        if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
            throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {token} for {path} must declare kind color.", path);
        if (!definition.TryGetProperty("value", out var tokenColor) || tokenColor.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ color token {token} for {path} must resolve to a color string.", path);
        return ParseSourceBoundColor(tokenColor, path);
    }

    private static (string Rgb, double Alpha) ParseSourceBoundColor(JsonElement value, string path)
    {
        if (value.ValueKind != JsonValueKind.String)
            throw Unsupported(path, "theme or computed source-bound color");
        var normalized = value.GetString()!.TrimStart('#');
        if (normalized.Length is not (6 or 8) || !normalized.All(Uri.IsHexDigit))
            throw Unsupported(path, "non-RGB chart text color");
        var alpha = normalized.Length == 8
            ? Convert.ToByte(normalized[6..], 16) / 255d
            : 1d;
        return (normalized[..6].ToUpperInvariant(), alpha);
    }

    private static string MixRgb(string source, string target, double amount)
    {
        var t = Math.Clamp(amount, 0, 1);
        var r = MixRgbChannel(source[..2], target[..2], t);
        var g = MixRgbChannel(source[2..4], target[2..4], t);
        var b = MixRgbChannel(source[4..6], target[4..6], t);
        return $"{r:X2}{g:X2}{b:X2}";
    }

    private static byte MixRgbChannel(string source, string target, double amount)
    {
        var from = byte.Parse(source, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture);
        var to = byte.Parse(target, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture);
        return checked((byte)Math.Round(from + (to - from) * amount, MidpointRounding.AwayFromZero));
    }

    private static double ResolveGrammarNumberToken(JsonElement root, JsonElement value, string expectedKind, string path)
    {
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var literal) && double.IsFinite(literal))
            return literal;
        if (value.ValueKind != JsonValueKind.Object ||
            !value.TryGetProperty("token", out var tokenValue) ||
            tokenValue.ValueKind != JsonValueKind.String)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} must be a finite number or a {expectedKind} grammar token.", path);
        var token = tokenValue.GetString()!;
        var design = root.GetProperty("design");
        if (!design.TryGetProperty("grammar", out var grammar) ||
            !grammar.TryGetProperty("tokens", out var tokens) ||
            !tokens.TryGetProperty(token, out var definition) ||
            definition.ValueKind != JsonValueKind.Object)
            throw new CodecException("ppj.grammar.tokenUnknown", $"PPJ grammar token {token} for {path} is not declared.", path);
        if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != expectedKind)
            throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {token} for {path} must declare kind {expectedKind}.", path);
        if (!definition.TryGetProperty("value", out var tokenNumber) ||
            tokenNumber.ValueKind != JsonValueKind.Number ||
            !tokenNumber.TryGetDouble(out var resolved) ||
            !double.IsFinite(resolved))
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ grammar token {token} for {path} must resolve to a number.", path);
        return resolved;
    }

    private static bool IsGrammarTokenReference(JsonElement value) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty("token", out var token) &&
        token.ValueKind == JsonValueKind.String;

    private static double ValidateOpacity(double value, string path)
    {
        if (!double.IsFinite(value) || value < 0 || value > 1)
            throw new CodecException("ppj.grammar.tokenValue", $"PPJ {path} opacity must be finite and between 0 and 1.", path);
        return value;
    }

    private static PresentationBackground? BuildBackground(
        JsonElement fill,
        double canvasWidth,
        double canvasHeight,
        Func<string, string> resolveAsset,
        Func<string, (double Width, double Height)?> assetDimensions,
        string path,
        Func<JsonElement, double>? resolveOpacity = null,
        Func<JsonElement, string>? resolveFit = null,
        Func<JsonElement, string, (string Rgb, double Alpha)>? resolveColor = null)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return null;
        if (type == "image") return new PresentationBackground
        {
            ImagePaint = PpjImagePaintLowering.Build(
                fill,
                canvasWidth,
                canvasHeight,
                resolveAsset,
                assetDimensions,
                path,
                resolveOpacity: resolveOpacity,
                resolveFit: resolveFit),
        };
        if (type == "gradient") return new PresentationBackground { GradientFill = BuildGradientFill(fill, path, resolveColor) };
        if (type == "solid")
        {
            var color = fill.GetProperty("color");
            (string Rgb, double Alpha) resolvedColor = resolveColor is not null
                ? resolveColor(color, path + ".color")
                : (Rgb(color, path + ".color"), 1d);
            var background = new PresentationBackground
            {
                Solid = true,
                ColorRgb = resolvedColor.Rgb,
            };
            var opacity = resolvedColor.Alpha;
            if (fill.TryGetProperty("opacity", out var opacityValue))
            {
                opacity = resolveOpacity is not null
                    ? resolveOpacity(opacityValue)
                    : opacityValue.GetDouble();
            }
            if (opacity != 1)
                background.OpacityThousandthPercent = checked((uint)Math.Round(ValidateOpacity(opacity, path + ".opacity") * 100_000d, MidpointRounding.AwayFromZero));
            return background;
        }
        throw Unsupported(path, $"background fill {type}");
    }

    private static bool ApplyCanvas(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        ISet<string> changedNodeIds,
        MutationState mutations)
    {
        if (baseline.Design.Width.Equals(requested.Design.Width) &&
            baseline.Design.Height.Equals(requested.Design.Height))
            return false;

        var before = baseline.Root.GetProperty("design").GetProperty("canvas");
        var after = requested.Root.GetProperty("design").GetProperty("canvas");
        RequireNativeRef(before, after, "$.design.canvas");
        RequireEqualExcept(before, after, "$.design.canvas", "width", "height");
        RequireCapability(requested.Design.CanvasNativeRef, "setCanvas", "$.design.canvas");

        presentation.SlideWidthEmu = Emu(requested.Design.Width);
        presentation.SlideHeightEmu = Emu(requested.Design.Height);
        foreach (var page in requested.Pages) changedNodeIds.Add(page.Id);
        mutations.SemanticChanges = true;
        return true;
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
        if (baseline.Pages.Count != presentation.Slides.Count)
            throw Unsupported("$.pages", "inconsistent source topology");
        var sourcePages = baseline.Pages.Select((page, index) => new
        {
            Program = page,
            Wire = presentation.Slides[index],
        }).ToDictionary(item => item.Program.Id, StringComparer.Ordinal);
        var requestedPageIds = requested.Pages.Select(page => page.Id).ToArray();
        if (requestedPageIds.Distinct(StringComparer.Ordinal).Count() != requestedPageIds.Length)
            throw Unsupported("$.pages", "duplicate page identity");
        var pendingClones = requested.Pages.Where(page => page.SourceClone is not null).ToArray();
        var retainedPages = requested.Pages.Where(page => page.SourceClone is null).ToArray();
        if (retainedPages.Any(page => !sourcePages.ContainsKey(page.Id)))
            throw Unsupported("$.pages", "source-bound page insertion or identity change");
        if (pendingClones.Length > 0)
        {
            if (!retainedPages.Select(page => page.Id).SequenceEqual(baseline.Pages.Select(page => page.Id), StringComparer.Ordinal))
                throw Unsupported("$.pages", "combining source slide reuse with page deletion or reorder");
            if (requested.Pages.Count != baseline.Pages.Count + pendingClones.Length)
                throw Unsupported("$.pages", "inconsistent source clone topology");
        }
        else if (requested.Pages.Count > baseline.Pages.Count)
        {
            throw Unsupported("$.pages", "source-bound page insertion without sourceClone");
        }
        var retainedPageSourceOrder = baseline.Pages
            .Where(page => retainedPages.Any(requestedPage => requestedPage.Id.Equals(page.Id, StringComparison.Ordinal)))
            .Select(page => page.Id)
            .ToArray();
        var retainedRequestedIds = retainedPages.Select(page => page.Id).ToArray();
        var pagesReordered = pendingClones.Length == 0 &&
            !retainedPageSourceOrder.SequenceEqual(retainedRequestedIds, StringComparer.Ordinal);
        if (pagesReordered)
        {
            if (requested.Pages.Count != baseline.Pages.Count)
                throw Unsupported("$.pages", "combining source-bound page deletion with page reorder");
            for (var index = 0; index < requested.Pages.Count; index++)
            {
                if (retainedPageSourceOrder[index].Equals(requestedPageIds[index], StringComparison.Ordinal)) continue;
                RequireCapability(requested.Pages[index].NativeRef, "reorder", $"$.pages[{index}]");
                changedNodeIds.Add(requested.Pages[index].Id);
            }
            mutations.SemanticChanges = true;
        }

        var changed = pagesReordered;
        var assetDimensions = AssetDimensions(requested.Root);
        var requestedSlides = new List<PresentationSlide>(requested.Pages.Count);
        var clonedSourcePageIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < requested.Pages.Count; index++)
        {
            var after = requested.Pages[index];
            var path = $"$.pages[{index}]";
            if (after.SourceClone is not null)
            {
                var sourceClone = after.SourceClone;
                if (!sourcePages.TryGetValue(sourceClone.PageId, out var cloneOrigin))
                    throw Unsupported(path + ".sourceClone.page", $"unknown source page {sourceClone.PageId}");
                if (index == 0 || !requested.Pages[index - 1].Id.Equals(sourceClone.PageId, StringComparison.Ordinal))
                    throw Unsupported(path + ".sourceClone.page", "a pending source clone must immediately follow its retained source page");
                if (!clonedSourcePageIds.Add(sourceClone.PageId))
                    throw Unsupported(path + ".sourceClone.page", "more than one pending clone for the same source page");
                RequireSourceClonePage(after, path);
                RequireCapability(
                    cloneOrigin.Program.NativeRef,
                    "duplicate",
                    sourceClone.CapabilityId,
                    "pageClone",
                    path + ".sourceClone.capability");
                if (cloneOrigin.Wire.Source?.CloneCapability?.Supported != true)
                    throw Unsupported(path + ".sourceClone", "source slide clone graph is no longer supported");

                var clone = cloneOrigin.Wire.Clone();
                clone.Id = after.Id;
                clone.CloneSource = cloneOrigin.Wire.Source.Clone();
                clone.Source = null;
                clone.ElementDeletions.Clear();
                if (sourceClone.RetainElementId is { } retainedElementId)
                {
                    if (cloneOrigin.Program.Elements.Count != clone.Elements.Count)
                        throw Unsupported(path + ".sourceClone.retainElement", "inconsistent source element topology");
                    var retainedIndex = -1;
                    for (var elementIndex = 0; elementIndex < cloneOrigin.Program.Elements.Count; elementIndex++)
                    {
                        if (!cloneOrigin.Program.Elements[elementIndex].Id.Equals(retainedElementId, StringComparison.Ordinal)) continue;
                        retainedIndex = elementIndex;
                        break;
                    }
                    if (retainedIndex < 0)
                        throw Unsupported(path + ".sourceClone.retainElement", $"unknown direct source element {retainedElementId}");

                    var retainedWire = clone.Elements[retainedIndex].Clone();
                    for (var elementIndex = 0; elementIndex < cloneOrigin.Program.Elements.Count; elementIndex++)
                    {
                        if (elementIndex == retainedIndex) continue;
                        var sibling = cloneOrigin.Program.Elements[elementIndex];
                        var siblingWire = clone.Elements[elementIndex];
                        RequireCapabilityField(sibling, "delete", "element", path + ".sourceClone.retainElement");
                        if (siblingWire.Source?.DeletionCapability?.Supported != true)
                            throw Unsupported(path + ".sourceClone.retainElement", $"deleting sibling {sibling.Id} without a re-proven native deletion profile");
                        clone.ElementDeletions.Add(new PresentationElementDeletion
                        {
                            Id = siblingWire.Id,
                            Source = siblingWire.Source.Clone(),
                        });
                    }
                    clone.Elements.Clear();
                    clone.Elements.Add(retainedWire);
                }
                requestedSlides.Add(clone);
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
                continue;
            }

            var sourcePage = sourcePages[after.Id];
            var before = sourcePage.Program;
            var slide = sourcePage.Wire;
            RequireNativeRef(before.Raw, after.Raw, path);
            RequireEqualExcept(before.Raw, after.Raw, path, "name", "role", "claim", "background", "transition", "notes", "animations", "elements", "readingOrder", "hidden");
            var animationsChanged = AnimationStateChanged(before, after);
            if (PropertyChanged(before.Raw, after.Raw, "name"))
            {
                RequireCapability(after.NativeRef, "setName", path + ".name");
                slide.Name = after.Name ?? string.Empty;
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (PropertyChanged(before.Raw, after.Raw, "hidden"))
            {
                RequireCapability(after.NativeRef, "setHidden", path + ".hidden");
                if (after.Hidden is null)
                    throw Unsupported(path + ".hidden", "removing an issued slide visibility state; set an explicit boolean instead");
                slide.Hidden = after.Hidden.Value;
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
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
                        path + ".background",
                        resolveOpacity: opacity => ResolveGrammarOpacityToken(requested.Root, opacity, path + ".background.opacity"),
                        resolveFit: fit => ResolveGrammarStringToken(requested.Root, fit, path + ".background.fit"),
                        resolveColor: (color, colorPath) => ResolveGrammarColorValue(requested.Root, color, colorPath))
                    : null;
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (PropertyChanged(before.Raw, after.Raw, "transition"))
            {
                RequireCapability(after.NativeRef, "setTransition", path + ".transition");
                if (before.Transition?.Type == "morph" || after.Transition?.Type == "morph")
                    throw Unsupported(path + ".transition", "source-bound Morph mutation requires a dedicated paired-object capability");
                slide.Transition = after.Transition is null || after.Transition.Type == "none"
                    ? null
                    : PpjTransitionLowering.BuildBase(after.Transition);
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (PropertyChanged(before.Raw, after.Raw, "notes"))
            {
                RequireCapability(after.NativeRef, "setNotes", path + ".notes");
                ApplySpeakerNotes(before, after, slide, path + ".notes");
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (before.Elements.Count != slide.Elements.Count)
                throw Unsupported(path + ".elements", "inconsistent source topology");
            var sourceElements = before.Elements.Select((element, elementIndex) => new
            {
                Program = element,
                Wire = slide.Elements[elementIndex],
            }).ToDictionary(item => item.Program.Id, StringComparer.Ordinal);
            var requestedIds = after.Elements.Select(element => element.Id).ToArray();
            if (requestedIds.Distinct(StringComparer.Ordinal).Count() != requestedIds.Length)
                throw Unsupported(path + ".elements", "duplicate element identity");
            var requestedElementOrder = after.ReadingOrder.Count == 0
                ? requestedIds
                : after.ReadingOrder.ToArray();
            if (after.ReadingOrder.Count > 0)
            {
                var requestedIdSet = requestedIds.ToHashSet(StringComparer.Ordinal);
                var sourceIdSet = sourcePages[after.Id].Program.Elements
                    .Select(element => element.Id)
                    .ToHashSet(StringComparer.Ordinal);
                var current = new List<string>(requestedElementOrder.Length);
                var staleOnly = true;
                foreach (var id in requestedElementOrder)
                {
                    if (requestedIdSet.Contains(id)) current.Add(id);
                    else if (!sourceIdSet.Contains(id)) staleOnly = false;
                }
                var missing = requestedIds.Where(id => !current.Contains(id, StringComparer.Ordinal)).ToArray();
                var missingAreAuthored = missing.All(id => !sourceIdSet.Contains(id));
                if (staleOnly && missingAreAuthored)
                {
                    // The projector emits the source page's direct order. If
                    // a request deletes a source element, filter that proven
                    // stale ID; if it appends an overlay, append the newly
                    // authored IDs in their requested topmost order.
                    requestedElementOrder = current
                        .Concat(missing)
                        .ToArray();
                }
            }
            if (requestedElementOrder.Length != requestedIds.Length ||
                requestedElementOrder.Distinct(StringComparer.Ordinal).Count() != requestedElementOrder.Length ||
                !requestedElementOrder.ToHashSet(StringComparer.Ordinal).SetEquals(requestedIds))
                throw Unsupported(path + ".readingOrder", "reading order must be a complete permutation of the direct source elements");
            var retainedCount = 0;
            while (retainedCount < after.Elements.Count && sourceElements.ContainsKey(after.Elements[retainedCount].Id))
                retainedCount++;
            var retained = after.Elements.Take(retainedCount).ToArray();
            var authored = after.Elements.Skip(retainedCount).ToArray();
            if (authored.Any(element => sourceElements.ContainsKey(element.Id)))
                throw Unsupported(path + ".elements", "new overlays must remain a topmost suffix after the complete source-owned prefix");

            if (authored.Length > 0)
            {
                RequireCapability(after.NativeRef, "appendElement", path + ".elements");
                if (animationsChanged)
                    throw Unsupported(path + ".animations", "animating a new source overlay before build and reimport");
                if (mutations.SemanticChanges || mutations.NativeLeaves.Count > 0 || changedNodeIds.Contains(after.Id))
                    throw Unsupported(path + ".elements", "mixing a source overlay with another mutation; build and reimport between edits");
                if (retained.Length != before.Elements.Count ||
                    !retained.Select(element => element.Id).SequenceEqual(before.Elements.Select(element => element.Id), StringComparer.Ordinal))
                    throw Unsupported(path + ".elements", "an overlay requires the complete source element prefix in its original order");
                for (var elementIndex = 0; elementIndex < retained.Length; elementIndex++)
                {
                    if (!JsonEqual(before.Elements[elementIndex].Raw, retained[elementIndex].Raw))
                        throw Unsupported($"{path}.elements[{elementIndex}]", "an overlay cannot be combined with a source element edit");
                }

                var currentOverlayNodeIds = new HashSet<string>(StringComparer.Ordinal) { after.Id };
                var overlayWire = slide.Elements.Select(element => element.Clone()).ToList();
                foreach (var element in authored)
                {
                    overlayWire.Add(PpjAuthoredPresentationCompiler.BuildSourceBoundOverlayElement(requested, element));
                    currentOverlayNodeIds.Add(element.Id);
                    changedNodeIds.Add(element.Id);
                }
                slide.Elements.Clear();
                slide.Elements.Add(overlayWire);
                changedNodeIds.Add(after.Id);
                mutations.AuthoredOverlayNodeIds = currentOverlayNodeIds;
                mutations.SemanticChanges = true;
                changed = true;
                requestedSlides.Add(slide);
                continue;
            }

            if (requestedIds.Any(id => !sourceElements.ContainsKey(id)))
                throw Unsupported(path + ".elements", "source-bound element identity change");

            var requestedWireById = new Dictionary<string, PresentationElement>(StringComparer.Ordinal);
            for (var elementIndex = 0; elementIndex < after.Elements.Count; elementIndex++)
            {
                var elementPath = $"{path}.elements[{elementIndex}]";
                var beforeElement = sourceElements[after.Elements[elementIndex].Id];
                var wireElement = beforeElement.Wire;
                var shapeTreePath = new[] { wireElement.Source?.ShapeTreeIndex ?? checked((uint)elementIndex) };
                if (ApplyElement(requested, beforeElement.Program, after.Elements[elementIndex], wireElement, slide, shapeTreePath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, elementPath))
                {
                    changed = true;
                    changedNodeIds.Add(after.Id);
                }
                requestedWireById[after.Elements[elementIndex].Id] = wireElement;
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
            if (!retainedSourceOrder.SequenceEqual(requestedElementOrder, StringComparer.Ordinal))
            {
                for (var elementIndex = 0; elementIndex < requestedElementOrder.Length; elementIndex++)
                {
                    if (retainedSourceOrder[elementIndex] == requestedElementOrder[elementIndex]) continue;
                    var requestedElement = after.Elements.Single(element => element.Id.Equals(requestedElementOrder[elementIndex], StringComparison.Ordinal));
                    RequireCapability(requestedElement, "reorder", $"{path}.readingOrder[{elementIndex}]");
                    changedNodeIds.Add(requestedElement.Id);
                }
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            if (animationsChanged)
            {
                RequireCapabilityField(after.NativeRef, "setAnimations", "animations", path + ".animations");
                if (slide.Morph is not null ||
                    slide.Source is null ||
                    slide.Source.TimingEditable != true && slide.Source.TimingAddable != true)
                    throw Unsupported(path + ".animations", "the source timing graph is no longer editable or addable");

                var animationWire = requestedElementOrder.Select(id => requestedWireById[id]).ToList();
                var targetIds = SourceAnimationTargetIds(after.Elements, animationWire, path + ".elements");
                slide.Animations.Clear();
                foreach (var animation in after.Animations)
                {
                    var lowered = PpjAuthoredPresentationCompiler.BuildAnimation(animation, after.Elements);
                    if (!targetIds.TryGetValue(animation.TargetId, out var nativeTargetId))
                        throw Unsupported(path + ".animations", $"animation target {animation.TargetId} has no source wire binding");
                    lowered.TargetId = nativeTargetId;
                    slide.Animations.Add(lowered);
                    changedNodeIds.Add(animation.Id);
                }
                foreach (var animation in before.Animations) changedNodeIds.Add(animation.Id);
                changedNodeIds.Add(after.Id);
                mutations.SemanticChanges = true;
                changed = true;
            }
            var orderedRequestedWire = requestedElementOrder.Select(id => requestedWireById[id]).ToList();
            slide.Elements.Clear();
            slide.Elements.Add(orderedRequestedWire);
            slide.ReadingOrder.Clear();
            slide.ReadingOrder.Add(orderedRequestedWire.Select(element => element.Id));
            requestedSlides.Add(slide);
        }

        if (mutations.AuthoredOverlayNodeIds is { } overlayNodeIds &&
            changedNodeIds.Any(id => !overlayNodeIds.Contains(id)))
            throw Unsupported("$.pages", "mixing a source overlay with another page mutation; build and reimport between edits");

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

    private static void RejectMixedAuthoredOverlay(MutationState mutations, string path)
    {
        if (mutations.AuthoredOverlayNodeIds is not null)
            throw Unsupported(path, "mixing source overlay creation with another source transaction; build and reimport between edits");
    }

    private static void RequireSourceClonePage(PpjPageModel page, string path)
    {
        if (page.NativeRef is not null || page.Elements.Count != 0 || page.Name is not null ||
            page.LayoutId is not null || page.Notes is not null || page.Hidden is not null ||
            page.Transition is not null || page.Animations.Count != 0 || page.Raw.TryGetProperty("background", out _))
            throw Unsupported(path, "editing a pending source clone before build and reimport");
        var allowed = new HashSet<string>(["id", "role", "claim", "elements", "sourceClone"], StringComparer.Ordinal);
        foreach (var property in page.Raw.EnumerateObject())
            if (!allowed.Contains(property.Name))
                throw Unsupported(path + "." + property.Name, "declaring source-owned state on a pending source clone");
    }

    private static bool ApplyComments(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        ISet<string> changedNodeIds)
    {
        if (baseline.Comments.Count != requested.Comments.Count)
            throw Unsupported("$.comments", "adding or removing source-bound comments");
        if (baseline.Comments.Count == 0) return false;
        if (presentation.Slides.Count != requested.Pages.Count)
            throw Unsupported("$.comments", "comment editing with changed page topology");

        var slideByPageId = requested.Pages.Select((page, index) => new
        {
            page.Id,
            Slide = presentation.Slides[index],
        }).ToDictionary(item => item.Id, item => item.Slide, StringComparer.Ordinal);
        var sourceComments = new Dictionary<string, SourceCommentTarget>(StringComparer.Ordinal);
        foreach (var pageGroup in baseline.Comments.GroupBy(comment => comment.PageId, StringComparer.Ordinal))
        {
            if (!slideByPageId.TryGetValue(pageGroup.Key, out var slide))
                throw Unsupported("$.comments", "comments attached to a removed or unknown page");
            var projected = pageGroup.ToArray();
            var modern = slide.ModernComments
                .SelectMany(thread => new[] { (Thread: thread, Comment: thread.Root, IsRoot: true) }
                    .Concat(thread.Replies.Select(reply => (Thread: thread, Comment: reply, IsRoot: false))))
                .Where(item => item.Comment is not null)
                .Select(item => (item.Thread, Comment: item.Comment!, item.IsRoot))
                .ToArray();
            if (projected.Length != slide.LegacyComments.Count + modern.Length)
                throw Unsupported("$.comments", "comment projection no longer matches the fresh source slide");
            var cursor = 0;
            for (var index = 0; index < slide.LegacyComments.Count; index++)
            {
                var comment = projected[cursor++];
                if (comment.Kind != "legacy")
                    throw Unsupported("$.comments", "legacy and modern comment order or family changed");
                sourceComments.Add(comment.Id, new SourceCommentTarget(slide, slide.LegacyComments[index], null, null));
            }
            foreach (var item in modern)
            {
                var comment = projected[cursor++];
                if (comment.Kind != "modern" || (item.IsRoot ? comment.ParentId is not null : comment.ParentId is null))
                    throw Unsupported("$.comments", "modern comment thread order or parent topology changed");
                sourceComments.Add(comment.Id, new SourceCommentTarget(slide, null, item.Thread, item.Comment));
            }
        }

        var changed = false;
        for (var index = 0; index < baseline.Comments.Count; index++)
        {
            var before = baseline.Comments[index];
            var after = requested.Comments[index];
            var path = $"$.comments[{index}]";
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal) ||
                !before.PageId.Equals(after.PageId, StringComparison.Ordinal) ||
                !before.Kind.Equals(after.Kind, StringComparison.Ordinal) ||
                !sourceComments.TryGetValue(before.Id, out var target))
                throw Unsupported(path, "comment reorder, identity, or page change");
            RequireNativeRef(before.Raw, after.Raw, path);
            if (before.Kind == "legacy")
            {
                RequireEqualExcept(before.Raw, after.Raw, path, "text");
                if (before.Text.Equals(after.Text, StringComparison.Ordinal)) continue;
                RequireCapability(after.NativeRef, "replaceText", path + ".text");
                target.Legacy!.Text = after.Text;
                changedNodeIds.Add(after.Id);
                changedNodeIds.Add(after.PageId);
                changed = true;
                continue;
            }

            if (target.Modern is null || target.ModernThread is null)
                throw Unsupported(path, "modern comment source binding is missing");
            RequireEqualExcept(before.Raw, after.Raw, path, "text", "resolved", "status");
            var textChanged = !before.Text.Equals(after.Text, StringComparison.Ordinal);
            var resolvedChanged = before.Resolved != after.Resolved;
            var statusChanged = !string.Equals(before.Status, after.Status, StringComparison.Ordinal);
            if (!textChanged && !resolvedChanged && !statusChanged) continue;
            if (textChanged)
                RequireCapability(after.NativeRef, "replaceText", path + ".text");
            if (resolvedChanged || statusChanged)
            {
                RequireCapabilityField(after.NativeRef, "setCommentStatus", "status", path + ".status");
                RequireCapabilityField(after.NativeRef, "setCommentStatus", "resolved", path + ".resolved");
            }

            var requestedStatus = target.Modern.Status;
            if (statusChanged)
            {
                if (after.Status is not ("active" or "resolved" or "closed"))
                    throw Unsupported(path + ".status", "modern comment status must be active, resolved, or closed");
                if (after.Resolved != !after.Status.Equals("active", StringComparison.Ordinal))
                    throw Unsupported(path, "modern comment resolved must agree with status");
                requestedStatus = after.Status;
            }
            else if (resolvedChanged)
            {
                requestedStatus = after.Resolved
                    ? (target.Modern.Status == "closed" ? "closed" : "resolved")
                    : "active";
            }
            target.Modern.Text = after.Text;
            target.Modern.Status = requestedStatus;
            changedNodeIds.Add(after.Id);
            changedNodeIds.Add(after.PageId);
            changed = true;
        }
        return changed;
    }

    private sealed record SourceCommentTarget(
        PresentationSlide Slide,
        PresentationLegacyComment? Legacy,
        PresentationModernCommentThread? ModernThread,
        PresentationModernComment? Modern);

    private static bool ApplySections(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        ISet<string> changedNodeIds)
    {
        if (baseline.Sections.Count != requested.Sections.Count ||
            presentation.Sections.Count != baseline.Sections.Count)
            throw Unsupported("$.sections", "adding or removing source-bound sections");
        if (baseline.Sections.Count == 0) return false;
        if (baseline.Sections.Select(section => section.Raw).Zip(requested.Sections.Select(section => section.Raw)).All(pair => JsonEqual(pair.First, pair.Second)))
            return false;
        var slideIdByPageId = RouteSlideIds(baseline, requested, presentation, "$.sections");

        var changed = false;
        for (var index = 0; index < baseline.Sections.Count; index++)
        {
            var before = baseline.Sections[index];
            var after = requested.Sections[index];
            var path = $"$.sections[{index}]";
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal))
                throw Unsupported(path, "section reorder or identity change");
            RequireNativeRef(before.Raw, after.Raw, path);
            RequireEqualExcept(before.Raw, after.Raw, path, "name", "pages");
            var nameChanged = PropertyChanged(before.Raw, after.Raw, "name");
            var pagesChanged = PropertyChanged(before.Raw, after.Raw, "pages");
            if (nameChanged) RequireCapability(after.NativeRef, "setName", path + ".name");
            if (pagesChanged) RequireCapability(after.NativeRef, "setPages", path + ".pages");
            if (!nameChanged && !pagesChanged) continue;

            var target = presentation.Sections[index];
            target.Name = after.Name;
            target.SlideIds.Clear();
            foreach (var pageId in after.PageIds)
                target.SlideIds.Add(slideIdByPageId[pageId]);
            changedNodeIds.Add(after.Id);
            changed = true;
        }
        return changed;
    }

    private static bool ApplyCustomShows(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        ISet<string> changedNodeIds)
    {
        if (baseline.CustomShows.Count != requested.CustomShows.Count ||
            presentation.CustomShows.Count != baseline.CustomShows.Count)
            throw Unsupported("$.customShows", "adding or removing source-bound custom shows");
        if (baseline.CustomShows.Count == 0) return false;
        if (baseline.CustomShows.Select(show => show.Raw).Zip(requested.CustomShows.Select(show => show.Raw)).All(pair => JsonEqual(pair.First, pair.Second)))
            return false;
        var slideIdByPageId = RouteSlideIds(baseline, requested, presentation, "$.customShows");

        var changed = false;
        for (var index = 0; index < baseline.CustomShows.Count; index++)
        {
            var before = baseline.CustomShows[index];
            var after = requested.CustomShows[index];
            var path = $"$.customShows[{index}]";
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal))
                throw Unsupported(path, "custom-show reorder or identity change");
            RequireNativeRef(before.Raw, after.Raw, path);
            RequireEqualExcept(before.Raw, after.Raw, path, "name", "pages");
            var nameChanged = PropertyChanged(before.Raw, after.Raw, "name");
            var pagesChanged = PropertyChanged(before.Raw, after.Raw, "pages");
            if (nameChanged) RequireCapability(after.NativeRef, "setName", path + ".name");
            if (pagesChanged) RequireCapability(after.NativeRef, "setPages", path + ".pages");
            if (!nameChanged && !pagesChanged) continue;

            var target = presentation.CustomShows[index];
            target.Name = after.Name;
            target.SlideIds.Clear();
            foreach (var pageId in after.PageIds)
                target.SlideIds.Add(slideIdByPageId[pageId]);
            changedNodeIds.Add(after.Id);
            changed = true;
        }
        return changed;
    }

    private static IReadOnlyDictionary<string, string> RouteSlideIds(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        string path)
    {
        if (baseline.Pages.Count != requested.Pages.Count ||
            presentation.Slides.Count != requested.Pages.Count ||
            !baseline.Pages.Select(page => page.Id).ToHashSet(StringComparer.Ordinal)
                .SetEquals(requested.Pages.Select(page => page.Id)))
            throw Unsupported(path, "route editing with changed page topology");
        return requested.Pages.Select((page, index) => (PageId: page.Id, SlideId: presentation.Slides[index].Id))
            .ToDictionary(item => item.PageId, item => item.SlideId, StringComparer.Ordinal);
    }

    private static bool ApplyElement(
        PpjProgramModel program,
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
        var semanticChangesBefore = mutations.SemanticChanges;
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
        var frameLeafFastPath = TryCollectFrameLeafMutations(
            before,
            after,
            target,
            slide,
            shapeTreePath,
            nativeLeafBindings,
            mutations,
            path);
        var groupChildFrameLeafFastPath = TryCollectGroupChildFrameLeafMutations(
            before,
            after,
            target,
            slide,
            shapeTreePath,
            nativeLeafBindings,
            mutations,
            path);
        var stateChanged = ApplyElementState(before, after, target, path);
        if (stateChanged) mutations.SemanticChanges = true;

        // Accessibility is an owner-local, non-visual leaf.  Apply it before
        // the typed content switch so every recognized source-bound element
        // shares one capability gate while the existing PPTX writers remain
        // responsible for the exact cNvPr/adec serialization and residual
        // checks.  The PPJ model keeps explicit decorative=false presence, so
        // this must not be treated as a truthy-only convenience flag.
        var accessibilityChanged = PropertyChanged(before.Raw, after.Raw, "accessibility");
        if (accessibilityChanged)
        {
            RequireCapability(after, "setAccessibility", path + ".accessibility");
            ApplySourceBoundAccessibility(after.Accessibility, target, path + ".accessibility");
            mutations.SemanticChanges = true;
        }

        // Action is a shape-tree-owned relationship leaf. Apply it before
        // the typed content switch so text, line, icon, and placeholder
        // projections share one source-bound closure. PptxCodec performs the
        // relationship-safe patch when the artifact is exported.
        var actionChanged = PropertyChanged(before.Raw, after.Raw, "action");
        if (actionChanged)
        {
            if (target.ContentCase != PresentationElement.ContentOneofCase.Shape)
                throw Unsupported(path + ".action", "source-bound action edits require a recognized shape-producing element");
            RequireCapability(after, "setAction", path + ".action");
            target.Shape.Action = after.Raw.TryGetProperty("action", out var action)
                ? PpjAuthoredPresentationCompiler.BuildAction(action, after.Id)
                : null;
            mutations.SemanticChanges = true;
        }
        var hoverActionChanged = PropertyChanged(before.Raw, after.Raw, "hoverAction");
        if (hoverActionChanged)
        {
            if (target.ContentCase != PresentationElement.ContentOneofCase.Shape)
                throw Unsupported(path + ".hoverAction", "source-bound hoverAction edits require a recognized shape-producing element");
            RequireCapability(after, "setHoverAction", path + ".hoverAction");
            target.Shape.HoverAction = after.Raw.TryGetProperty("hoverAction", out var hoverAction)
                ? PpjAuthoredPresentationCompiler.BuildAction(hoverAction, after.Id)
                : null;
            mutations.SemanticChanges = true;
        }

        bool changed;
        switch (before)
        {
            case PpjTextElementModel beforeText when after is PpjTextElementModel afterText && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyTextElement(program, beforeText, afterText, target, slide, shapeTreePath, assets, assetDimensions, mutations, path);
                break;
            case PpjShapeElementModel beforeShape when after is PpjShapeElementModel afterShape && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyShapeElement(program, beforeShape, afterShape, target, slide, shapeTreePath, assets, assetDimensions, mutations, path);
                break;
            case PpjPlaceholderElementModel beforePlaceholder when after is PpjPlaceholderElementModel afterPlaceholder && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                changed = ApplyPlaceholderElement(program, beforePlaceholder, afterPlaceholder, target, slide, shapeTreePath, mutations, path);
                break;
            case PpjIconElementModel when after is PpjIconElementModel && target.ContentCase == PresentationElement.ContentOneofCase.Shape:
                RequireEqualExcept(before.Raw, after.Raw, path, "action", "hoverAction", "accessibility");
                changed = actionChanged || hoverActionChanged;
                break;
            case PpjImageElementModel beforeImage when after is PpjImageElementModel afterImage && target.ContentCase == PresentationElement.ContentOneofCase.Image:
                changed = ApplyImageElement(program, beforeImage, afterImage, target.Image, assets, assetDimensions, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjChartElementModel beforeChart when after is PpjChartElementModel afterChart && target.ContentCase == PresentationElement.ContentOneofCase.Chart:
                changed = ApplyChartElement(program, beforeChart, afterChart, target.Chart, assets, assetDimensions, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjTableElementModel beforeTable when after is PpjTableElementModel afterTable && target.ContentCase == PresentationElement.ContentOneofCase.Table:
                changed = ApplyTableElement(program, beforeTable, afterTable, target.Table, assets, assetDimensions, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjConnectorElementModel beforeConnector when after is PpjConnectorElementModel afterConnector && target.ContentCase == PresentationElement.ContentOneofCase.Connector:
                changed = ApplyConnectorElement(program, beforeConnector, afterConnector, target.Connector, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjGroupElementModel beforeGroup when after is PpjGroupElementModel afterGroup && target.ContentCase == PresentationElement.ContentOneofCase.Group:
                changed = ApplyGroupElement(program, beforeGroup, afterGroup, target.Group, slide, shapeTreePath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, path);
                break;
            case PpjSmartArtElementModel beforeSmartArt when after is PpjSmartArtElementModel afterSmartArt &&
                target.ContentCase == PresentationElement.ContentOneofCase.Opaque && target.Opaque.DiagramText is not null:
                changed = ApplySourceSmartArtElement(beforeSmartArt, afterSmartArt, target.Opaque, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjSmartArtElementModel beforeSmartArt when after is PpjSmartArtElementModel afterSmartArt &&
                target.ContentCase == PresentationElement.ContentOneofCase.Diagram:
                changed = ApplyNativeSmartArtElement(beforeSmartArt, afterSmartArt, target.Diagram, assets, program.Root, path);
                if (afterSmartArt.DetachToShapes && !beforeSmartArt.DetachToShapes)
                {
                    var detached = target.Diagram.Drawing.Clone();
                    target.Group = detached;
                    mutations.SemanticChanges = true;
                    mutations.Warnings.Add(new Diagnostic
                    {
                        Severity = DiagnosticSeverity.Warning,
                        Code = "ppj.smartArt.detachedToShapes",
                        Message = "SmartArt was explicitly detached to ordinary shapes; diagram semantics were removed.",
                        SourcePath = path + ".detachToShapes",
                    });
                }
                break;
            case PpjOleElementModel beforeOle when after is PpjOleElementModel afterOle &&
                target.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                (target.Opaque.OleWorkbook is not null || target.Opaque.OleOfficePackage is not null):
                changed = ApplySourceOleElement(beforeOle, afterOle, target.Opaque, assets, path);
                if (changed) mutations.SemanticChanges = true;
                break;
            case PpjOpaqueElementModel beforeOpaque when after is PpjOpaqueElementModel afterOpaque:
                changed = ApplyOpaqueElement(beforeOpaque, afterOpaque, target, slide, shapeTreePath, mutations, path);
                break;
            default:
                throw Unsupported(path, "the exact source object no longer matches its PPJ projection type");
        }
        // ApplyFrame updates the in-memory presentation for the conservative
        // writer path and marks a semantic mutation. If the frame is the only
        // PPJ field that changed and every changed coordinate has a fresh,
        // hash-bound native leaf, the token-splice edit plan can carry the same
        // operation without materializing/reserializing the full source IR.
        // Restore the prior semantic flag so the caller selects that plan.
        if (frameLeafFastPath || groupChildFrameLeafFastPath)
            mutations.SemanticChanges = semanticChangesBefore;
        changed |= nativeLeafChanged || stateChanged;
        changed |= nativeLeafChanged || stateChanged || actionChanged || hoverActionChanged || accessibilityChanged;
        if (changed) changedNodeIds.Add(after.Id);
        return changed;
    }

    private static void ApplySourceBoundAccessibility(
        PpjAccessibilityModel? source,
        PresentationElement target,
        string path)
    {
        switch (target.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                target.Shape.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Image:
                target.Image.ClearAccessibilityDecorative();
                target.Image.AccessibilityTitle = string.Empty;
                target.Image.AltText = string.Empty;
                if (source is null) return;
                // Unlike the other wire owners, PresentationImage exposes
                // title/description through compatibility string fields.  Set
                // the optional decorative bit even when false so the PPJ
                // object's explicit classification survives the round trip.
                target.Image.AccessibilityDecorative = source.Decorative;
                if (!source.Decorative)
                {
                    if (source.Title is not null) target.Image.AccessibilityTitle = source.Title;
                    if (source.Description is not null) target.Image.AltText = source.Description;
                }
                return;
            case PresentationElement.ContentOneofCase.Chart:
                target.Chart.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Table:
                target.Table.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Connector:
                target.Connector.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Group:
                target.Group.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Diagram:
                target.Diagram.Accessibility = SourceBoundAccessibility(source);
                return;
            case PresentationElement.ContentOneofCase.Opaque when target.Opaque.NativeKind == "media":
                target.Opaque.Accessibility = SourceBoundAccessibility(source);
                return;
            default:
                throw Unsupported(path, "source-bound accessibility requires a recognized presentation owner");
        }
    }

    private static PresentationNonVisualAccessibility? SourceBoundAccessibility(PpjAccessibilityModel? source)
    {
        if (source is null) return null;
        var output = new PresentationNonVisualAccessibility
        {
            // `decorative` is required by PPJ even when false.  Preserve its
            // presence in the optional wire field rather than collapsing
            // explicit false into omission.
            Decorative = source.Decorative,
        };
        if (source.Title is not null) output.Title = source.Title;
        if (source.Description is not null) output.Description = source.Description;
        return output;
    }

    private static bool ApplyElementState(
        PpjElementModel before,
        PpjElementModel after,
        PresentationElement target,
        string path)
    {
        var changed = false;
        if ((before.Hidden ?? false) != (after.Hidden ?? false))
        {
            if (after.Hidden is null)
                throw Unsupported(path + ".hidden", "removing an issued hidden state; set an explicit boolean instead");
            RequireCapability(after, "setHidden", path + ".hidden");
            target.Hidden = after.Hidden.Value;
            changed = true;
        }
        if ((before.Locked ?? false) != (after.Locked ?? false))
        {
            if (after.Locked is null)
                throw Unsupported(path + ".locked", "removing an issued locked state; set an explicit boolean instead");
            RequireCapability(after, "setLocked", path + ".locked");
            target.Locked = after.Locked.Value;
            changed = true;
        }
        return changed;
    }

    private static bool ApplyTextElement(
        PpjProgramModel program,
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
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "text", "style", "fill", "stroke", "action", "hoverAction", "accessibility");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "style"))
        {
            RequireCapabilityField(after.NativeRef, "setTextBodyStyle", "text.style", path + ".style");
            if (!after.Raw.TryGetProperty("style", out var style))
                throw Unsupported(path + ".style", "removing source-bound text body style is not an explicit bounded operation");
            PpjAuthoredPresentationCompiler.MergeSourceBoundTextBodyStyle(target.TextBody!, style, path + ".style");
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, after.NativeRef, slide, element, shapeTreePath, mutations, path + ".text", program.Root);
        }
        semanticChanged |= ApplyFillProperty(before, after, target, "fill", assets, assetDimensions, program.Root, path);
        semanticChanged |= ApplyStrokeProperty(before, after, target, "stroke", program.Root, path);
        mutations.SemanticChanges |= semanticChanged;
        changed |= semanticChanged;
        return changed;
    }

    private static bool ApplyShapeElement(
        PpjProgramModel program,
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
        if (before.Type == "line")
            return ApplyLineElement(program, before, after, element, mutations, path);
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "geometry", "text", "style", "textStyle", "compositing", "action", "hoverAction", "accessibility");
        var styleChanged = PropertyChanged(before.Raw, after.Raw, "style");
        var compositingChanged = PropertyChanged(before.Raw, after.Raw, "compositing");
        if (styleChanged && compositingChanged)
            throw Unsupported(path + ".compositing", "source-bound shape style and compound opacity cannot be changed together");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "textStyle"))
        {
            RequireCapabilityField(after.NativeRef, "setTextBodyStyle", "textStyle", path + ".textStyle");
            if (!after.Raw.TryGetProperty("textStyle", out var textStyle))
                throw Unsupported(path + ".textStyle", "removing source-bound text body style is not an explicit bounded operation");
            PpjAuthoredPresentationCompiler.MergeSourceBoundTextBodyStyle(target.TextBody!, textStyle, path + ".textStyle");
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source text body");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, after.NativeRef, slide, element, shapeTreePath, mutations, path + ".text", program.Root);
        }
        if (PropertyChanged(before.Raw, after.Raw, "geometry"))
        {
            var oldGeometry = before.Raw.GetProperty("geometry");
            var newGeometry = after.Raw.GetProperty("geometry");
            if (before.GeometryKind == "custom" && after.GeometryKind == "custom")
            {
                RequireCapability(after, "setGeometry", path + ".geometry.paths");
                RequireEqualExcept(oldGeometry, newGeometry, path + ".geometry", "paths");
                if (!IsLiteralCustomGeometry(target))
                    throw Unsupported(path + ".geometry", "source custom geometry is outside the literal path edit profile");
                target.CustomPaths.Clear();
                target.CustomAdjustments.Clear();
                target.CustomGuides.Clear();
                target.CustomConnectionSites.Clear();
                target.CustomAdjustmentHandles.Clear();
                target.TextRectangle = null;
                PpjAuthoredPresentationCompiler.ApplyCustomGeometry(target, newGeometry, after.Id);
            }
            else
            {
                RequireCapability(after, "setGeometry", path + ".geometry.adjustments");
                RequireEqualExcept(oldGeometry, newGeometry, path + ".geometry", "adjustments");
                target.PresetAdjustments.Clear();
                target.PresetAdjustments.Add(after.GeometryAdjustments);
            }
            semanticChanged = true;
        }
        semanticChanged |= ApplyShapeStyle(before, after, target, assets, assetDimensions, program.Root, path);
        semanticChanged |= ApplySourceBoundShapeCompositing(before, after, target, program.Root, path);
        mutations.SemanticChanges |= semanticChanged;
        changed |= semanticChanged;
        return changed;
    }

    private static bool ApplySourceBoundShapeCompositing(
        PpjShapeElementModel before,
        PpjShapeElementModel after,
        PresentationShape target,
        JsonElement grammarRoot,
        string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, "compositing")) return false;
        var beforeCompositing = OptionalProperty(before.Raw, "compositing");
        var afterCompositing = OptionalProperty(after.Raw, "compositing");
        if (beforeCompositing is { } oldValue && afterCompositing is { } newValue)
            RequireEqualExcept(oldValue, newValue, path + ".compositing", "opacity");
        else
        {
            var present = beforeCompositing ?? afterCompositing;
            if (present is { } presentValue)
                foreach (var property in presentValue.EnumerateObject())
                    if (property.Name != "opacity")
                        throw Unsupported(path + ".compositing", $"changing {property.Name}");
        }
        if (afterCompositing is { } requested && !requested.TryGetProperty("opacity", out _))
            throw Unsupported(path + ".compositing.opacity", "source-bound shape opacity must remain explicit");
        if (!PpjPresentationProjector.TryGetCompoundShapeOpacity(target, out _))
            throw Unsupported(path + ".compositing.opacity", "source-bound shape is outside the compound opacity owner profile");
        RequireCapabilityField(after.NativeRef, "setOpacity", "compositing.opacity", path + ".compositing.opacity");
        var opacity = afterCompositing is { } compositing && compositing.TryGetProperty("opacity", out var value)
            ? ResolveGrammarOpacityToken(grammarRoot, value, path + ".compositing.opacity")
            : 1d;
        PpjAuthoredPresentationCompiler.SetCompoundShapeOpacity(target, opacity, after.Id);
        return true;
    }

    private static bool IsLiteralCustomGeometry(PresentationShape shape)
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
            PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => IsLiteral(command.MoveTo),
            PresentationCustomGeometryCommand.CommandOneofCase.LineTo => IsLiteral(command.LineTo),
            PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo =>
                IsLiteral(command.QuadraticBezierTo.Control) && IsLiteral(command.QuadraticBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo =>
                IsLiteral(command.CubicBezierTo.Control1) && IsLiteral(command.CubicBezierTo.Control2) && IsLiteral(command.CubicBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => IsLiteral(command.ArcTo),
            PresentationCustomGeometryCommand.CommandOneofCase.Close => true,
            _ => false,
        });
    }

    private static bool IsLiteral(PresentationCustomGeometryPoint point) =>
        !point.HasXReference && !point.HasYReference;

    private static bool IsLiteral(PresentationCustomGeometryArc arc) =>
        !arc.HasWidthRadiusReference && !arc.HasHeightRadiusReference &&
        !arc.HasStartAngleReference && !arc.HasSweepAngleReference;

    private static bool ApplyLineElement(
        PpjProgramModel program,
        PpjShapeElementModel before,
        PpjShapeElementModel after,
        PresentationElement element,
        MutationState mutations,
        string path)
    {
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "path", "points", "viewBox", "curve", "stroke", "shadow", "glow", "innerShadow", "reflection", "softEdge", "startArrow", "endArrow", "action", "hoverAction", "accessibility");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        var pathChanged = PropertyChanged(before.Raw, after.Raw, "path") ||
            PropertyChanged(before.Raw, after.Raw, "points") ||
            PropertyChanged(before.Raw, after.Raw, "viewBox") ||
            PropertyChanged(before.Raw, after.Raw, "curve");
        var curveChanged = PropertyChanged(before.Raw, after.Raw, "curve");
        if (curveChanged && !after.Raw.TryGetProperty("points", out _))
            throw Unsupported(path + ".curve", "source-bound Kimi curve edits require the compact points form");
        if (pathChanged)
        {
            RequireCapability(after, "setLinePath", path + ".path");
            if (after.Raw.TryGetProperty("path", out var pathValue) && !curveChanged)
                PpjLinePathCodec.Apply(target, pathValue, after.Id);
            else
                PpjLinePathCodec.Apply(target, PpjLinePathCodec.KimiPath(after.Raw, after.Frame.Width, after.Frame.Height, after.Id), after.Id);
            semanticChanged = true;
        }
        semanticChanged |= ApplyStrokeProperty(before, after, target, "stroke", program.Root, path);
        if (curveChanged && !HasExplicitStrokeJoin(after.Raw))
        {
            var curve = OptionalString(after.Raw, "curve") ?? "round";
            if (curve is "round" or "sharp") target.LineJoin = curve == "round" ? "round" : "miter";
        }
        if (PropertyChanged(before.Raw, after.Raw, "startArrow") || PropertyChanged(before.Raw, after.Raw, "endArrow"))
        {
            RequireCapability(after, "setStroke", path + ".stroke");
            target.StartArrow = ArrowValue(after.Raw, "startArrow");
            target.EndArrow = ArrowValue(after.Raw, "endArrow");
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "shadow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.shadow", path + ".shadow");
            target.Shadow = after.Raw.TryGetProperty("shadow", out var shadow)
                ? SourceBoundShadow(shadow, path + ".shadow", "shape", program.Root)
                : null;
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "glow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.glow", path + ".glow");
            target.Glow = after.Raw.TryGetProperty("glow", out var glow)
                ? SourceBoundGlow(glow, path + ".glow", "shape", program.Root)
                : null;
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "innerShadow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.innerShadow", path + ".innerShadow");
            target.InnerShadow = after.Raw.TryGetProperty("innerShadow", out var innerShadow)
                ? SourceBoundInnerShadow(innerShadow, path + ".innerShadow", "shape", program.Root)
                : null;
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "reflection"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.reflection", path + ".reflection");
            target.Reflection = after.Raw.TryGetProperty("reflection", out var reflection)
                ? SourceBoundReflection(reflection, path + ".reflection", "shape", program.Root)
                : null;
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "softEdge"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.softEdge", path + ".softEdge");
            target.SoftEdge = after.Raw.TryGetProperty("softEdge", out var softEdge)
                ? SourceBoundSoftEdge(softEdge, path + ".softEdge", "shape", program.Root)
                : null;
            semanticChanged = true;
        }
        mutations.SemanticChanges |= semanticChanged;
        return changed | semanticChanged;
    }

    private static bool HasExplicitStrokeJoin(JsonElement raw) =>
        raw.TryGetProperty("stroke", out var stroke) && stroke.ValueKind == JsonValueKind.Object &&
        stroke.TryGetProperty("join", out _);

    private static bool ApplyPlaceholderElement(
        PpjProgramModel program,
        PpjPlaceholderElementModel before,
        PpjPlaceholderElementModel after,
        PresentationElement element,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        MutationState mutations,
        string path)
    {
        var target = element.Shape;
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "text", "style", "action", "hoverAction", "accessibility");
        var semanticChanged = ApplyFrame(before, after, target, path);
        var changed = semanticChanged;
        if (PropertyChanged(before.Raw, after.Raw, "style"))
        {
            RequireCapabilityField(after.NativeRef, "setTextBodyStyle", "text.style", path + ".style");
            if (!after.Raw.TryGetProperty("style", out var style))
                throw Unsupported(path + ".style", "removing source-bound placeholder text body style is not an explicit bounded operation");
            PpjAuthoredPresentationCompiler.MergeSourceBoundTextBodyStyle(target.TextBody!, style, path + ".style");
            semanticChanged = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source placeholder text body");
            changed |= CollectTextLeafMutations(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"),
                target, after.Id, after.NativeRef, slide, element, shapeTreePath, mutations, path + ".text", program.Root);
        }
        mutations.SemanticChanges |= semanticChanged;
        return changed;
    }

    private static bool ApplyImageElement(
        PpjProgramModel program,
        PpjImageElementModel before,
        PpjImageElementModel after,
        PresentationImage target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "asset", "svgAsset", "styleRef", "style", "fit", "crop", "focus", "opacity", "mask", "border", "shadow", "glow", "innerShadow", "reflection", "softEdge", "accessibility");
        var changed = ApplyFrame(before, after, target, path);
        var catalog = new PpjAuthoredPresentationCompiler.Catalog(program.Root);
        if (!before.AssetId.Equals(after.AssetId, StringComparison.Ordinal))
        {
            RequireCapability(after, "replaceImage", path + ".asset");
            if (!assets.TryGetValue(after.AssetId, out var nativeAssetId))
                throw new CodecException("ppj.asset.missing", $"PPJ image asset {after.AssetId} has no validated bytes.", path + ".asset");
            target.AssetId = nativeAssetId;
            changed = true;
        }
        if (!string.Equals(before.SvgAssetId, after.SvgAssetId, StringComparison.Ordinal))
        {
            RequireCapability(after, "replaceSvg", path + ".svgAsset");
            if (before.SvgAssetId is null || after.SvgAssetId is null)
                throw Unsupported(path + ".svgAsset", "adding or removing a native raster/SVG fallback pair");
            if (!assets.TryGetValue(after.SvgAssetId, out var nativeSvgAssetId))
                throw new CodecException("ppj.asset.missing", $"PPJ SVG asset {after.SvgAssetId} has no validated bytes.", path + ".svgAsset");
            target.SvgAssetId = nativeSvgAssetId;
            changed = true;
        }
        var beforeInlineStyle = OptionalProperty(before.Raw, "style");
        var afterInlineStyle = OptionalProperty(after.Raw, "style");
        var beforeNamedStyle = catalog.ImageStyle(before.StyleRef);
        var afterNamedStyle = catalog.ImageStyle(after.StyleRef);
        var beforeFit = EffectiveImageFit(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, path + ".fit");
        var afterFit = EffectiveImageFit(after.Raw, afterInlineStyle, afterNamedStyle, catalog, path + ".fit");
        var beforeCrop = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "crop");
        var afterCrop = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "crop");
        var beforeFocus = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "focus");
        var afterFocus = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "focus");
        var cropChanged = !JsonEqual(beforeCrop, afterCrop) || !JsonEqual(beforeFocus, afterFocus);
        var fitChanged = !string.Equals(beforeFit, afterFit, StringComparison.Ordinal);
        if (cropChanged) RequireCapability(after, "setImageCrop", path + ".crop");
        if (fitChanged) RequireCapability(after, "setImageFit", path + ".fit");
        if (cropChanged || fitChanged)
        {
            var paint = BuildImagePaint(
                EffectiveImagePaintSource(after, after.Raw, afterInlineStyle, afterNamedStyle, catalog),
                after.Frame.Width,
                after.Frame.Height,
                assets,
                assetDimensions,
                path,
                resolveOpacity: opacity => ResolveGrammarOpacityToken(program.Root, opacity, path + ".opacity"),
                resolveFit: fit => ResolveGrammarStringToken(program.Root, fit, path + ".fit"));
            target.Crop = paint.Crop;
            target.Tiled = paint.Mode == PresentationImagePaint.Types.Mode.Tile;
            changed = true;
        }
        var beforeOpacity = EffectiveImageOpacity(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, path + ".opacity");
        var afterOpacity = EffectiveImageOpacity(after.Raw, afterInlineStyle, afterNamedStyle, catalog, path + ".opacity");
        if (beforeOpacity != afterOpacity)
        {
            RequireCapability(after, "setOpacity", path + ".opacity");
            if (afterOpacity is { } opacity)
                target.OpacityThousandthPercent = Unit(opacity);
            else target.ClearOpacityThousandthPercent();
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "mask"))
        {
            var hasBeforeMask = before.Raw.TryGetProperty("mask", out var beforeMask);
            var hasAfterMask = after.Raw.TryGetProperty("mask", out var afterMask);
            // PPJ omits the native rectangle mask by default.  Normalize that
            // omission to the same canonical preset identity so a source-bound
            // edit can change rect <-> another supported preset without
            // confusing absence with a topology change.
            var beforePreset = before.MaskKind == "preset"
                ? before.MaskPreset ?? "rect"
                : !hasBeforeMask ? "rect" : null;
            var afterPreset = after.MaskKind == "preset"
                ? after.MaskPreset ?? "rect"
                : !hasAfterMask ? "rect" : null;
            if (beforePreset is not null && afterPreset is not null &&
                !beforePreset.Equals(afterPreset, StringComparison.Ordinal))
            {
                RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.preset", path + ".mask.preset");
                if (!PptxCustomGeometryCodec.TryPreset(afterPreset, out _))
                    throw Unsupported(path + ".mask.preset", "source-bound picture mask preset is outside the supported DrawingML catalog");
                if (before.MaskAdjustments.Count > 0 || after.MaskAdjustments.Count > 0)
                    RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.adjustments", path + ".mask.adjustments");
                target.MaskPreset = afterPreset.Equals("rect", StringComparison.Ordinal) ? string.Empty : afterPreset;
                target.MaskPresetAdjustments.Clear();
                target.MaskPresetAdjustments.Add(after.MaskAdjustments);
                changed = true;
            }
            else if (beforePreset is not null && after.MaskKind == "custom" &&
                     before.MaskAdjustments.Count == 0)
            {
                // A literal custom mask is an owner-local geometry.  When the
                // source preset is the native rectangle/default (or another
                // no-adjustment preset), replacing that geometry is safe as a
                // single picture-owned topology edit; no relationship or
                // descendant shape identity is involved.
                RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.preset", path + ".mask.preset");
                RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.paths", path + ".mask.paths");
                if (!hasAfterMask)
                    throw Unsupported(path + ".mask", "source-bound custom picture mask is missing its mask object");
                var maskShape = new PresentationShape { Geometry = "custom" };
                PpjAuthoredPresentationCompiler.ApplyCustomGeometry(maskShape, afterMask, after.Id + " image mask");
                target.CustomMaskPaths.Clear();
                target.CustomMaskPaths.Add(maskShape.CustomPaths);
                target.MaskPreset = string.Empty;
                target.MaskPresetAdjustments.Clear();
            }
            else if (before.MaskKind == "custom" && afterPreset is not null)
            {
                // The inverse transition is equally local: replace a
                // recognized literal custom geometry with a supported preset
                // and keep the picture's existing relationship untouched.
                RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.preset", path + ".mask.preset");
                RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.paths", path + ".mask.paths");
                if (!PptxCustomGeometryCodec.TryPreset(afterPreset, out _))
                    throw Unsupported(path + ".mask.preset", "source-bound picture mask preset is outside the supported DrawingML catalog");
                if (after.MaskAdjustments.Count > 0)
                    RequireCapabilityField(after.NativeRef, "setImageMask", "image.mask.adjustments", path + ".mask.adjustments");
                target.CustomMaskPaths.Clear();
                target.MaskPreset = afterPreset.Equals("rect", StringComparison.Ordinal) ? string.Empty : afterPreset;
                target.MaskPresetAdjustments.Clear();
                target.MaskPresetAdjustments.Add(after.MaskAdjustments);
            }
            else if (before.MaskKind == "custom" && after.MaskKind == "custom")
            {
                RequireCapability(after, "setImageMask", path + ".mask.paths");
                if (!hasBeforeMask || !hasAfterMask)
                    throw Unsupported(path + ".mask", "source-bound custom picture mask is missing its mask object");
                RequireEqualExcept(beforeMask, afterMask, path + ".mask", "paths");
                var maskShape = new PresentationShape { Geometry = "custom" };
                PpjAuthoredPresentationCompiler.ApplyCustomGeometry(maskShape, afterMask, after.Id + " image mask");
                target.CustomMaskPaths.Clear();
                target.CustomMaskPaths.Add(maskShape.CustomPaths);
                target.MaskPreset = string.Empty;
                target.MaskPresetAdjustments.Clear();
            }
            else
            {
                RequireCapability(after, "setImageMask", path + ".mask.adjustments");
                if (beforePreset is null || afterPreset is null ||
                    !beforePreset.Equals(afterPreset, StringComparison.Ordinal) ||
                    !hasBeforeMask || !hasAfterMask)
                    throw Unsupported(path + ".mask", "source-bound picture mask topology or preset identity change");
                RequireEqualExcept(beforeMask, afterMask, path + ".mask", "adjustments");
                target.MaskPresetAdjustments.Clear();
                target.MaskPresetAdjustments.Add(after.MaskAdjustments);
            }
            changed = true;
        }
        var beforeBorder = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "border");
        var afterBorder = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "border");
        if (!JsonEqual(beforeBorder, afterBorder))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.border", path + ".border");
            target.Border = afterBorder is { } border
                ? SourceBoundImageBorder(border, path + ".border", program.Root)
                : null;
            changed = true;
        }
        var beforeShadow = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "shadow");
        var afterShadow = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "shadow");
        if (!JsonEqual(beforeShadow, afterShadow))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.shadow", path + ".shadow");
            target.Shadow = afterShadow is { } shadow
                ? SourceBoundShadow(shadow, path + ".shadow", "image", program.Root)
                : null;
            changed = true;
        }
        var beforeGlow = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "glow");
        var afterGlow = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "glow");
        if (!JsonEqual(beforeGlow, afterGlow))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.glow", path + ".glow");
            target.Glow = afterGlow is { } glow
                ? SourceBoundGlow(glow, path + ".glow", "image", program.Root)
                : null;
            changed = true;
        }
        var beforeInnerShadow = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "innerShadow");
        var afterInnerShadow = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "innerShadow");
        if (!JsonEqual(beforeInnerShadow, afterInnerShadow))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.innerShadow", path + ".innerShadow");
            target.InnerShadow = afterInnerShadow is { } innerShadow
                ? SourceBoundInnerShadow(innerShadow, path + ".innerShadow", "image", program.Root)
                : null;
            changed = true;
        }
        var beforeReflection = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "reflection");
        var afterReflection = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "reflection");
        if (!JsonEqual(beforeReflection, afterReflection))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.reflection", path + ".reflection");
            target.Reflection = afterReflection is { } reflection
                ? SourceBoundReflection(reflection, path + ".reflection", "image", program.Root)
                : null;
            changed = true;
        }
        var beforeSoftEdge = EffectiveImageProperty(before.Raw, beforeInlineStyle, beforeNamedStyle, catalog, "softEdge");
        var afterSoftEdge = EffectiveImageProperty(after.Raw, afterInlineStyle, afterNamedStyle, catalog, "softEdge");
        if (!JsonEqual(beforeSoftEdge, afterSoftEdge))
        {
            RequireCapabilityField(after.NativeRef, "setImageEffects", "image.softEdge", path + ".softEdge");
            target.SoftEdge = afterSoftEdge is { } softEdge
                ? SourceBoundSoftEdge(softEdge, path + ".softEdge", "image", program.Root)
                : null;
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartElement(
        PpjProgramModel program,
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path,
            "role", "tags", "hidden", "locked", "frame", "title", "data", "style",
            "titlePlacement", "displayBlanksAs", "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis", "spokeAxis", "accessibility");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "displayBlanksAs"))
        {
            RequireCapability(after, "setChartPlot", path + ".displayBlanksAs");
            if (after.Raw.TryGetProperty("displayBlanksAs", out var displayBlanksAs))
                target.DisplayBlanksAs = ResolveGrammarEnumToken(
                    program.Root,
                    displayBlanksAs,
                    path + ".displayBlanksAs",
                    "zero", "gap", "span");
            else
                target.ClearDisplayBlanksAs();
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "title"))
        {
            RequireCapability(after, "setChartTitle", path + ".title");
            if (!after.Raw.TryGetProperty("title", out var title))
            {
                target.Title = string.Empty;
                target.TitleBody = null;
            }
            else if (title.ValueKind == JsonValueKind.String)
            {
                target.Title = title.GetString()!;
                target.TitleBody = null;
            }
            else
            {
                target.TitleBody = PpjAuthoredPresentationCompiler.BuildChartTitleBody(program, after);
                target.Title = PptxTextCodec.Flatten(target.TitleBody);
            }
            if (!after.Raw.TryGetProperty("titlePlacement", out _)) target.ClearTitlePlacement();
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "titlePlacement"))
        {
            RequireCapability(after, "setChartTitle", path + ".titlePlacement");
            if (after.Raw.TryGetProperty("titlePlacement", out var titlePlacement))
                target.TitlePlacement = ResolveGrammarEnumToken(
                    program.Root,
                    titlePlacement,
                    path + ".titlePlacement",
                    "none", "aboveChart", "centeredOverlay");
            else
                target.ClearTitlePlacement();
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "data"))
        {
            ApplyChartData(before, after, target, program.Root, path + ".data");
            changed = true;
        }
        changed |= ApplyChartStyles(before, after, target, assets, assetDimensions, program.Root, path);
        if (PropertyChanged(OptionalProperty(before.Raw, "style"), OptionalProperty(after.Raw, "style"), "titleTextStyle") &&
            after.Raw.TryGetProperty("title", out var currentTitle) &&
            currentTitle.ValueKind != JsonValueKind.String)
        {
            target.TitleBody = PpjAuthoredPresentationCompiler.BuildChartTitleBody(program, after);
            target.Title = PptxTextCodec.Flatten(target.TitleBody);
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartStyles(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot,
        string path)
    {
        var changed = ApplyChartStyleTextStyles(before, after, target, assets, assetDimensions, grammarRoot, path);
        changed |= ApplyChartAxisStyle(before, after, target.XAxis, "xAxis", grammarRoot, path);
        changed |= ApplyChartAxisStyle(before, after, target.YAxis, "yAxis", grammarRoot, path);
        changed |= ApplyChartAxisStyle(before, after, target.SecondaryXAxis, "secondaryXAxis", grammarRoot, path);
        changed |= ApplyChartAxisStyle(before, after, target.SecondaryYAxis, "secondaryYAxis", grammarRoot, path);
        changed |= ApplyRadarSpokeAxis(before, after, target, grammarRoot, path);
        return changed;
    }

    private static bool ApplyChartStyleTextStyles(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot,
        string path)
    {
        var oldStyle = OptionalProperty(before.Raw, "style");
        var newStyle = OptionalProperty(after.Raw, "style");
        if (JsonEqual(oldStyle, newStyle)) return false;
        RequireOnlyBoundedProperties(
            oldStyle,
            newStyle,
            path + ".style",
            "legend",
            "legendOverlay",
            "legendFill",
            "legendLine",
            "stacking",
            "gapWidth",
            "overlap",
            "showCategoryAxis",
            "showValueAxis",
            "showGridlines",
            "titleTextStyle",
            "legendTextStyle",
            "dataLabels",
            "chartAreaFill",
            "plotAreaFill",
            "frame",
            "startAngle",
            "holeSize",
            "bubbleScale",
            "bubbleSizeMode",
            "smooth",
            "varyColors");

        var changed = false;
        if (PropertyChanged(oldStyle, newStyle, "legend"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.legend");
            var legend = newStyle is { } owner && owner.TryGetProperty("legend", out var value)
                ? ResolveGrammarEnumToken(grammarRoot, value, path + ".style.legend", "none", "top", "topRight", "bottom", "left", "right")
                : "none";
            target.HasLegend = !string.Equals(legend, "none", StringComparison.Ordinal);
            target.LegendPosition = target.HasLegend ? legend : string.Empty;
            if (!target.HasLegend)
            {
                target.LegendFill = null;
                target.LegendLine = null;
            }
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "legendOverlay"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.legendOverlay");
            if (newStyle is { } owner && owner.TryGetProperty("legendOverlay", out var value))
            {
                if (!target.HasLegend)
                    throw Unsupported(path + ".style.legendOverlay", "legendOverlay requires a visible legend");
                target.LegendOverlay = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.legendOverlay");
            }
            else
                target.ClearLegendOverlay();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "stacking"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.stacking");
            if (target.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area or SpreadsheetChartType.Combo))
                throw Unsupported(path + ".style.stacking", "stacking on a chart without a bounded categorical plot");
            var stacking = newStyle is { } owner && owner.TryGetProperty("stacking", out var value)
                ? ResolveGrammarEnumToken(grammarRoot, value, path + ".style.stacking", "none", "stacked", "percent-stacked")
                : "none";
            target.Grouping = stacking;
            if (target.Type == SpreadsheetChartType.Line && target.LineOptions?.HasGrouping == true)
                target.LineOptions.Grouping = stacking switch
                {
                    "none" => SpreadsheetChartLineGrouping.Standard,
                    "stacked" => SpreadsheetChartLineGrouping.Stacked,
                    "percent-stacked" => SpreadsheetChartLineGrouping.PercentStacked,
                    _ => throw new InvalidOperationException("Validated chart stacking changed unexpectedly."),
                };
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "gapWidth"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.gapWidth");
            var supportsGapWidth = target.Type == SpreadsheetChartType.Bar ||
                target.Type == SpreadsheetChartType.Combo && target.ComboSeries.Any(item => item.Type == SpreadsheetChartType.Bar);
            if (!supportsGapWidth)
                throw Unsupported(path + ".style.gapWidth", "gap width without a bounded column plot");
            if (newStyle is { } owner && owner.TryGetProperty("gapWidth", out var value))
                target.GapWidth = ResolveGrammarIntegerToken(grammarRoot, value, path + ".style.gapWidth", 0, 500);
            else
                target.ClearGapWidth();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "overlap"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.overlap");
            var supportsOverlap = target.Type == SpreadsheetChartType.Bar ||
                target.Type == SpreadsheetChartType.Combo && target.ComboSeries.Any(item => item.Type == SpreadsheetChartType.Bar);
            if (!supportsOverlap)
                throw Unsupported(path + ".style.overlap", "overlap without a bounded column plot");
            if (newStyle is { } owner && owner.TryGetProperty("overlap", out var value))
                target.Overlap = ResolveGrammarSignedIntegerToken(grammarRoot, value, path + ".style.overlap", -100, 100);
            else
                target.ClearOverlap();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "showCategoryAxis"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.showCategoryAxis");
            if (target.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut || target.XAxis is null)
                throw Unsupported(path + ".style.showCategoryAxis", "category-axis visibility without an existing categorical axis");
            if (newStyle is { } owner && owner.TryGetProperty("showCategoryAxis", out var value))
            {
                var visible = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.showCategoryAxis");
                target.ShowCategoryAxis = visible;
                target.XAxis.Visible = visible;
            }
            else
            {
                target.ClearShowCategoryAxis();
                target.XAxis.ClearVisible();
            }
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "showValueAxis"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.showValueAxis");
            if (target.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut || target.YAxis is null)
                throw Unsupported(path + ".style.showValueAxis", "value-axis visibility without an existing numeric axis");
            if (newStyle is { } owner && owner.TryGetProperty("showValueAxis", out var value))
            {
                var visible = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.showValueAxis");
                target.ShowValueAxis = visible;
                target.YAxis.Visible = visible;
            }
            else
            {
                target.ClearShowValueAxis();
                target.YAxis.ClearVisible();
            }
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "showGridlines"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.showGridlines");
            if (target.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut || target.YAxis is null)
                throw Unsupported(path + ".style.showGridlines", "gridline visibility without an existing numeric axis");
            if (newStyle is { } owner && owner.TryGetProperty("showGridlines", out var value))
            {
                var visible = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.showGridlines");
                if (visible)
                {
                    target.ShowGridlines = true;
                    target.YAxis.ShowMajorGridlines = true;
                }
                else
                {
                    // The native default for c:majorGridlines is omission.
                    // Clear both optional carriers so a false request has the
                    // same semantic hash as the post-write omitted node.
                    target.ClearShowGridlines();
                    target.YAxis.ClearShowMajorGridlines();
                }
            }
            else
            {
                target.ClearShowGridlines();
                target.YAxis.ClearShowMajorGridlines();
            }
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "titleTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + ".style.titleTextStyle");
            var titleTextStyle = SourceBoundChartTextStyle(
                newStyle,
                "titleTextStyle",
                path + ".style.titleTextStyle",
                grammarRoot);
            if (titleTextStyle is not null && target.Title.Length == 0)
                throw Unsupported(path + ".style.titleTextStyle", "chart-title style without an existing title");
            target.TitleTextStyle = titleTextStyle;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "legendTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + ".style.legendTextStyle");
            var legendTextStyle = SourceBoundChartTextStyle(
                newStyle,
                "legendTextStyle",
                path + ".style.legendTextStyle",
                grammarRoot);
            if (legendTextStyle is not null && !target.HasLegend)
                throw Unsupported(path + ".style.legendTextStyle", "chart-legend style without an existing legend");
            target.LegendTextStyle = legendTextStyle;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "legendFill"))
        {
            RequireCapabilityField(after, "setChartFill", "chart.legendFill", path + ".style.legendFill");
            var legendFill = SourceBoundChartFill(newStyle, "legendFill", path + ".style.legendFill", grammarRoot);
            if (legendFill is not null && !target.HasLegend)
                throw Unsupported(path + ".style.legendFill", "legendFill requires a visible legend");
            target.LegendFill = legendFill;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "legendLine"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.legendLine");
            var legendLine = newStyle is { } owner && owner.TryGetProperty("legendLine", out var value)
                ? SourceBoundChartLine(value, path + ".style.legendLine", grammarRoot)
                : null;
            if (legendLine is not null && !target.HasLegend)
                throw Unsupported(path + ".style.legendLine", "legendLine requires a visible legend");
            target.LegendLine = legendLine;
            changed = true;
        }

        var oldLabels = oldStyle is { } oldStyleValue ? OptionalProperty(oldStyleValue, "dataLabels") : null;
        var newLabels = newStyle is { } newStyleValue ? OptionalProperty(newStyleValue, "dataLabels") : null;
        if (!JsonEqual(oldLabels, newLabels))
        {
            if (oldLabels is null || newLabels is null || target.DataLabels is null)
                throw Unsupported(path + ".style.dataLabels", "source-bound chart-data-label topology change");
            RequireEqualExcept(oldLabels.Value, newLabels.Value, path + ".style.dataLabels",
                "showValue", "showCategory", "showSeries", "showPercent", "showBubbleSize", "showLeaderLines", "position", "textStyle", "numberFormat", "fill", "line");
            if (PropertyChanged(oldLabels, newLabels, "showValue"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showValue");
                target.DataLabels.ShowValue = newLabels.Value.TryGetProperty("showValue", out var showValue) &&
                    ResolveGrammarBooleanToken(grammarRoot, showValue, path + ".style.dataLabels.showValue");
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "showCategory"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showCategory");
                target.DataLabels.ShowCategoryName = newLabels.Value.TryGetProperty("showCategory", out var showCategory) &&
                    ResolveGrammarBooleanToken(grammarRoot, showCategory, path + ".style.dataLabels.showCategory");
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "showSeries"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showSeries");
                if (newLabels.Value.TryGetProperty("showSeries", out var showSeries))
                    target.DataLabels.ShowSeriesName = ResolveGrammarBooleanToken(
                        grammarRoot,
                        showSeries,
                        path + ".style.dataLabels.showSeries");
                else
                    target.DataLabels.ClearShowSeriesName();
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "showPercent"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showPercent");
                if (newLabels.Value.TryGetProperty("showPercent", out var showPercent))
                    target.DataLabels.ShowPercent = ResolveGrammarBooleanToken(
                        grammarRoot,
                        showPercent,
                        path + ".style.dataLabels.showPercent");
                else
                    target.DataLabels.ClearShowPercent();
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "showBubbleSize"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showBubbleSize");
                if (newLabels.Value.TryGetProperty("showBubbleSize", out var showBubbleSize))
                    target.DataLabels.ShowBubbleSize = ResolveGrammarBooleanToken(
                        grammarRoot,
                        showBubbleSize,
                        path + ".style.dataLabels.showBubbleSize");
                else
                    target.DataLabels.ClearShowBubbleSize();
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "showLeaderLines"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.showLeaderLines");
                if (newLabels.Value.TryGetProperty("showLeaderLines", out var showLeaderLines))
                    target.DataLabels.ShowLeaderLines = ResolveGrammarBooleanToken(
                        grammarRoot,
                        showLeaderLines,
                        path + ".style.dataLabels.showLeaderLines");
                else
                    target.DataLabels.ClearShowLeaderLines();
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "position"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.position");
                if (newLabels.Value.TryGetProperty("position", out var position))
                    target.DataLabels.Position = SourceBoundDataLabelPosition(
                        ResolveGrammarEnumToken(
                            grammarRoot,
                            position,
                            path + ".style.dataLabels.position",
                            "best-fit", "bottom", "center", "inside-base", "inside-end", "left", "outside-end", "right", "top"));
                else
                    target.DataLabels.ClearPosition();
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "textStyle"))
            {
                RequireCapability(after, "setChartTextStyle", path + ".style.dataLabels.textStyle");
                target.DataLabels.TextStyle = SourceBoundChartTextStyle(
                    newLabels,
                    "textStyle",
                    path + ".style.dataLabels.textStyle",
                    grammarRoot);
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "numberFormat"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.numberFormat");
                target.DataLabels.NumberFormatCode = newLabels.Value.TryGetProperty("numberFormat", out var format)
                    ? ResolveGrammarStringToken(grammarRoot, format, path + ".style.dataLabels.numberFormat")
                    : string.Empty;
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "fill"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.fill");
                target.DataLabels.Fill = SourceBoundChartFill(newLabels, "fill", path + ".style.dataLabels.fill", grammarRoot);
                changed = true;
            }
            if (PropertyChanged(oldLabels, newLabels, "line"))
            {
                RequireCapability(after, "setChartLabels", path + ".style.dataLabels.line");
                target.DataLabels.Line = newLabels.Value.TryGetProperty("line", out var line)
                    ? SourceBoundChartLine(line, path + ".style.dataLabels.line", grammarRoot)
                    : null;
                changed = true;
            }
        }
        if (PropertyChanged(oldStyle, newStyle, "chartAreaFill"))
        {
            RequireCapability(after, "setChartFill", path + ".style.chartAreaFill");
            target.ChartAreaFill = SourceBoundChartFill(newStyle, "chartAreaFill", path + ".style.chartAreaFill", grammarRoot);
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "plotAreaFill"))
        {
            RequireCapability(after, "setChartFill", path + ".style.plotAreaFill");
            target.PlotAreaFill = SourceBoundChartFill(newStyle, "plotAreaFill", path + ".style.plotAreaFill", grammarRoot);
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "frame"))
        {
            RequireCapability(after, "setChartFrame", path + ".style.frame");
            var oldFrame = oldStyle is { } oldOwner ? OptionalProperty(oldOwner, "frame") : null;
            var newFrame = newStyle is { } newOwner ? OptionalProperty(newOwner, "frame") : null;
            if (oldFrame is { } oldValue && newFrame is { } newValue)
                RequireEqualExcept(oldValue, newValue, path + ".style.frame", "fill", "stroke", "shadow");
            else if (oldFrame is { } presentOld)
                foreach (var property in presentOld.EnumerateObject())
                    if (property.Name is not ("fill" or "stroke" or "shadow"))
                        throw Unsupported(path + ".style.frame." + property.Name, "changing source-owned chart frame property");
            else if (newFrame is { } presentNew)
                foreach (var frameProperty in presentNew.EnumerateObject())
                    if (frameProperty.Name is not ("fill" or "stroke" or "shadow"))
                        throw Unsupported(path + ".style.frame." + frameProperty.Name, "changing source-owned chart frame property");
            EnsureSourceBoundChartFrameImageTopology(oldFrame, newFrame, path + ".style.frame.fill");
            target.Frame = newFrame is { } frame
                ? SourceBoundChartFrame(
                    frame,
                    path + ".style.frame",
                    assets,
                    assetDimensions,
                    grammarRoot,
                    target.WidthEmu / EmuPerPoint,
                    target.HeightEmu / EmuPerPoint)
                : null;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "startAngle"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.startAngle");
            if (target.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
                throw Unsupported(path + ".style.startAngle", "first-slice angle on a non-circular chart");
            if (newStyle is { } owner && owner.TryGetProperty("startAngle", out var value))
                target.FirstSliceAngle = ResolveGrammarIntegerToken(grammarRoot, value, path + ".style.startAngle", 0, 360);
            else
                target.ClearFirstSliceAngle();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "holeSize"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.holeSize");
            if (target.Type != SpreadsheetChartType.Doughnut)
                throw Unsupported(path + ".style.holeSize", "center-hole size on a non-doughnut chart");
            if (newStyle is { } owner && owner.TryGetProperty("holeSize", out var value))
                target.DoughnutHoleSize = ResolveGrammarIntegerToken(grammarRoot, value, path + ".style.holeSize", 10, 90);
            else
                target.ClearDoughnutHoleSize();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "bubbleScale"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.bubbleScale");
            if (target.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(path + ".style.bubbleScale", "bubble scale on a non-bubble chart");
            if (newStyle is { } owner && owner.TryGetProperty("bubbleScale", out var value))
                target.BubbleScale = ResolveGrammarIntegerToken(grammarRoot, value, path + ".style.bubbleScale", 0, 300);
            else
                target.ClearBubbleScale();
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "bubbleSizeMode"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.bubbleSizeMode");
            if (target.Type != SpreadsheetChartType.Bubble)
                throw Unsupported(path + ".style.bubbleSizeMode", "bubble size mode on a non-bubble chart");
            target.BubbleSizeMode = newStyle is { } owner && owner.TryGetProperty("bubbleSizeMode", out var value)
                ? ResolveGrammarEnumToken(grammarRoot, value, path + ".style.bubbleSizeMode", "area", "width")
                : string.Empty;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "smooth"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.smooth");
            if (target.Type != SpreadsheetChartType.Line)
                throw Unsupported(path + ".style.smooth", "smooth interpolation on a non-line chart");
            var options = target.LineOptions?.Clone() ?? new SpreadsheetChartLineOptionsArtifact();
            if (newStyle is { } owner && owner.TryGetProperty("smooth", out var value))
                options.Smooth = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.smooth");
            else
                options.ClearSmooth();
            target.LineOptions = options;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "varyColors"))
        {
            RequireCapability(after, "setChartPlot", path + ".style.varyColors");
            if (target.Type is not (SpreadsheetChartType.Line or SpreadsheetChartType.Bar or SpreadsheetChartType.Combo))
                throw Unsupported(path + ".style.varyColors", "color variation on an unsupported chart family");
            if (target.Type == SpreadsheetChartType.Line)
            {
                var options = target.LineOptions?.Clone() ?? new SpreadsheetChartLineOptionsArtifact();
                options.VaryColors = newStyle is { } owner && owner.TryGetProperty("varyColors", out var value) &&
                    ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.varyColors");
                target.LineOptions = options;
            }
            else
            {
                if (target.Type == SpreadsheetChartType.Combo && !target.ComboSeries.Any(item => item.Type == SpreadsheetChartType.Bar))
                    throw Unsupported(path + ".style.varyColors", "categorical combo color variation requires a column plot");
                if (newStyle is { } owner && owner.TryGetProperty("varyColors", out var value))
                    target.VaryColors = ResolveGrammarBooleanToken(grammarRoot, value, path + ".style.varyColors");
                else
                    target.ClearVaryColors();
            }
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartAxisStyle(
        PpjChartElementModel before,
        PpjChartElementModel after,
        SpreadsheetChartAxisArtifact? target,
        string axisName,
        JsonElement grammarRoot,
        string path)
    {
        var oldAxis = OptionalProperty(before.Raw, axisName);
        var newAxis = OptionalProperty(after.Raw, axisName);
        if (JsonEqual(oldAxis, newAxis)) return false;
        if (oldAxis is null || newAxis is null || target is null)
            throw Unsupported(path + "." + axisName, "source-bound chart-axis topology change");
        RequireEqualExcept(oldAxis.Value, newAxis.Value, path + "." + axisName,
            "textStyle", "title", "titleTextStyle", "visible", "numberFormat", "tickLabelInterval",
            "min", "max", "majorUnit", "minorUnit", "position", "majorTickMark", "minorTickMark", "tickLabelsVisible", "tickLabelPosition", "reverse", "axisLine", "axisLineArrow", "gridLine", "minorGridLine");
        var changed = false;
        if (PropertyChanged(oldAxis, newAxis, "title"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".title");
            target.Title = newAxis.Value.TryGetProperty("title", out var title)
                ? ResolveGrammarStringToken(grammarRoot, title, path + "." + axisName + ".title")
                : string.Empty;
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "textStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + "." + axisName + ".textStyle");
            target.TextStyle = SourceBoundChartTextStyle(
                newAxis,
                "textStyle",
                path + "." + axisName + ".textStyle",
                grammarRoot);
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "titleTextStyle"))
        {
            RequireCapability(after, "setChartTextStyle", path + "." + axisName + ".titleTextStyle");
            var titleTextStyle = SourceBoundChartTextStyle(
                newAxis,
                "titleTextStyle",
                path + "." + axisName + ".titleTextStyle",
                grammarRoot);
            if (titleTextStyle is not null && target.Title.Length == 0)
                throw Unsupported(path + "." + axisName + ".titleTextStyle", "axis-title style without an existing title");
            target.TitleTextStyle = titleTextStyle;
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "reverse"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".reverse");
            if (newAxis.Value.TryGetProperty("reverse", out var reverse))
                target.Reverse = ResolveGrammarBooleanToken(grammarRoot, reverse, path + "." + axisName + ".reverse");
            else target.ClearReverse();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "position"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".position");
            if (newAxis.Value.TryGetProperty("position", out var position))
                target.Position = ResolveGrammarEnumToken(
                    grammarRoot,
                    position,
                    path + "." + axisName + ".position",
                    "bottom", "left", "right", "top");
            else
                target.ClearPosition();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "visible"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".visible");
            if (newAxis.Value.TryGetProperty("visible", out var visible))
                target.Visible = ResolveGrammarBooleanToken(grammarRoot, visible, path + "." + axisName + ".visible");
            else target.ClearVisible();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "numberFormat"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".numberFormat");
            target.NumberFormatCode = newAxis.Value.TryGetProperty("numberFormat", out var format)
                ? ResolveGrammarStringToken(grammarRoot, format, path + "." + axisName + ".numberFormat")
                : string.Empty;
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "tickLabelInterval"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".tickLabelInterval");
            if (newAxis.Value.TryGetProperty("tickLabelInterval", out var interval))
                target.TickLabelInterval = ResolveGrammarIntegerToken(
                    grammarRoot,
                    interval,
                    path + "." + axisName + ".tickLabelInterval",
                    1,
                    10_000);
            else
                target.ClearTickLabelInterval();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "min"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".min");
            if (newAxis.Value.TryGetProperty("min", out var minimum))
                target.Minimum = ResolveGrammarNumberToken(grammarRoot, minimum, "size", path + "." + axisName + ".min");
            else target.ClearMinimum();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "max"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".max");
            if (newAxis.Value.TryGetProperty("max", out var maximum))
                target.Maximum = ResolveGrammarNumberToken(grammarRoot, maximum, "size", path + "." + axisName + ".max");
            else target.ClearMaximum();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "majorUnit"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".majorUnit");
            if (newAxis.Value.TryGetProperty("majorUnit", out var unit))
                target.MajorUnit = ResolveGrammarPositiveNumberToken(
                    grammarRoot,
                    unit,
                    path + "." + axisName + ".majorUnit");
            else target.ClearMajorUnit();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "minorUnit"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".minorUnit");
            if (newAxis.Value.TryGetProperty("minorUnit", out var unit))
                target.MinorUnit = ResolveGrammarPositiveNumberToken(
                    grammarRoot,
                    unit,
                    path + "." + axisName + ".minorUnit");
            else target.ClearMinorUnit();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "majorTickMark"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".majorTickMark");
            if (newAxis.Value.TryGetProperty("majorTickMark", out var majorTickMark))
                target.MajorTickMark = ResolveGrammarEnumToken(
                    grammarRoot,
                    majorTickMark,
                    path + "." + axisName + ".majorTickMark",
                    "cross", "in", "out", "none");
            else
                target.ClearMajorTickMark();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "minorTickMark"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".minorTickMark");
            if (newAxis.Value.TryGetProperty("minorTickMark", out var minorTickMark))
                target.MinorTickMark = ResolveGrammarEnumToken(
                    grammarRoot,
                    minorTickMark,
                    path + "." + axisName + ".minorTickMark",
                    "cross", "in", "out", "none");
            else
                target.ClearMinorTickMark();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "tickLabelsVisible"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".tickLabelsVisible");
            if (newAxis.Value.TryGetProperty("tickLabelsVisible", out var visible))
                target.TickLabelsVisible = ResolveGrammarBooleanToken(
                    grammarRoot,
                    visible,
                    path + "." + axisName + ".tickLabelsVisible");
            else
                target.ClearTickLabelsVisible();
            target.ClearTickLabelPosition();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "tickLabelPosition"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".tickLabelPosition");
            if (newAxis.Value.TryGetProperty("tickLabelPosition", out var position))
                target.TickLabelPosition = ResolveGrammarEnumToken(
                    grammarRoot,
                    position,
                    path + "." + axisName + ".tickLabelPosition",
                    "nextTo", "high", "low", "none");
            else
                target.ClearTickLabelPosition();
            target.ClearTickLabelsVisible();
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "axisLine"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".axisLine");
            target.AxisLine = null;
            if (!newAxis.Value.TryGetProperty("axisLine", out var axisLine)) target.ClearAxisLineVisible();
            else if (axisLine.ValueKind is JsonValueKind.True or JsonValueKind.False || IsGrammarTokenReference(axisLine))
                target.AxisLineVisible = ResolveGrammarBooleanToken(
                    grammarRoot,
                    axisLine,
                    path + "." + axisName + ".axisLine");
            else
            {
                target.AxisLineVisible = true;
                target.AxisLine = SourceBoundChartLine(
                    axisLine,
                    path + "." + axisName + ".axisLine",
                    grammarRoot);
            }
            if (newAxis.Value.TryGetProperty("axisLineArrow", out _))
                ApplySourceBoundChartAxisArrows(target, newAxis.Value);
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "axisLineArrow"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".axisLineArrow");
            if (newAxis.Value.TryGetProperty("axisLine", out var axisLine) && axisLine.ValueKind == JsonValueKind.False)
                throw Unsupported(path + "." + axisName + ".axisLineArrow", "arrowheads on a hidden axis line");
            ApplySourceBoundChartAxisArrows(target, newAxis.Value);
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "gridLine"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".gridLine");
            target.MajorGridlineStyle = null;
            if (!newAxis.Value.TryGetProperty("gridLine", out var gridLine))
            {
                target.ClearShowMajorGridlines();
                target.ClearMajorGridlineVisible();
            }
            else if (gridLine.ValueKind is JsonValueKind.True or JsonValueKind.False || IsGrammarTokenReference(gridLine))
            {
                target.ShowMajorGridlines = true;
                var visible = ResolveGrammarBooleanToken(
                    grammarRoot,
                    gridLine,
                    path + "." + axisName + ".gridLine");
                // In DrawingML a majorGridlines node without spPr already
                // means the default visible line.  Keep that canonical form
                // for true instead of recording a redundant explicit
                // MajorGridlineVisible field that the writer will drop.
                if (visible) target.ClearMajorGridlineVisible();
                else target.MajorGridlineVisible = false;
            }
            else
            {
                target.ShowMajorGridlines = true;
                target.ClearMajorGridlineVisible();
                target.MajorGridlineStyle = SourceBoundChartLine(
                    gridLine,
                    path + "." + axisName + ".gridLine",
                    grammarRoot);
            }
            changed = true;
        }
        if (PropertyChanged(oldAxis, newAxis, "minorGridLine"))
        {
            RequireCapability(after, "setChartAxis", path + "." + axisName + ".minorGridLine");
            target.MinorGridlineStyle = null;
            if (!newAxis.Value.TryGetProperty("minorGridLine", out var minorGridLine))
            {
                target.ClearShowMinorGridlines();
                target.ClearMinorGridlineVisible();
            }
            else if (minorGridLine.ValueKind is JsonValueKind.True or JsonValueKind.False || IsGrammarTokenReference(minorGridLine))
            {
                target.ShowMinorGridlines = true;
                var visible = ResolveGrammarBooleanToken(
                    grammarRoot,
                    minorGridLine,
                    path + "." + axisName + ".minorGridLine");
                if (visible) target.ClearMinorGridlineVisible();
                else target.MinorGridlineVisible = false;
            }
            else
            {
                target.ShowMinorGridlines = true;
                target.ClearMinorGridlineVisible();
                target.MinorGridlineStyle = SourceBoundChartLine(
                    minorGridLine,
                    path + "." + axisName + ".minorGridLine",
                    grammarRoot);
            }
            changed = true;
        }
        return changed;
    }

    private static void ApplySourceBoundChartAxisArrows(
        SpreadsheetChartAxisArtifact target,
        JsonElement source)
    {
        if (!source.TryGetProperty("axisLineArrow", out var arrows))
        {
            if (target.AxisLine is null) return;
            target.AxisLine.StartArrow = string.Empty;
            target.AxisLine.EndArrow = string.Empty;
            if (!HasDirectChartLineStyle(target.AxisLine)) target.AxisLine = null;
            return;
        }

        target.AxisLineVisible = true;
        target.AxisLine ??= new SpreadsheetChartLineStyleArtifact();
        target.AxisLine.StartArrow = arrows.TryGetProperty("start", out var start) ? start.GetString()! : string.Empty;
        target.AxisLine.EndArrow = arrows.TryGetProperty("end", out var end) ? end.GetString()! : string.Empty;
    }

    private static bool HasDirectChartLineStyle(SpreadsheetChartLineStyleArtifact line) =>
        line.Color is not null || line.DashStyle != SpreadsheetChartLineDashStyle.Unspecified ||
        line.HasWidthPoints || line.HasOpacityThousandthPercent || line.Cap.Length > 0 || line.Join.Length > 0;

    private static bool ApplyRadarSpokeAxis(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        JsonElement grammarRoot,
        string path)
    {
        var oldSpoke = OptionalProperty(before.Raw, "spokeAxis");
        var newSpoke = OptionalProperty(after.Raw, "spokeAxis");
        if (JsonEqual(oldSpoke, newSpoke)) return false;
        if (oldSpoke is null || newSpoke is null || target.Type != SpreadsheetChartType.Radar ||
            target.XAxis is null || target.YAxis is null)
            throw Unsupported(path + ".spokeAxis", "source-bound radar spoke-axis topology change");

        RequireCapability(after, "setChartAxis", path + ".spokeAxis");
        var changed = false;
        if (PropertyChanged(oldSpoke, newSpoke, "show"))
        {
            var show = !newSpoke.Value.TryGetProperty("show", out var value) ||
                ResolveGrammarBooleanToken(grammarRoot, value, path + ".spokeAxis.show");
            target.XAxis.Visible = show;
            target.YAxis.Visible = show;
            // PresentationChart keeps the legacy showCategoryAxis/showValueAxis
            // carriers alongside the concrete axis objects.  ToSpreadsheet
            // uses those carriers when patching ChartPart; keep them in sync
            // or a radar `show` edit would be silently rewritten to visible.
            target.ShowCategoryAxis = show;
            target.ShowValueAxis = show;
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "min"))
        {
            if (newSpoke.Value.TryGetProperty("min", out var minimum))
                target.YAxis.Minimum = ResolveGrammarNumberToken(grammarRoot, minimum, "size", path + ".spokeAxis.min");
            else target.YAxis.ClearMinimum();
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "max"))
        {
            if (newSpoke.Value.TryGetProperty("max", out var maximum))
                target.YAxis.Maximum = ResolveGrammarNumberToken(grammarRoot, maximum, "size", path + ".spokeAxis.max");
            else target.YAxis.ClearMaximum();
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "majorUnit"))
        {
            if (newSpoke.Value.TryGetProperty("majorUnit", out var majorUnit))
                target.YAxis.MajorUnit = ResolveGrammarPositiveNumberToken(
                    grammarRoot,
                    majorUnit,
                    path + ".spokeAxis.majorUnit");
            else target.YAxis.ClearMajorUnit();
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "axisLine"))
        {
            ApplySourceBoundRadarGuideLine(
                target.XAxis,
                newSpoke.Value,
                "axisLine",
                grammarRoot,
                path + ".spokeAxis.axisLine");
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "gridLine"))
        {
            ApplySourceBoundRadarGuideLine(
                target.YAxis,
                newSpoke.Value,
                "gridLine",
                grammarRoot,
                path + ".spokeAxis.gridLine");
            changed = true;
        }
        if (PropertyChanged(oldSpoke, newSpoke, "label"))
        {
            var oldLabel = oldSpoke.Value.TryGetProperty("label", out var oldLabelValue) ? oldLabelValue : (JsonElement?)null;
            var newLabel = newSpoke.Value.TryGetProperty("label", out var newLabelValue) ? newLabelValue : (JsonElement?)null;
            if (HasRadarLabelTextStyle(oldLabel) || HasRadarLabelTextStyle(newLabel))
                RequireCapability(after, "setChartTextStyle", path + ".spokeAxis.label");
            ApplySourceBoundRadarLabel(target.YAxis, newSpoke.Value, grammarRoot, path + ".spokeAxis.label");
            changed = true;
        }
        return changed;
    }

    private static void ApplySourceBoundRadarGuideLine(
        SpreadsheetChartAxisArtifact target,
        JsonElement owner,
        string propertyName,
        JsonElement grammarRoot,
        string path)
    {
        target.ShowMajorGridlines = true;
        target.MajorGridlineStyle = null;
        if (!owner.TryGetProperty(propertyName, out var line))
        {
            target.ClearMajorGridlineVisible();
            return;
        }
        if (line.ValueKind is JsonValueKind.True or JsonValueKind.False || IsGrammarTokenReference(line))
        {
            var visible = ResolveGrammarBooleanToken(grammarRoot, line, path);
            if (visible) target.ClearMajorGridlineVisible();
            else target.MajorGridlineVisible = false;
            return;
        }
        target.ClearMajorGridlineVisible();
        target.MajorGridlineStyle = SourceBoundChartLine(line, path, grammarRoot);
    }

    private static void ApplySourceBoundRadarLabel(
        SpreadsheetChartAxisArtifact target,
        JsonElement owner,
        JsonElement grammarRoot,
        string path)
    {
        target.NumberFormatCode = string.Empty;
        target.TextStyle = null;
        if (!owner.TryGetProperty("label", out var label))
        {
            target.TickLabelsVisible = true;
            return;
        }
        if (label.ValueKind is JsonValueKind.True or JsonValueKind.False || IsGrammarTokenReference(label))
        {
            target.TickLabelsVisible = label.ValueKind is JsonValueKind.True or JsonValueKind.False
                ? label.GetBoolean()
                : ResolveGrammarBooleanToken(grammarRoot, label, path);
            return;
        }
        target.TickLabelsVisible = true;
        target.NumberFormatCode = label.TryGetProperty("numberFormat", out var format)
            ? ResolveGrammarStringToken(grammarRoot, format, path + ".numberFormat")
            : string.Empty;
        target.TextStyle = HasRadarLabelTextStyle(label) ? SourceBoundChartTextStyle(label, path, grammarRoot) : null;
    }

    private static bool HasRadarLabelTextStyle(JsonElement? label) =>
        label is { ValueKind: JsonValueKind.Object } value && !IsGrammarTokenReference(value) &&
        value.EnumerateObject().Any(property => property.Name != "numberFormat");

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
        string path,
        JsonElement? grammarRoot = null)
    {
        if (owner is null || !owner.Value.TryGetProperty(property, out var source)) return null;
        return SourceBoundChartTextStyle(source, path, grammarRoot);
    }

    private static SpreadsheetChartTextStyleArtifact SourceBoundChartTextStyle(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var output = new SpreadsheetChartTextStyleArtifact();
        if (source.TryGetProperty("fontSize", out var fontSize))
        {
            var resolved = grammarRoot is { } root
                ? ResolveGrammarPositiveSizeToken(root, fontSize, path + ".fontSize")
                : fontSize.GetDouble();
            output.FontSizePoints = resolved;
        }
        if (source.TryGetProperty("fontFamily", out var fontFamily))
            output.FontFamily = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, fontFamily, path + ".fontFamily")
                : fontFamily.GetString()!;
        if (source.TryGetProperty("fontFamilyEastAsia", out var eastAsia))
            output.FontFamilyEastAsia = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, eastAsia, path + ".fontFamilyEastAsia")
                : eastAsia.GetString()!;
        if (source.TryGetProperty("fontFamilyComplexScript", out var complexScript))
            output.FontFamilyComplexScript = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, complexScript, path + ".fontFamilyComplexScript")
                : complexScript.GetString()!;
        if (source.TryGetProperty("bold", out var bold))
            output.Bold = grammarRoot is { } root
                ? ResolveGrammarBooleanToken(root, bold, path + ".bold")
                : bold.GetBoolean();
        if (source.TryGetProperty("italic", out var italic))
            output.Italic = grammarRoot is { } root
                ? ResolveGrammarBooleanToken(root, italic, path + ".italic")
                : italic.GetBoolean();
        if (source.TryGetProperty("underline", out var underline))
            output.Underline = PpjAuthoredPresentationCompiler.NativeUnderline(underline.GetString()!);
        if (source.TryGetProperty("alignment", out var alignment))
            output.Alignment = PpjAuthoredPresentationCompiler.NativeChartAlignment(alignment.GetString()!);
        if (source.TryGetProperty("fill", out var fill))
            ApplyChartTextFill(output, SourceBoundChartFill(fill, path + ".fill", grammarRoot));
        if (source.TryGetProperty("color", out var color))
        {
            var resolved = grammarRoot is { } root
                ? ResolveGrammarColorValue(root, color, path + ".color")
                : ParseSourceBoundColor(color, path + ".color");
            output.Fill = null;
            output.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) output.OpacityThousandthPercent = Unit(resolved.Alpha);
            else output.ClearOpacityThousandthPercent();
        }
        return output;
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

    private static SpreadsheetChartLineStyleArtifact SourceBoundChartLine(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var colorValue = source.GetProperty("color");
        var (color, colorAlpha) = grammarRoot is { } root
            ? ResolveGrammarColorValue(root, colorValue, path + ".color")
            : (Rgb(colorValue, path + ".color"), 1d);
        var output = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = color },
            DashStyle = source.TryGetProperty("dash", out var dash) ? dash.GetString() switch
            {
                "solid" => SpreadsheetChartLineDashStyle.Solid,
                "dash" or "long-dash" => SpreadsheetChartLineDashStyle.Dashed,
                "dot" => SpreadsheetChartLineDashStyle.Dotted,
                "dash-dot" => SpreadsheetChartLineDashStyle.DashDot,
                _ => throw Unsupported(path + ".dash", "unsupported chart line dash"),
            } : SpreadsheetChartLineDashStyle.Solid,
            WidthPoints = grammarRoot is { } widthRoot
                ? ResolveGrammarSizeToken(widthRoot, source.GetProperty("width"), path + ".width")
                : source.GetProperty("width").GetDouble(),
            Cap = source.TryGetProperty("cap", out var cap) ? cap.GetString()! : string.Empty,
            Join = source.TryGetProperty("join", out var join) ? join.GetString()! : string.Empty,
        };
        if (source.TryGetProperty("opacity", out var opacity))
        {
            var resolvedOpacity = grammarRoot is { } opacityRoot
                ? ResolveGrammarOpacityToken(opacityRoot, opacity, path + ".opacity")
                : opacity.GetDouble();
            if (resolvedOpacity < 1) output.OpacityThousandthPercent = Unit(resolvedOpacity);
        }
        else if (colorAlpha < 1)
            output.OpacityThousandthPercent = Unit(colorAlpha);
        return output;
    }

    private static SpreadsheetChartSurfaceFill? SourceBoundChartFill(
        JsonElement? owner,
        string property,
        string path,
        JsonElement? grammarRoot = null)
    {
        if (owner is null || !owner.Value.TryGetProperty(property, out var fill)) return null;
        return SourceBoundChartFill(fill, path, grammarRoot);
    }

    private static SpreadsheetChartSurfaceFill SourceBoundChartFill(JsonElement fill, string path, JsonElement? grammarRoot = null)
    {
        var type = fill.GetProperty("type").GetString();
        if (type == "none") return new SpreadsheetChartSurfaceFill { NoFill = true };
        if (type == "gradient") return new SpreadsheetChartSurfaceFill
        {
            GradientFill = BuildGradientFill(
                fill,
                path,
                grammarRoot is { } root
                    ? (color, colorPath) => ResolveGrammarColorValue(root, color, colorPath)
                    : null),
        };
        if (type != "solid") throw Unsupported(path, "source-bound chart paint outside none, solid, or bounded gradient fill");
        (string Rgb, double Alpha) resolvedColor = grammarRoot is { } colorRoot
            ? ResolveGrammarColorValue(colorRoot, fill.GetProperty("color"), path + ".color")
            : (Rgb(fill.GetProperty("color"), path + ".color"), 1d);
        var output = new SpreadsheetChartSurfaceFill
        {
            SolidRgb = resolvedColor.Rgb,
        };
        if (fill.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(grammarRoot is { } root
                ? ResolveGrammarOpacityToken(root, opacity, path + ".opacity")
                : opacity.GetDouble());
        else if (resolvedColor.Alpha < 1)
            output.OpacityThousandthPercent = Unit(resolvedColor.Alpha);
        return output;
    }

    private static PresentationChartFrame SourceBoundChartFrame(
        JsonElement source,
        string path,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot,
        double frameWidth,
        double frameHeight)
    {
        var output = new PresentationChartFrame();
        if (source.TryGetProperty("fill", out var fill))
        {
            if (fill.TryGetProperty("type", out var fillType) && fillType.GetString() == "image")
                output.ImageFill = PpjImagePaintLowering.Build(
                    fill,
                    frameWidth,
                    frameHeight,
                    id => ResolveAsset(assets, id, path + ".fill.asset"),
                    id => assetDimensions.TryGetValue(id, out var dimensions) ? dimensions : null,
                    path + ".fill",
                    resolveOpacity: opacity => ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".fill.opacity"),
                    resolveFit: fit => ResolveGrammarStringToken(grammarRoot, fit, path + ".fill.fit"));
            else
                output.Fill = SourceBoundChartFill(fill, path + ".fill", grammarRoot);
        }
        if (source.TryGetProperty("stroke", out var stroke))
            output.Line = SourceBoundChartLine(stroke, path + ".stroke", grammarRoot);
        if (source.TryGetProperty("shadow", out var shadow))
            output.Shadow = SourceBoundShadow(shadow, path + ".shadow", "chart frame", grammarRoot);
        PptxChartFrameCodec.Validate(output, path);
        return output;
    }

    private static void EnsureSourceBoundChartFrameImageTopology(
        JsonElement? before,
        JsonElement? after,
        string path)
    {
        var oldFill = before is { } oldFrame && oldFrame.TryGetProperty("fill", out var oldValue)
            ? oldValue
            : (JsonElement?)null;
        var newFill = after is { } newFrame && newFrame.TryGetProperty("fill", out var newValue)
            ? newValue
            : (JsonElement?)null;
        var oldIsImage = IsImageFill(oldFill);
        var newIsImage = IsImageFill(newFill);
        if (!oldIsImage && !newIsImage) return;
        if (!oldIsImage || !newIsImage)
            throw Unsupported(path, "source-bound chart image frame fill cannot be added, removed, or changed to another paint topology");

        // The bounded chart-frame codec owns the chart-part relationship and
        // removes the previous image part when it becomes unreferenced.  Asset
        // replacement is therefore safe here as long as the image-fill
        // topology itself remains stable.
    }

    private static bool IsImageFill(JsonElement? fill) =>
        fill is { ValueKind: JsonValueKind.Object } value &&
        value.TryGetProperty("type", out var type) &&
        type.ValueKind == JsonValueKind.String &&
        string.Equals(type.GetString(), "image", StringComparison.Ordinal);

    private static PresentationImageBorder SourceBoundImageBorder(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "image border must be an object");
        var color = source.GetProperty("color");
        var width = source.GetProperty("width");
        var colorAlpha = 1d;
        var output = new PresentationImageBorder
        {
            WidthEmu = Emu(grammarRoot is { } widthRoot
                ? ResolveGrammarSizeToken(widthRoot, width, path + ".width")
                : width.GetDouble()),
            Style = source.TryGetProperty("dash", out var dash) ? dash.GetString() switch
            {
                "solid" => "solid",
                "dash" => "dashed",
                "dot" => "dotted",
                "dash-dot" => "dash-dot",
                "long-dash" => "dash-dot-dot",
                _ => throw Unsupported(path + ".dash", "unsupported image border dash style"),
            } : "solid",
            Cap = OptionalString(source, "cap") ?? string.Empty,
            Join = OptionalString(source, "join") ?? string.Empty,
        };
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            token.ValueKind == JsonValueKind.String)
        {
            var tokenName = token.GetString()!;
            if (grammarRoot is { } colorRoot && TryDeclaredGrammarToken(colorRoot, tokenName, out var definition))
            {
                if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {tokenName} for {path}.color must declare kind color.", path + ".color");
                var resolved = ResolveGrammarColorValue(colorRoot, color, path + ".color");
                output.ColorRgb = resolved.Rgb;
                colorAlpha = resolved.Alpha;
            }
            else
            {
                // An undeclared token retains the source-bound meaning of a
                // standard DrawingML theme color; only declared tokens are
                // interpreted as PPJ design grammar references.
                output.ColorScheme = PptxColor.NormalizeScheme(tokenName);
            }
        }
        else
        {
            var resolved = grammarRoot is { } literalRoot
                ? ResolveGrammarColorValue(literalRoot, color, path + ".color")
                : ParseSourceBoundColor(color, path + ".color");
            output.ColorRgb = resolved.Rgb;
            colorAlpha = resolved.Alpha;
        }
        if (source.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(grammarRoot is { } opacityRoot
                ? ResolveGrammarOpacityToken(opacityRoot, opacity, path + ".opacity")
                : opacity.GetDouble());
        else if (colorAlpha < 1)
            output.OpacityThousandthPercent = Unit(colorAlpha);
        if (output.WidthEmu is < 0 or > int.MaxValue ||
            output.Style is not ("solid" or "dashed" or "dotted" or "dash-dot" or "dash-dot-dot") ||
            output.Cap is not ("" or "flat" or "round" or "square") ||
            output.Join is not ("" or "miter" or "round" or "bevel") ||
            output.HasOpacityThousandthPercent && output.OpacityThousandthPercent > 100_000)
            throw Unsupported(path, "image border uses unsupported width, dash, cap, join, or opacity");
        return output;
    }

    private static PresentationShadow SourceBoundShadow(
        JsonElement source,
        string path,
        string subject = "image",
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, $"{subject} shadow must be an object");
        var color = source.GetProperty("color");
        var colorAlpha = 1d;
        var output = new PresentationShadow
        {
            BlurRadiusEmu = Emu(source.GetProperty("blur").GetDouble()),
            DistanceEmu = Emu(source.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = RotationAngle(NormalizeAngle(source.GetProperty("angle").GetDouble())),
        };
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            token.ValueKind == JsonValueKind.String)
        {
            var tokenName = token.GetString()!;
            if (grammarRoot is { } root && TryDeclaredGrammarToken(root, tokenName, out var definition))
            {
                if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {tokenName} for {path}.color must declare kind color.", path + ".color");
                var resolved = ResolveGrammarColorValue(root, color, path + ".color");
                output.ColorRgb = resolved.Rgb;
                colorAlpha = resolved.Alpha;
            }
            else
            {
                // An undeclared token keeps the source-bound meaning of a
                // standard DrawingML theme color.  Do not reinterpret an
                // unknown token as a grammar reference.
                output.ColorScheme = PptxColor.NormalizeScheme(tokenName);
            }
        }
        else
        {
            var resolved = grammarRoot is { } root
                ? ResolveGrammarColorValue(root, color, path + ".color")
                : ParseSourceBoundColor(color, path + ".color");
            output.ColorRgb = resolved.Rgb;
            colorAlpha = resolved.Alpha;
        }
        if (source.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(grammarRoot is { } root
                ? ResolveGrammarOpacityToken(root, opacity, path + ".opacity")
                : opacity.GetDouble());
        else if (colorAlpha < 1)
            output.OpacityThousandthPercent = Unit(colorAlpha);
        if (source.TryGetProperty("alignment", out var alignment)) output.Alignment = alignment.GetString()!;
        if (source.TryGetProperty("rotateWithShape", out var rotateWithShape)) output.RotateWithShape = rotateWithShape.GetBoolean();
        PptxShadowCodec.Validate(output, path, subject);
        return output;
    }

    private static PresentationGlow SourceBoundGlow(
        JsonElement source,
        string path,
        string subject = "shape",
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, $"{subject} glow must be an object");
        var color = source.GetProperty("color");
        var colorAlpha = 1d;
        var radius = grammarRoot is { } radiusRoot
            ? ResolveGrammarSizeToken(radiusRoot, source.GetProperty("radius"), path + ".radius")
            : source.GetProperty("radius").GetDouble();
        var output = new PresentationGlow { RadiusEmu = Emu(radius) };
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            token.ValueKind == JsonValueKind.String)
        {
            var tokenName = token.GetString()!;
            if (grammarRoot is { } colorRoot && TryDeclaredGrammarToken(colorRoot, tokenName, out var definition))
            {
                if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {tokenName} for {path}.color must declare kind color.", path + ".color");
                var resolved = ResolveGrammarColorValue(colorRoot, color, path + ".color");
                output.ColorRgb = resolved.Rgb;
                colorAlpha = resolved.Alpha;
            }
            else
            {
                output.ColorScheme = PptxColor.NormalizeScheme(tokenName);
            }
        }
        else
        {
            var resolved = grammarRoot is { } literalRoot
                ? ResolveGrammarColorValue(literalRoot, color, path + ".color")
                : ParseSourceBoundColor(color, path + ".color");
            output.ColorRgb = resolved.Rgb;
            colorAlpha = resolved.Alpha;
        }
        if (source.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(grammarRoot is { } opacityRoot
                ? ResolveGrammarOpacityToken(opacityRoot, opacity, path + ".opacity")
                : opacity.GetDouble());
        else if (colorAlpha < 1)
            output.OpacityThousandthPercent = Unit(colorAlpha);
        PptxGlowCodec.Validate(output, path, subject);
        return output;
    }

    private static PresentationInnerShadow SourceBoundInnerShadow(
        JsonElement source,
        string path,
        string subject = "shape",
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, $"{subject} innerShadow must be an object");
        var color = source.GetProperty("color");
        var colorAlpha = 1d;
        var output = new PresentationInnerShadow
        {
            BlurRadiusEmu = Emu(source.GetProperty("blur").GetDouble()),
            DistanceEmu = Emu(source.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = RotationAngle(NormalizeAngle(source.GetProperty("angle").GetDouble())),
        };
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            token.ValueKind == JsonValueKind.String)
        {
            var tokenName = token.GetString()!;
            if (grammarRoot is { } root && TryDeclaredGrammarToken(root, tokenName, out var definition))
            {
                if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {tokenName} for {path}.color must declare kind color.", path + ".color");
                var resolved = ResolveGrammarColorValue(root, color, path + ".color");
                output.ColorRgb = resolved.Rgb;
                colorAlpha = resolved.Alpha;
            }
            else
            {
                output.ColorScheme = PptxColor.NormalizeScheme(tokenName);
            }
        }
        else
        {
            var resolved = grammarRoot is { } root
                ? ResolveGrammarColorValue(root, color, path + ".color")
                : ParseSourceBoundColor(color, path + ".color");
            output.ColorRgb = resolved.Rgb;
            colorAlpha = resolved.Alpha;
        }
        if (source.TryGetProperty("opacity", out var opacity))
            output.OpacityThousandthPercent = Unit(grammarRoot is { } root
                ? ResolveGrammarOpacityToken(root, opacity, path + ".opacity")
                : opacity.GetDouble());
        else if (colorAlpha < 1)
            output.OpacityThousandthPercent = Unit(colorAlpha);
        PptxInnerShadowCodec.Validate(output, path, subject);
        return output;
    }

    private static PresentationReflection SourceBoundReflection(
        JsonElement source,
        string path,
        string subject = "shape",
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, $"{subject} reflection must be an object");
        var output = new PresentationReflection
        {
            BlurRadiusEmu = Emu(source.GetProperty("blur").GetDouble()),
            StartOpacityThousandthPercent = Unit(grammarRoot is { } startRoot
                ? ResolveGrammarOpacityToken(startRoot, source.GetProperty("startOpacity"), path + ".startOpacity")
                : source.GetProperty("startOpacity").GetDouble()),
            EndOpacityThousandthPercent = Unit(grammarRoot is { } endRoot
                ? ResolveGrammarOpacityToken(endRoot, source.GetProperty("endOpacity"), path + ".endOpacity")
                : source.GetProperty("endOpacity").GetDouble()),
            DistanceEmu = Emu(source.GetProperty("distance").GetDouble()),
            DirectionAngle60000 = RotationAngle(NormalizeAngle(source.GetProperty("angle").GetDouble())),
        };
        PptxReflectionCodec.Validate(output, path, subject);
        return output;
    }

    private static PresentationSoftEdge SourceBoundSoftEdge(
        JsonElement source,
        string path,
        string subject = "shape",
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, $"{subject} softEdge must be an object");
        var radius = grammarRoot is { } root
            ? ResolveGrammarSizeToken(root, source.GetProperty("radius"), path + ".radius")
            : source.GetProperty("radius").GetDouble();
        var output = new PresentationSoftEdge { RadiusEmu = Emu(radius) };
        PptxSoftEdgeCodec.Validate(output, path, subject);
        return output;
    }

    private static double NormalizeAngle(double degrees)
    {
        var normalized = degrees % 360;
        return normalized < 0 ? normalized + 360 : normalized;
    }

    private static bool ApplyTableElement(
        PpjProgramModel program,
        PpjTableElementModel before,
        PpjTableElementModel after,
        PresentationTable target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "columns", "rows", "style", "accessibility");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "columns"))
        {
            RequireCapability(after, "setTableGeometry", path + ".columns");
            if (before.Columns.Count != after.Columns.Count || before.Columns.Count != target.ColumnWidthsEmu.Count)
                throw Unsupported(path + ".columns", "table column topology change");
            for (var column = 0; column < before.Columns.Count; column++)
            {
                var oldColumn = before.Columns[column];
                var newColumn = after.Columns[column];
                if (oldColumn.Id != newColumn.Id)
                    throw Unsupported($"{path}.columns[{column}]", "table column identity change");
                target.ColumnWidthsEmu[column] = Emu(newColumn.Width);
            }
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "style"))
        {
            RequireCapability(after, "setTableStyle", path + ".style");
            ApplyTableStyle(OptionalProperty(after.Raw, "style"), target, path + ".style");
            changed = true;
        }
        if (!PropertyChanged(before.Raw, after.Raw, "rows")) return changed;
        if (before.Rows.Count != after.Rows.Count || before.Rows.Count != target.Rows.Count)
            throw Unsupported(path + ".rows", "table row topology change");
        for (var row = 0; row < before.Rows.Count; row++)
        {
            var oldRow = before.Rows[row];
            var newRow = after.Rows[row];
            if (oldRow.Id != newRow.Id || oldRow.Cells.Count != newRow.Cells.Count)
                throw Unsupported($"{path}.rows[{row}]", "table row topology change");
            if (oldRow.Height != newRow.Height)
            {
                RequireCapability(after, "setTableGeometry", $"{path}.rows[{row}].height");
                if (newRow.Height is null)
                    throw Unsupported($"{path}.rows[{row}].height", "source-bound table row height must remain explicit");
                target.Rows[row].HeightEmu = Emu(newRow.Height.Value);
                changed = true;
            }
            for (var cell = 0; cell < oldRow.Cells.Count; cell++)
            {
                var oldCell = oldRow.Cells[cell];
                var newCell = newRow.Cells[cell];
                var cellPath = $"{path}.rows[{row}].cells[{cell}]";
                RequireEqualExcept(oldCell.Raw, newCell.Raw, cellPath, "id", "text", "rowSpan", "columnSpan", "fill", "borders", "textStyle");
                if (oldCell.Id != newCell.Id || oldCell.RowSpan != newCell.RowSpan || oldCell.ColumnSpan != newCell.ColumnSpan)
                    throw Unsupported(cellPath, "table cell topology change");
                var rawTextChanged = PropertyChanged(oldCell.Raw, newCell.Raw, "text");
                var oldStructuredText = oldCell.Raw.TryGetProperty("text", out var oldTextRaw) && oldTextRaw.ValueKind == JsonValueKind.Object;
                var newStructuredText = newCell.Raw.TryGetProperty("text", out var newTextRaw) && newTextRaw.ValueKind == JsonValueKind.Object;
                if (oldStructuredText && rawTextChanged && !newStructuredText)
                    throw Unsupported(cellPath + ".text", "mixed-run table-cell text must retain its structured text body");
                var textChanged = !TextEqual(oldCell.Text, newCell.Text);
                var textBodyStyleChanged = rawTextChanged && newStructuredText &&
                    TextBodyStyleChanged(oldTextRaw, newTextRaw);
                var styleChanged = PropertyChanged(oldCell.Raw, newCell.Raw, "fill") ||
                    PropertyChanged(oldCell.Raw, newCell.Raw, "borders") ||
                    PropertyChanged(oldCell.Raw, newCell.Raw, "textStyle");
                if (styleChanged)
                {
                    if (!TryProjectedCellCoordinates(oldCell.Id, out var styleRow, out var styleColumn) ||
                        styleRow != row || styleColumn < 0 || styleColumn >= target.Rows[row].Cells.Count)
                        throw Unsupported(cellPath, "table cell style requires a canonical visible cell coordinate");
                    var physicalCell = target.Rows[styleRow].Cells[styleColumn];
                    if (PropertyChanged(oldCell.Raw, newCell.Raw, "fill"))
                    {
                        RequireCapabilityField(after, "setTableCellStyle", "table.cell.fill", cellPath + ".fill");
                        physicalCell.Fill = newCell.Raw.TryGetProperty("fill", out var fill)
                            ? SourceBoundTableCellFill(
                                fill,
                                cellPath + ".fill",
                                target.ColumnWidthsEmu[styleColumn] / 12_700d,
                                target.Rows[styleRow].HeightEmu / 12_700d,
                                assets,
                                assetDimensions,
                                program.Root)
                            : null;
                    }
                    if (PropertyChanged(oldCell.Raw, newCell.Raw, "borders"))
                    {
                        RequireCapabilityField(after, "setTableCellStyle", "table.cell.borders", cellPath + ".borders");
                        physicalCell.Borders = newCell.Raw.TryGetProperty("borders", out var borders)
                            ? SourceBoundTableCellBorders(borders, cellPath + ".borders", program.Root)
                            : null;
                    }
                    if (PropertyChanged(oldCell.Raw, newCell.Raw, "textStyle"))
                    {
                        RequireCapabilityField(after, "setTableCellStyle", "table.cell.textStyle", cellPath + ".textStyle");
                        physicalCell.TextStyle = newCell.Raw.TryGetProperty("textStyle", out var textStyle)
                            ? SourceBoundTableCellTextStyle(textStyle, cellPath + ".textStyle", program.Root)
                            : null;
                    }
                    changed = true;
                }
                if (textBodyStyleChanged)
                {
                    RequireCapabilityField(after, "setTableCellStyle", "table.cell.textStyle", cellPath + ".text");
                    if (!TryProjectedCellCoordinates(oldCell.Id, out var bodyStyleRow, out var bodyStyleColumn) ||
                        bodyStyleRow != row || bodyStyleColumn < 0 || bodyStyleColumn >= target.Rows[row].Cells.Count)
                        throw Unsupported(cellPath + ".text", "mixed-run table-cell text requires a canonical visible cell coordinate");
                    target.Rows[bodyStyleRow].Cells[bodyStyleColumn].TextBody =
                        SourceBoundTableCellTextBody(newTextRaw, cellPath + ".text", program.Root);
                    target.Rows[bodyStyleRow].Cells[bodyStyleColumn].Text =
                        PptxTextCodec.Flatten(target.Rows[bodyStyleRow].Cells[bodyStyleColumn].TextBody);
                    changed = true;
                }
                if (!textChanged) continue;
                RequireCapability(after, "replaceText", cellPath + ".text");
                if (!newStructuredText && newCell.Text.PlainText is null)
                    throw Unsupported(cellPath + ".text", "rich or noncanonical imported table-cell text");
                if (!TryProjectedCellCoordinates(oldCell.Id, out var physicalRow, out var physicalColumn) ||
                    physicalRow != row || physicalColumn < 0 || physicalColumn >= target.Rows[row].Cells.Count)
                    throw Unsupported(cellPath + ".text", "rich or noncanonical imported table-cell text");
                if (newStructuredText)
                {
                    target.Rows[physicalRow].Cells[physicalColumn].TextBody =
                        SourceBoundTableCellTextBody(newTextRaw, cellPath + ".text", program.Root);
                    target.Rows[physicalRow].Cells[physicalColumn].Text =
                        PptxTextCodec.Flatten(target.Rows[physicalRow].Cells[physicalColumn].TextBody);
                }
                else target.Rows[physicalRow].Cells[physicalColumn].Text = newCell.Text.PlainText!;
                changed = true;
            }
        }
        return changed;
    }

    private static PresentationTextBody SourceBoundTableCellTextBody(
        JsonElement source,
        string path,
        JsonElement programRoot)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "source-bound mixed-run table-cell text must be a structured text body");
        PresentationTextBody body;
        try
        {
            body = PpjAuthoredPresentationCompiler.BuildSourceBoundTextBody(source, programRoot);
        }
        catch (CodecException error)
        {
            throw Unsupported(path, $"mixed-run table-cell text contains unsupported direct properties ({error.Message})");
        }
        if (!PptxTableCodec.IsBoundedMixedRunTextBody(body))
            throw Unsupported(path, "mixed-run table-cell text must contain only fixed-topology plain runs with bounded direct styles");
        return body;
    }

    private static PresentationTableCellFill SourceBoundTableCellFill(
        JsonElement source,
        string path,
        double frameWidth,
        double frameHeight,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "source-bound table cell fill must be an object");
        if (source.TryGetProperty("type", out var type) && type.GetString() == "image")
            return new PresentationTableCellFill
            {
                ImagePaint = BuildImagePaint(
                    source,
                    frameWidth,
                    frameHeight,
                    assets,
                    assetDimensions,
                    path,
                    resolveOpacity: opacity => ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".opacity"),
                    resolveFit: fit => ResolveGrammarStringToken(grammarRoot, fit, path + ".fit")),
            };
        var fill = SourceBoundChartFill(source, path, grammarRoot);
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            return new PresentationTableCellFill { NoFill = true };
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb)
        {
            var output = new PresentationTableCellFill { SolidRgb = fill.SolidRgb };
            if (fill.HasOpacityThousandthPercent)
                output.OpacityThousandthPercent = fill.OpacityThousandthPercent;
            return output;
        }
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            return new PresentationTableCellFill { GradientFill = fill.GradientFill.Clone() };
        throw Unsupported(path, "source-bound table cell paint outside no-fill, solid, or bounded gradient fill");
    }

    private static PresentationTableCellBorders SourceBoundTableCellBorders(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "table cell borders must be an object");
        var output = new PresentationTableCellBorders();
        if (source.TryGetProperty("left", out var left)) output.Left = SourceBoundChartLine(left, path + ".left", grammarRoot);
        if (source.TryGetProperty("top", out var top)) output.Top = SourceBoundChartLine(top, path + ".top", grammarRoot);
        if (source.TryGetProperty("right", out var right)) output.Right = SourceBoundChartLine(right, path + ".right", grammarRoot);
        if (source.TryGetProperty("bottom", out var bottom)) output.Bottom = SourceBoundChartLine(bottom, path + ".bottom", grammarRoot);
        return output;
    }

    private static PresentationTextStyle SourceBoundTableCellTextStyle(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "table cell textStyle must be an object");
        var value = source.TryGetProperty("defaultText", out var defaultText) ? defaultText : source;
        if (value.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "table cell textStyle.defaultText must be an object");
        var allowed = new HashSet<string>(StringComparer.Ordinal)
        {
            "bold", "italic", "size", "fontFamily", "font", "fontFamilyEastAsia", "fontFamilyComplexScript", "color", "underline", "strike",
        };
        foreach (var property in value.EnumerateObject())
            if (!allowed.Contains(property.Name))
                throw Unsupported(path + ".defaultText." + property.Name, "source-bound table cell text style is outside the direct run profile");
        var output = new PresentationTextStyle();
        if (value.TryGetProperty("bold", out var bold))
            output.Bold = grammarRoot is { } root
                ? ResolveGrammarBooleanToken(root, bold, path + ".defaultText.bold")
                : bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic))
            output.Italic = grammarRoot is { } root
                ? ResolveGrammarBooleanToken(root, italic, path + ".defaultText.italic")
                : italic.GetBoolean();
        if (value.TryGetProperty("size", out var size))
        {
            var points = grammarRoot is { } root
                ? ResolveGrammarPositiveSizeToken(root, size, path + ".defaultText.size")
                : size.GetDouble();
            if (!double.IsFinite(points) || points <= 0 || points > 768)
                throw Unsupported(path + ".defaultText.size", "table cell text size must be finite and between 0 and 768 points");
            output.FontSizePoints = points;
        }
        if (value.TryGetProperty("fontFamily", out var family))
            output.FontFamily = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, family, path + ".defaultText.fontFamily")
                : family.GetString()!;
        else if (value.TryGetProperty("font", out var font))
            throw Unsupported(path + ".defaultText.font", "source-bound table cell text style cannot resolve design font tokens");
        if (value.TryGetProperty("fontFamilyEastAsia", out var eastAsia))
            output.FontFamilyEastAsia = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, eastAsia, path + ".defaultText.fontFamilyEastAsia")
                : eastAsia.GetString()!;
        if (value.TryGetProperty("fontFamilyComplexScript", out var complexScript))
            output.FontFamilyComplexScript = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, complexScript, path + ".defaultText.fontFamilyComplexScript")
                : complexScript.GetString()!;
        if (value.TryGetProperty("color", out var color))
        {
            if (grammarRoot is { } root)
            {
                var resolved = ResolveGrammarColorValue(root, color, path + ".defaultText.color");
                output.ColorRgb = resolved.Rgb;
                if (resolved.Alpha < 1) output.ColorOpacityThousandthPercent = Unit(resolved.Alpha);
            }
            else
            {
                if (color.ValueKind != JsonValueKind.String)
                    throw Unsupported(path + ".defaultText.color", "source-bound table cell text style requires direct RGB color");
                output.ColorRgb = Rgb(color, path + ".defaultText.color");
            }
        }
        if (value.TryGetProperty("underline", out var underline)) output.Underline = underline.GetString()!;
        if (value.TryGetProperty("strike", out var strike)) output.Strike = strike.GetString()!;
        return output;
    }

    private static void ApplyTableStyle(JsonElement? style, PresentationTable target, string path)
    {
        if (style is { } value && value.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "table style must be an object");

        var supported = new HashSet<string>(StringComparer.Ordinal)
        {
            "headerRows",
            "bandedRows",
            "bandedColumns",
            "firstColumnEmphasis",
            "lastColumnEmphasis",
            "lastRow",
        };
        if (style is { } objectValue)
        {
            foreach (var property in objectValue.EnumerateObject())
                if (!supported.Contains(property.Name))
                    throw Unsupported(path + "." + property.Name, "source-bound table style field is not represented by the native table-property profile");
        }

        var headerRows = OptionalTableStyleInt(style, "headerRows", path);
        if (headerRows is null) target.ClearFirstRow();
        else if (headerRows is 0 or 1) target.FirstRow = headerRows == 1;
        else throw Unsupported(path + ".headerRows", "source-bound table style supports only zero or one header row");

        ApplyTableStyleBool(style, "bandedRows", target, path, value =>
        {
            if (value is { } flag) target.BandedRows = flag;
            else target.ClearBandedRows();
        });
        ApplyTableStyleBool(style, "bandedColumns", target, path, value =>
        {
            if (value is { } flag) target.BandedColumns = flag;
            else target.ClearBandedColumns();
        });
        ApplyTableStyleBool(style, "firstColumnEmphasis", target, path, value =>
        {
            if (value is { } flag) target.FirstColumn = flag;
            else target.ClearFirstColumn();
        });
        ApplyTableStyleBool(style, "lastColumnEmphasis", target, path, value =>
        {
            if (value is { } flag) target.LastColumn = flag;
            else target.ClearLastColumn();
        });
        ApplyTableStyleBool(style, "lastRow", target, path, value =>
        {
            if (value is { } flag) target.LastRow = flag;
            else target.ClearLastRow();
        });
    }

    private static int? OptionalTableStyleInt(JsonElement? style, string name, string path)
    {
        if (style is not { } value || !value.TryGetProperty(name, out var property)) return null;
        if (property.ValueKind != JsonValueKind.Number || !property.TryGetInt32(out var result))
            throw Unsupported(path + "." + name, "table style headerRows must be an integer");
        return result;
    }

    private static void ApplyTableStyleBool(
        JsonElement? style,
        string name,
        PresentationTable target,
        string path,
        Action<bool?> apply)
    {
        if (style is not { } value || !value.TryGetProperty(name, out var property))
        {
            apply(null);
            return;
        }
        if (property.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw Unsupported(path + "." + name, "table style flag must be a boolean");
        apply(property.GetBoolean());
    }

    private static bool ApplyConnectorElement(PpjProgramModel program, PpjConnectorElementModel before, PpjConnectorElementModel after, PresentationConnector target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "stroke", "accessibility");
        var oldFrame = before.Frame;
        var changed = ApplyConnectorFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".stroke");
            ApplyConnectorStroke(after.Raw.GetProperty("stroke"), target, program.Root, path + ".stroke");
            changed = true;
        }
        _ = oldFrame;
        return changed;
    }

    private static bool ApplyGroupElement(
        PpjProgramModel program,
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
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "childFrame", "elements", "readingOrder", "accessibility");
        var changed = ApplyFrame(before, after, target, path);
        if (changed) mutations.SemanticChanges = true;
        if (ChildFrameChanged(before, after))
        {
            target.ChildLeftEmu = Emu(after.ChildFrame.X);
            target.ChildTopEmu = Emu(after.ChildFrame.Y);
            target.ChildWidthEmu = Emu(after.ChildFrame.Width);
            target.ChildHeightEmu = Emu(after.ChildFrame.Height);
            changed = true;
            mutations.SemanticChanges = true;
        }
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
        var requestedElementOrder = after.ReadingOrder.Count == 0
            ? requestedIds
            : after.ReadingOrder.ToArray();
        if (requestedElementOrder.Length != requestedIds.Length ||
            requestedElementOrder.Distinct(StringComparer.Ordinal).Count() != requestedElementOrder.Length ||
            !requestedElementOrder.ToHashSet(StringComparer.Ordinal).SetEquals(requestedIds))
            throw Unsupported(path + ".readingOrder", "group reading order must be a complete permutation of the direct source children");
        for (var index = 0; index < after.Elements.Count; index++)
        {
            var sourceChild = sourceChildren[after.Elements[index].Id];
            var child = sourceChild.Wire;
            var childPath = shapeTreePath.Concat([child.Source?.ShapeTreeIndex ?? checked((uint)index)]).ToArray();
            changed |= ApplyElement(program, sourceChild.Program, after.Elements[index], child, slide, childPath, assets, assetDimensions, nativeLeafBindings, changedNodeIds, mutations, $"{path}.elements[{index}]");
        }
        var sourceOrder = before.Elements.Select(element => element.Id).ToArray();
        if (!sourceOrder.SequenceEqual(requestedElementOrder, StringComparer.Ordinal))
        {
            for (var index = 0; index < requestedElementOrder.Length; index++)
            {
                if (sourceOrder[index] == requestedElementOrder[index]) continue;
                var requestedElement = after.Elements.Single(element =>
                    element.Id.Equals(requestedElementOrder[index], StringComparison.Ordinal));
                RequireCapability(requestedElement, "reorder", $"{path}.readingOrder[{index}]");
                changedNodeIds.Add(requestedElement.Id);
            }
            mutations.SemanticChanges = true;
            changed = true;
        }
        var requestedChildren = requestedElementOrder
            .Select(id => sourceChildren[id].Wire)
            .ToList();
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
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "visibleText", "accessibility");
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

    private static bool ApplySourceSmartArtElement(
        PpjSmartArtElementModel before,
        PpjSmartArtElementModel after,
        PresentationOpaqueElement target,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "nodes", "accessibility");
        if (before.Mode != "source-bound" || after.Mode != "source-bound" || target.DiagramText is null)
            throw Unsupported(path, "source-bound SmartArt identity or mode change");
        if (before.Nodes.Count != after.Nodes.Count || before.Nodes.Count != target.DiagramText.Nodes.Count)
            throw Unsupported(path + ".nodes", "source-bound SmartArt node topology change");

        var changed = ApplyFrame(before, after, target, path);
        for (var index = 0; index < before.Nodes.Count; index++)
        {
            var oldNode = before.Nodes[index];
            var newNode = after.Nodes[index];
            var targetNode = target.DiagramText.Nodes[index];
            var nodePath = $"{path}.nodes[{index}]";
            if (!oldNode.Id.Equals(newNode.Id, StringComparison.Ordinal))
                throw Unsupported(nodePath + ".id", "source-bound SmartArt node identity change");
            RequireNativeRef(oldNode.Raw, newNode.Raw, nodePath);
            RequireEqualExcept(oldNode.Raw, newNode.Raw, nodePath, "text");

            var oldRuns = SmartArtTextRuns(oldNode.Text, nodePath + ".text");
            var newRuns = SmartArtTextRuns(newNode.Text, nodePath + ".text");
            var sourceRuns = targetNode.RunTexts.Count > 0 ? targetNode.RunTexts.ToArray() : [targetNode.Text];
            if (!oldRuns.SequenceEqual(sourceRuns, StringComparer.Ordinal))
                throw new CodecException(
                    "ppj.nativeRef.stale",
                    "The projected SmartArt node text no longer matches the exact source binding.",
                    nodePath + ".text");
            if (oldRuns.SequenceEqual(newRuns, StringComparer.Ordinal)) continue;
            if (oldRuns.Count != newRuns.Count)
                throw Unsupported(nodePath + ".text", "source-bound SmartArt run topology change");
            RequireCapabilityField(after.NativeRef, "setSmartArtText", "smartArt.text", nodePath + ".text");
            RequireCapabilityField(newNode.NativeRef, "setSmartArtText", "smartArt.text", nodePath + ".text");

            targetNode.Text = string.Concat(newRuns);
            targetNode.RunTexts.Clear();
            targetNode.RunTexts.Add(newRuns);
            changed = true;
        }
        return changed;
    }

    private static bool ApplyNativeSmartArtElement(
        PpjSmartArtElementModel before,
        PpjSmartArtElementModel after,
        PresentationDiagram target,
        IReadOnlyDictionary<string, string> assets,
        JsonElement grammarRoot,
        string path)
    {
        RequireEqualExcept(
            before.Raw,
            after.Raw,
            path,
            "role",
            "tags",
            "hidden",
            "locked",
            "frame",
            "nodes",
            "connections",
            "detachToShapes",
            "accessibility");
        if (before.Mode != "source-bound" || after.Mode != "source-bound" ||
            before.Nodes.Count != after.Nodes.Count || before.Nodes.Count != target.Nodes.Count ||
            before.Connections.Count != after.Connections.Count || before.Connections.Count != target.Connections.Count)
            throw Unsupported(path, "native SmartArt identity, mode, or topology change");
        for (var index = 0; index < before.Nodes.Count; index++)
        {
            var oldNode = before.Nodes[index];
            var newNode = after.Nodes[index];
            var targetNode = target.Nodes[index];
            var nodePath = $"{path}.nodes[{index}]";
            if (oldNode.Id != targetNode.Id || oldNode.Id != newNode.Id ||
                string.Concat(SmartArtTextRuns(oldNode.Text, nodePath + ".text")) != PptxTextCodec.Flatten(targetNode.TextBody))
                throw new CodecException("ppj.nativeRef.stale", "The projected SmartArt node no longer matches the exact source binding.", $"{path}.nodes[{index}]");
            RequireEqualExcept(oldNode.Raw, newNode.Raw, nodePath, "text", "asset", "image");
            var oldRuns = SmartArtTextRuns(oldNode.Text, nodePath + ".text");
            var newRuns = SmartArtTextRuns(newNode.Text, nodePath + ".text");
            var cached = target.Drawing?.Children
                .SingleOrDefault(child => child.Name == targetNode.Id && child.Shape is not null);
            var assetChanged = !string.Equals(oldNode.AssetId, newNode.AssetId, StringComparison.Ordinal);
            if (assetChanged)
            {
                if (string.IsNullOrWhiteSpace(oldNode.AssetId) || string.IsNullOrWhiteSpace(newNode.AssetId))
                    throw Unsupported(nodePath + ".asset", "source-bound SmartArt picture assets cannot be added or removed");
                if (!assets.TryGetValue(oldNode.AssetId, out var oldNativeAssetId) ||
                    !string.Equals(oldNativeAssetId, targetNode.AssetId, StringComparison.Ordinal))
                    throw new CodecException(
                        "ppj.nativeRef.stale",
                        "The projected SmartArt node picture no longer matches the exact source asset binding.",
                        nodePath + ".asset");
                if (!assets.TryGetValue(newNode.AssetId, out var newNativeAssetId))
                    throw new CodecException(
                        "ppj.asset.missing",
                        $"SmartArt node picture asset {newNode.AssetId} is not declared in the source-bound asset catalog.",
                        nodePath + ".asset");
                RequireCapabilityField(after.NativeRef, "setSmartArtImage", "smartArt.nodes[].asset", nodePath + ".asset");
                if (cached?.Shape?.ImageFill is not { } cachedPaint)
                    throw Unsupported(nodePath + ".asset", "source-bound SmartArt picture replacement requires a recognized cached image paint");
                targetNode.AssetId = newNativeAssetId;
                cachedPaint.AssetId = newNativeAssetId;
            }
            var imageChanged = !JsonEqual(oldNode.Image, newNode.Image);
            if (imageChanged)
            {
                if (oldNode.Image is not { } oldImage || newNode.Image is not { } newImage)
                    throw Unsupported(nodePath + ".image", "source-bound SmartArt image paint cannot be added or removed");
                if (cached?.Shape?.ImageFill is not { } cachedPaint)
                    throw Unsupported(nodePath + ".image", "source-bound SmartArt image paint requires a recognized cached image paint");
                var expectedOldPaint = BuildSmartArtNodeImagePaint(
                    oldImage,
                    oldNode.AssetId ?? string.Empty,
                    assets,
                    grammarRoot,
                    nodePath + ".image");
                if (!SmartArtImagePaintVisualEqual(expectedOldPaint, cachedPaint))
                    throw new CodecException(
                        "ppj.nativeRef.stale",
                        "The projected SmartArt node image paint no longer matches the exact source binding.",
                        nodePath + ".image");
                var requestedPaint = BuildSmartArtNodeImagePaint(
                    newImage,
                    newNode.AssetId ?? string.Empty,
                    assets,
                    grammarRoot,
                    nodePath + ".image");
                cachedPaint.Crop = requestedPaint.Crop;
                cachedPaint.Mode = requestedPaint.Mode;
                if (requestedPaint.HasOpacityThousandthPercent)
                    cachedPaint.OpacityThousandthPercent = requestedPaint.OpacityThousandthPercent;
                else
                    cachedPaint.ClearOpacityThousandthPercent();
            }
            if (oldRuns.SequenceEqual(newRuns, StringComparer.Ordinal)) continue;
            if (oldRuns.Count != newRuns.Count)
                throw Unsupported(nodePath + ".text", "native SmartArt run topology change");
            RequireCapabilityField(after.NativeRef, "setSmartArtText", "smartArt.text", nodePath + ".text");
            ReplaceDiagramTextRuns(targetNode.TextBody, newRuns, nodePath + ".text");
            if (cached is null)
                throw Unsupported(nodePath + ".text", "native SmartArt text edit without a matching cached drawing node");
            cached.Shape.TextBody = targetNode.TextBody.Clone();
            cached.Shape.Text = PptxTextCodec.Flatten(targetNode.TextBody);
        }
        var graphChanged = false;
        for (var index = 0; index < before.Connections.Count; index++)
        {
            var projected = before.Connections[index];
            var requested = after.Connections[index];
            var source = target.Connections[index];
            if (projected.Id != source.Id || projected.FromId != source.FromId || projected.ToId != source.ToId ||
                projected.Role != source.Role || projected.Order != source.Order)
                throw new CodecException("ppj.nativeRef.stale", "The projected SmartArt connection no longer matches the exact source binding.", $"{path}.connections[{index}]");
            if (projected.Id != requested.Id)
                throw Unsupported($"{path}.connections[{index}].id", "native SmartArt connection identity change");
            if (projected.FromId == requested.FromId && projected.ToId == requested.ToId &&
                projected.Role == requested.Role && projected.Order == requested.Order) continue;
            graphChanged = true;
            source.FromId = requested.FromId;
            source.ToId = requested.ToId;
            source.Role = requested.Role;
            source.Order = requested.Order;
        }
        if (graphChanged)
            RequireCapabilityField(after.NativeRef, "setSmartArtGraph", "smartArt.connections", path + ".connections");
        var frameChanged = FrameChanged(before, after);
        if (frameChanged)
        {
            RequireCapability(after, "setFrame", path + ".frame");
            ApplyDiagramFrame(before, after, target, path + ".frame");
        }
        if (after.DetachToShapes != before.DetachToShapes)
        {
            if (!after.DetachToShapes)
                throw Unsupported(path + ".detachToShapes", "reattaching ordinary shapes as SmartArt");
            if (!target.DrawingCacheVerified || target.Drawing is null || target.Drawing.Children.Count == 0)
                throw Unsupported(path + ".detachToShapes", "detaching SmartArt without a verified cached drawing");
            RequireCapabilityField(
                after.NativeRef,
                "detachSmartArt",
                "smartArt.detachToShapes",
                path + ".detachToShapes");
            return true;
        }
        return graphChanged || frameChanged || before.Nodes.Zip(after.Nodes).Any(pair =>
            !string.Equals(pair.First.AssetId, pair.Second.AssetId, StringComparison.Ordinal) ||
            !JsonEqual(pair.First.Image, pair.Second.Image) ||
            !SmartArtTextRuns(pair.First.Text, path + ".nodes.text")
                .SequenceEqual(SmartArtTextRuns(pair.Second.Text, path + ".nodes.text"), StringComparer.Ordinal));
    }

    private static PresentationImagePaint BuildSmartArtNodeImagePaint(
        JsonElement image,
        string assetId,
        IReadOnlyDictionary<string, string> assets,
        JsonElement grammarRoot,
        string path)
    {
        if (string.IsNullOrWhiteSpace(assetId) || image.ValueKind != JsonValueKind.Object)
            throw Unsupported(path, "SmartArt node image paint requires a bound asset and object value");
        var imageObject = System.Text.Json.Nodes.JsonNode.Parse(image.GetRawText())?.AsObject() ??
            throw Unsupported(path, "SmartArt node image paint must be an object");
        imageObject["type"] = "image";
        imageObject["asset"] = assetId;
        using var document = JsonDocument.Parse(imageObject.ToJsonString());
        var output = BuildImagePaint(
            document.RootElement,
            1,
            1,
            assets,
            new Dictionary<string, (double Width, double Height)>(StringComparer.Ordinal),
            path,
            resolveOpacity: value => ResolveGrammarOpacityToken(grammarRoot, value, path + ".opacity"),
            resolveFit: value => ResolveGrammarStringToken(grammarRoot, value, path + ".fit"));
        if (output.Mode is not (PresentationImagePaint.Types.Mode.Stretch or PresentationImagePaint.Types.Mode.Tile))
            throw Unsupported(path, "SmartArt node image paint fit is outside the stretch/tile profile");
        return output;
    }

    private static bool SmartArtImagePaintVisualEqual(
        PresentationImagePaint left,
        PresentationImagePaint right)
    {
        if (left.Mode != right.Mode || left.HasOpacityThousandthPercent != right.HasOpacityThousandthPercent ||
            left.HasOpacityThousandthPercent && left.OpacityThousandthPercent != right.OpacityThousandthPercent)
            return false;
        if (left.Crop is null || right.Crop is null) return left.Crop is null && right.Crop is null;
        return left.Crop.LeftThousandthPercent == right.Crop.LeftThousandthPercent &&
            left.Crop.TopThousandthPercent == right.Crop.TopThousandthPercent &&
            left.Crop.RightThousandthPercent == right.Crop.RightThousandthPercent &&
            left.Crop.BottomThousandthPercent == right.Crop.BottomThousandthPercent;
    }

    private static void ReplaceDiagramTextRuns(
        PresentationTextBody? body,
        IReadOnlyList<string> values,
        string path)
    {
        if (body is null) throw Unsupported(path, "native SmartArt text without a text body");
        var runs = body.Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Text)
            .ToArray();
        if (runs.Length != values.Count)
            throw Unsupported(path, "native SmartArt text run topology mismatch");
        for (var index = 0; index < runs.Length; index++) runs[index].Text = values[index];
    }

    private static void ApplyDiagramFrame(
        PpjSmartArtElementModel before,
        PpjSmartArtElementModel after,
        PresentationDiagram target,
        string path)
    {
        if (target.LeftEmu != Emu(before.Frame.X) || target.TopEmu != Emu(before.Frame.Y) ||
            target.WidthEmu != Emu(before.Frame.Width) || target.HeightEmu != Emu(before.Frame.Height))
            throw new CodecException("ppj.nativeRef.stale", "The projected SmartArt frame no longer matches the exact source binding.", path);
        var newLeft = Emu(after.Frame.X);
        var newTop = Emu(after.Frame.Y);
        var newWidth = Emu(after.Frame.Width);
        var newHeight = Emu(after.Frame.Height);
        if (newWidth <= 0 || newHeight <= 0) throw Unsupported(path, "non-positive native SmartArt frame");
        var scaleX = newWidth / (double)target.WidthEmu;
        var scaleY = newHeight / (double)target.HeightEmu;
        foreach (var child in target.Drawing.Children)
        {
            if (child.ContentCase != PresentationElement.ContentOneofCase.Shape)
                throw Unsupported(path, "frame scaling with an unsupported cached drawing child");
            var shape = child.Shape;
            shape.LeftEmu = newLeft + checked((long)Math.Round((shape.LeftEmu - target.LeftEmu) * scaleX));
            shape.TopEmu = newTop + checked((long)Math.Round((shape.TopEmu - target.TopEmu) * scaleY));
            shape.WidthEmu = Math.Max(1, checked((long)Math.Round(shape.WidthEmu * scaleX)));
            shape.HeightEmu = Math.Max(1, checked((long)Math.Round(shape.HeightEmu * scaleY)));
        }
        target.LeftEmu = newLeft;
        target.TopEmu = newTop;
        target.WidthEmu = newWidth;
        target.HeightEmu = newHeight;
        target.Drawing.LeftEmu = newLeft;
        target.Drawing.TopEmu = newTop;
        target.Drawing.WidthEmu = newWidth;
        target.Drawing.HeightEmu = newHeight;
        target.Drawing.ChildLeftEmu = newLeft;
        target.Drawing.ChildTopEmu = newTop;
        target.Drawing.ChildWidthEmu = newWidth;
        target.Drawing.ChildHeightEmu = newHeight;
    }

    private static IReadOnlyList<string> SmartArtTextRuns(PpjTextContentModel text, string path)
    {
        if (text.PlainText is not null) return [text.PlainText];
        if (text.Paragraphs.Count != 1 || text.Paragraphs[0].Runs.Count == 0)
            throw Unsupported(path, "source-bound SmartArt paragraph or run topology change");
        if (text.Paragraphs[0].Runs.Any(run => run.Text is null))
            throw Unsupported(path, "source-bound SmartArt formula mutation");
        return text.Paragraphs[0].Runs.Select(run => run.Text!).ToArray();
    }

    private static bool ApplySourceOleElement(
        PpjOleElementModel before,
        PpjOleElementModel after,
        PresentationOpaqueElement target,
        IReadOnlyDictionary<string, string> assets,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "hidden", "locked", "frame", "payloadAsset");
        if (string.IsNullOrWhiteSpace(before.PayloadAssetId) || string.IsNullOrWhiteSpace(after.PayloadAssetId))
            throw Unsupported(path + ".payloadAsset", "removing or inventing a source-bound OLE payload binding");

        var changed = ApplyFrame(before, after, target, path);
        if (before.PayloadAssetId.Equals(after.PayloadAssetId, StringComparison.Ordinal)) return changed;

        RequireCapabilityField(after.NativeRef, "setOlePayload", "ole.payload", path + ".payloadAsset");
        var replacementAssetId = ResolveAsset(assets, after.PayloadAssetId, path + ".payloadAsset");
        if (target.OleWorkbook is not null)
            target.OleWorkbook.ReplacementAssetId = replacementAssetId;
        else if (target.OleOfficePackage is not null)
            target.OleOfficePackage.ReplacementAssetId = replacementAssetId;
        else
            throw Unsupported(path + ".payloadAsset", "the exact source OLE payload binding no longer exists");
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationShape target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        // A typed source-bound placeholder is emitted only after either its
        // direct frame or a unique layout/master effective frame was proved.
        // Both cases can safely materialize an owner-local rotation/flip.
        var allowTransform = true;
        RequireFrameChange(before, after, path, allowTransform);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        if (target.Placeholder is null)
            target.Transform = ShapeTransform(after.Frame);
        else
        {
            // A source-bound slide placeholder may own a direct frame while
            // retaining inherited text/style semantics. Update only the
            // owner-local frame fields represented by the PPJ projection,
            // including optional rotation/flip presence. When the source
            // slide inherited its geometry, this is the one explicit
            // materialization transition permitted by the issued setFrame
            // capability. The effective frame already came from the linked
            // layout/master projection, so no guessed geometry is introduced.
            if (target.DirectFrame is null)
            {
                target.DirectFrame = new PresentationPlaceholderFrame
                {
                    LeftEmu = target.LeftEmu,
                    TopEmu = target.TopEmu,
                    WidthEmu = target.WidthEmu,
                    HeightEmu = target.HeightEmu,
                };
                ApplyRequestedPlaceholderTransform(
                    target.DirectFrame,
                    before.Raw.GetProperty("frame"),
                    after.Raw.GetProperty("frame"),
                    after.Frame,
                    path + ".frame");
                target.Placeholder.InheritsGeometry = false;
            }
            else
            {
                target.DirectFrame.LeftEmu = target.LeftEmu;
                target.DirectFrame.TopEmu = target.TopEmu;
                target.DirectFrame.WidthEmu = target.WidthEmu;
                target.DirectFrame.HeightEmu = target.HeightEmu;
                ApplyRequestedPlaceholderTransform(
                    target.DirectFrame,
                    before.Raw.GetProperty("frame"),
                    after.Raw.GetProperty("frame"),
                    after.Frame,
                    path + ".frame");
            }
        }
        return true;
    }

    private static void ApplyRequestedPlaceholderTransform(
        PresentationPlaceholderFrame target,
        JsonElement beforeFrame,
        JsonElement afterFrame,
        PpjFrameModel requested,
        string path)
    {
        // Keep optional transform presence explicit.  A missing requested
        // property clears the corresponding native attribute; an explicitly
        // supplied zero/false value remains present and is written as such.
        if (afterFrame.TryGetProperty("rotation", out _))
            target.RotationAngle60000 = RotationAngle(requested.Rotation);
        else if (beforeFrame.TryGetProperty("rotation", out _))
            target.ClearRotationAngle60000();
        if (afterFrame.TryGetProperty("flipH", out _))
            target.FlipHorizontal = requested.FlipH;
        else if (beforeFrame.TryGetProperty("flipH", out _))
            target.ClearFlipHorizontal();
        if (afterFrame.TryGetProperty("flipV", out _))
            target.FlipVertical = requested.FlipV;
        else if (beforeFrame.TryGetProperty("flipV", out _))
            target.ClearFlipVertical();
    }

    private static bool TryCollectFrameLeafMutations(
        PpjElementModel before,
        PpjElementModel after,
        PresentationElement target,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> bindings,
        MutationState mutations,
        string path)
    {
        if (!FrameChanged(before, after) || !OnlyFramePropertyChanged(before.Raw, after.Raw) ||
            target.ContentCase is not (PresentationElement.ContentOneofCase.Shape or PresentationElement.ContentOneofCase.Image or PresentationElement.ContentOneofCase.Group) ||
            target.Source?.Editable != true)
            return false;

        var pending = new List<NativeLeafMutation>(7);
        var bindingByKind = bindings.Values
            .Where(binding => binding.ElementId.Equals(after.Id, StringComparison.Ordinal) &&
                              binding.Kind is "leftEmu" or "topEmu" or "widthEmu" or "heightEmu" or "rotationDegrees" or "flipHorizontal" or "flipVertical")
            .GroupBy(binding => binding.Kind, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);

        bool Add(string kind, string value)
        {
            if (!bindingByKind.TryGetValue(kind, out var binding)) return false;
            var next = value;
            if (binding.ExpectedValue.Equals(next, StringComparison.Ordinal)) return true;
            pending.Add(new NativeLeafMutation(
                after.Id,
                slide,
                target,
                shapeTreePath,
                binding.NativeLeafIndex,
                binding.TextLeafIndex,
                binding.ExpectedValue,
                next,
                kind,
                FrameFastPath: true));
            return true;
        }

        var changedCoordinateCount = 0;
        if (!before.Frame.X.Equals(after.Frame.X))
        {
            changedCoordinateCount++;
            if (!Add("leftEmu", Emu(after.Frame.X).ToString(CultureInfo.InvariantCulture))) return false;
        }
        if (!before.Frame.Y.Equals(after.Frame.Y))
        {
            changedCoordinateCount++;
            if (!Add("topEmu", Emu(after.Frame.Y).ToString(CultureInfo.InvariantCulture))) return false;
        }
        if (!before.Frame.Width.Equals(after.Frame.Width))
        {
            changedCoordinateCount++;
            if (!Add("widthEmu", Emu(after.Frame.Width).ToString(CultureInfo.InvariantCulture))) return false;
        }
        if (!before.Frame.Height.Equals(after.Frame.Height))
        {
            changedCoordinateCount++;
            if (!Add("heightEmu", Emu(after.Frame.Height).ToString(CultureInfo.InvariantCulture))) return false;
        }

        if (!before.Frame.Rotation.Equals(after.Frame.Rotation) &&
            !Add("rotationDegrees", RotationAngle(after.Frame.Rotation).ToString(CultureInfo.InvariantCulture))) return false;
        if (before.Frame.FlipH != after.Frame.FlipH &&
            !Add("flipHorizontal", after.Frame.FlipH ? "1" : "0")) return false;
        if (before.Frame.FlipV != after.Frame.FlipV &&
            !Add("flipVertical", after.Frame.FlipV ? "1" : "0")) return false;

        if (changedCoordinateCount == 0 &&
            before.Frame.Rotation.Equals(after.Frame.Rotation) &&
            before.Frame.FlipH == after.Frame.FlipH &&
            before.Frame.FlipV == after.Frame.FlipV)
            return false;
        if (pending.Count == 0) return false;

        mutations.NativeLeaves.AddRange(pending);
        return true;
    }

    private static bool TryCollectGroupChildFrameLeafMutations(
        PpjElementModel before,
        PpjElementModel after,
        PresentationElement target,
        PresentationSlide slide,
        IReadOnlyList<uint> shapeTreePath,
        IReadOnlyDictionary<string, PpjNativeLeafBinding> bindings,
        MutationState mutations,
        string path)
    {
        if (before is not PpjGroupElementModel beforeGroup ||
            after is not PpjGroupElementModel afterGroup ||
            target.ContentCase != PresentationElement.ContentOneofCase.Group ||
            target.Source?.Editable != true ||
            !ChildFrameChanged(beforeGroup, afterGroup) ||
            !OnlyGroupChildFramePropertyChanged(before.Raw, after.Raw))
            return false;

        var pending = new List<NativeLeafMutation>(4);
        var bindingByKind = bindings.Values
            .Where(binding => binding.ElementId.Equals(after.Id, StringComparison.Ordinal) &&
                              binding.Kind is "childLeftEmu" or "childTopEmu" or "childWidthEmu" or "childHeightEmu")
            .GroupBy(binding => binding.Kind, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);

        bool Add(string kind, string value)
        {
            if (!bindingByKind.TryGetValue(kind, out var binding)) return false;
            if (binding.ExpectedValue.Equals(value, StringComparison.Ordinal)) return true;
            pending.Add(new NativeLeafMutation(
                after.Id,
                slide,
                target,
                shapeTreePath,
                binding.NativeLeafIndex,
                binding.TextLeafIndex,
                binding.ExpectedValue,
                value,
                kind,
                FrameFastPath: true));
            return true;
        }

        if (!beforeGroup.ChildFrame.X.Equals(afterGroup.ChildFrame.X) &&
            !Add("childLeftEmu", Emu(afterGroup.ChildFrame.X).ToString(CultureInfo.InvariantCulture))) return false;
        if (!beforeGroup.ChildFrame.Y.Equals(afterGroup.ChildFrame.Y) &&
            !Add("childTopEmu", Emu(afterGroup.ChildFrame.Y).ToString(CultureInfo.InvariantCulture))) return false;
        if (!beforeGroup.ChildFrame.Width.Equals(afterGroup.ChildFrame.Width) &&
            !Add("childWidthEmu", Emu(afterGroup.ChildFrame.Width).ToString(CultureInfo.InvariantCulture))) return false;
        if (!beforeGroup.ChildFrame.Height.Equals(afterGroup.ChildFrame.Height) &&
            !Add("childHeightEmu", Emu(afterGroup.ChildFrame.Height).ToString(CultureInfo.InvariantCulture))) return false;

        if (pending.Count == 0) return false;
        mutations.NativeLeaves.AddRange(pending);
        return true;
    }

    private static bool ChildFrameChanged(PpjGroupElementModel before, PpjGroupElementModel after) =>
        !before.ChildFrame.Equals(after.ChildFrame);

    private static bool OnlyGroupChildFramePropertyChanged(JsonElement before, JsonElement after)
    {
        var frameChanged = false;
        var names = before.EnumerateObject().Select(property => property.Name)
            .Concat(after.EnumerateObject().Select(property => property.Name))
            .Distinct(StringComparer.Ordinal);
        foreach (var name in names)
        {
            var hasBefore = before.TryGetProperty(name, out var beforeValue);
            var hasAfter = after.TryGetProperty(name, out var afterValue);
            if (!hasBefore || !hasAfter) return false;
            if (name.Equals("childFrame", StringComparison.Ordinal))
            {
                frameChanged = !JsonEqual(beforeValue, afterValue);
                continue;
            }
            if (!JsonEqual(beforeValue, afterValue)) return false;
        }
        return frameChanged;
    }

    private static bool OnlyFramePropertyChanged(JsonElement before, JsonElement after)
    {
        var frameChanged = false;
        var names = before.EnumerateObject().Select(property => property.Name)
            .Concat(after.EnumerateObject().Select(property => property.Name))
            .Distinct(StringComparer.Ordinal);
        foreach (var name in names)
        {
            var hasBefore = before.TryGetProperty(name, out var beforeValue);
            var hasAfter = after.TryGetProperty(name, out var afterValue);
            if (!hasBefore || !hasAfter) return false;
            if (name.Equals("frame", StringComparison.Ordinal))
            {
                frameChanged = !JsonEqual(beforeValue, afterValue);
                continue;
            }
            if (!JsonEqual(beforeValue, afterValue)) return false;
        }
        return frameChanged;
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
        PpjNativeRefModel? nativeRef,
        PresentationSlide slide,
        PresentationElement element,
        IReadOnlyList<uint> shapeTreePath,
        MutationState mutations,
        string path,
        JsonElement programRoot)
    {
        if (JsonEqual(beforeRaw, afterRaw)) return false;
        var tabStopsChanged = !JsonEqual(
            MaskTextValues(beforeRaw, maskAlignment: true),
            MaskTextValues(afterRaw, maskAlignment: true));
        var alignmentChanged = !JsonEqual(
            MaskTextValues(beforeRaw, maskTabStops: true),
            MaskTextValues(afterRaw, maskTabStops: true));
        if (!JsonEqual(
                MaskTextValues(beforeRaw, maskTabStops: true, maskAlignment: true),
                MaskTextValues(afterRaw, maskTabStops: true, maskAlignment: true)))
            throw Unsupported(path, "rich-text topology or styling change");
        if (target.TextBody is null)
            throw Unsupported(path, "text edit without one imported bounded text body");
        if (tabStopsChanged || alignmentChanged)
        {
            if (tabStopsChanged)
                RequireCapabilityField(
                    nativeRef,
                    "setTextParagraphStyle",
                    "text.paragraphs[].style.tabStops",
                    path + ".paragraphStyle.tabStops");
            if (alignmentChanged)
                RequireCapabilityField(
                    nativeRef,
                    "setTextParagraphStyle",
                    "text.paragraphs[].style.alignment",
                    path + ".paragraphStyle.alignment");
            ApplyTextParagraphStyleMutation(afterRaw, target, programRoot, path);
            mutations.SemanticChanges = true;
        }

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
                    var beforeRun = before.Paragraphs[paragraph].Runs[run];
                    var afterRun = after.Paragraphs[paragraph].Runs[run];
                    if (targetRun.ContentCase == PresentationTextRun.ContentOneofCase.Text)
                    {
                        var oldText = beforeRun.Text;
                        var newText = afterRun.Text;
                        if (oldText is null || newText is null)
                            throw Unsupported(path, "source-bound formula mutation");
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
                    }
                    else if (targetRun.ContentCase == PresentationTextRun.ContentOneofCase.Field)
                    {
                        // Keep field identity and type source-owned. A static
                        // field display value is safe to update through the
                        // existing text-body export path without changing
                        // paragraph/run topology or host field semantics.
                        if (beforeRun.Field is null || afterRun.Field is null ||
                            !string.Equals(beforeRun.Field.Id, afterRun.Field.Id, StringComparison.Ordinal) ||
                            !string.Equals(beforeRun.Field.Type, afterRun.Field.Type, StringComparison.Ordinal) ||
                            beforeRun.Field.Automatic != afterRun.Field.Automatic ||
                            !string.Equals(targetRun.Field.Id, beforeRun.Field.Id, StringComparison.Ordinal) ||
                            !string.Equals(targetRun.Field.Type, beforeRun.Field.Type, StringComparison.Ordinal))
                            throw Unsupported(path, "source-bound field identity or type change");
                        if (!string.Equals(beforeRun.Field.Text, afterRun.Field.Text, StringComparison.Ordinal))
                        {
                            targetRun.Field.Automatic = afterRun.Field.Automatic;
                            targetRun.Field.Text = afterRun.Field.Text;
                            target.Text = PptxTextCodec.Flatten(target.TextBody);
                            mutations.SemanticChanges = true;
                        }
                    }
                    else if (targetRun.ContentCase == PresentationTextRun.ContentOneofCase.LineBreak)
                    {
                        if (!beforeRun.LineBreak || !afterRun.LineBreak)
                            throw Unsupported(path, "source-bound line-break identity change");
                    }
                    else
                        throw Unsupported(path, "non-text imported run mutation");
                    leafIndex++;
                }
            }
        }
        return true;
    }

    private static void ApplySpeakerNotes(
        PpjPageModel before,
        PpjPageModel after,
        PresentationSlide slide,
        string path)
    {
        if (after.Notes is null)
            throw Unsupported(path, "deleting an imported NotesSlide");

        var source = slide.SpeakerNotes;
        if (source is null)
        {
            if (before.Notes is not null)
                throw Unsupported(path, "notes state inconsistent with the fresh source projection");
            if (after.Notes.PlainText is null)
                throw Unsupported(path, "adding structured notes; the bounded add profile accepts plain text only");
            slide.SpeakerNotes = new PresentationSpeakerNotes { Text = after.Notes.PlainText };
            return;
        }

        var requested = source.Clone();
        if (before.Notes is null)
        {
            if (!string.IsNullOrEmpty(source.Text) || after.Notes.PlainText is null)
                throw Unsupported(path, "adding structured notes or changing an unprojected notes body");
            requested.Text = after.Notes.PlainText;
            requested.TextBody = null;
            slide.SpeakerNotes = requested;
            return;
        }

        var beforeRaw = before.Raw.GetProperty("notes");
        var afterRaw = after.Raw.GetProperty("notes");
        if (!JsonEqual(MaskTextValues(beforeRaw), MaskTextValues(afterRaw)))
            throw Unsupported(path, "rich-text topology or styling change");

        if (before.Notes.PlainText is not null || after.Notes.PlainText is not null)
        {
            if (before.Notes.PlainText is null || after.Notes.PlainText is null || source.TextBody is not null)
                throw Unsupported(path, "plain/rich notes conversion");
            requested.Text = after.Notes.PlainText;
            requested.TextBody = null;
            slide.SpeakerNotes = requested;
            return;
        }

        if (source.TextBody is null ||
            before.Notes.Paragraphs.Count != after.Notes.Paragraphs.Count ||
            before.Notes.Paragraphs.Count != source.TextBody.Paragraphs.Count)
            throw Unsupported(path, "paragraph topology change");

        var body = source.TextBody.Clone();
        for (var paragraphIndex = 0; paragraphIndex < before.Notes.Paragraphs.Count; paragraphIndex++)
        {
            var oldParagraph = before.Notes.Paragraphs[paragraphIndex];
            var newParagraph = after.Notes.Paragraphs[paragraphIndex];
            var targetParagraph = body.Paragraphs[paragraphIndex];
            if (oldParagraph.Runs.Count != newParagraph.Runs.Count ||
                oldParagraph.Runs.Count != targetParagraph.Runs.Count)
                throw Unsupported(path, "run topology change");
            for (var runIndex = 0; runIndex < oldParagraph.Runs.Count; runIndex++)
            {
                var targetRun = targetParagraph.Runs[runIndex];
                if (targetRun.ContentCase != PresentationTextRun.ContentOneofCase.Text)
                    throw Unsupported(path, "non-text imported run mutation");
                targetRun.Text = newParagraph.Runs[runIndex].Text ??
                    throw Unsupported(path, "source-bound formula mutation");
            }
        }
        requested.TextBody = body;
        requested.Text = string.Join("\n", body.Paragraphs.Select(paragraph =>
            string.Concat(paragraph.Runs.Select(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Text ? run.Text : string.Empty))));
        slide.SpeakerNotes = requested;
    }

    private static bool ApplyShapeStyle(
        PpjElementModel before,
        PpjElementModel after,
        PresentationShape target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot,
        string path)
    {
        var oldStyle = OptionalProperty(before.Raw, "style");
        var newStyle = OptionalProperty(after.Raw, "style");
        if (JsonEqual(oldStyle, newStyle)) return false;
        if (oldStyle is { } oldValue && newStyle is { } newValue)
            RequireEqualExcept(oldValue, newValue, path + ".style", "fill", "stroke", "shadow", "glow", "innerShadow", "reflection", "softEdge");
        else
        {
            var present = oldStyle ?? newStyle!.Value;
            foreach (var property in present.EnumerateObject())
                if (property.Name is not ("fill" or "stroke" or "shadow" or "glow" or "innerShadow" or "reflection" or "softEdge"))
                    throw Unsupported(path + ".style", $"changing {property.Name}");
        }
        var changed = false;
        if (PropertyChanged(oldStyle, newStyle, "fill"))
        {
            RequireCapability(after, "setFill", path + ".style.fill");
            ApplyFill(newStyle is { } style && style.TryGetProperty("fill", out var fill) ? fill : (JsonElement?)null,
                target, assets, assetDimensions, grammarRoot, path + ".style.fill");
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".style.stroke");
            ApplyStroke(newStyle is { } style && style.TryGetProperty("stroke", out var stroke) ? stroke : (JsonElement?)null, target, grammarRoot, path + ".style.stroke");
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "shadow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.shadow", path + ".style.shadow");
            target.Shadow = newStyle is { } style && style.TryGetProperty("shadow", out var shadow)
                ? SourceBoundShadow(shadow, path + ".style.shadow", "shape", grammarRoot)
                : null;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "glow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.glow", path + ".style.glow");
            target.Glow = newStyle is { } style && style.TryGetProperty("glow", out var glow)
                ? SourceBoundGlow(glow, path + ".style.glow", "shape", grammarRoot)
                : null;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "innerShadow"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.innerShadow", path + ".style.innerShadow");
            target.InnerShadow = newStyle is { } style && style.TryGetProperty("innerShadow", out var innerShadow)
                ? SourceBoundInnerShadow(innerShadow, path + ".style.innerShadow", "shape", grammarRoot)
                : null;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "reflection"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.reflection", path + ".style.reflection");
            target.Reflection = newStyle is { } style && style.TryGetProperty("reflection", out var reflection)
                ? SourceBoundReflection(reflection, path + ".style.reflection", "shape", grammarRoot)
                : null;
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "softEdge"))
        {
            RequireCapabilityField(after.NativeRef, "setShapeEffects", "shape.softEdge", path + ".style.softEdge");
            target.SoftEdge = newStyle is { } style && style.TryGetProperty("softEdge", out var softEdge)
                ? SourceBoundSoftEdge(softEdge, path + ".style.softEdge", "shape", grammarRoot)
                : null;
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
        JsonElement grammarRoot,
        string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setFill", $"{path}.{name}");
        ApplyFill(after.Raw.TryGetProperty(name, out var fill) ? fill : (JsonElement?)null,
            target, assets, assetDimensions, grammarRoot, $"{path}.{name}");
        return true;
    }

    private static bool ApplyStrokeProperty(
        PpjElementModel before,
        PpjElementModel after,
        PresentationShape target,
        string name,
        JsonElement grammarRoot,
        string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setStroke", $"{path}.{name}");
        ApplyStroke(after.Raw.TryGetProperty(name, out var stroke) ? stroke : (JsonElement?)null, target, grammarRoot, $"{path}.{name}");
        return true;
    }

    private static void ApplyFill(
        JsonElement? fill,
        PresentationShape target,
        IReadOnlyDictionary<string, string> assets,
        IReadOnlyDictionary<string, (double Width, double Height)> assetDimensions,
        JsonElement grammarRoot,
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
            target.GradientFill = BuildGradientFill(
                fill.Value,
                path,
                (color, colorPath) => ResolveGrammarColorValue(grammarRoot, color, colorPath));
            target.ImageFill = null;
            return;
        }
        if (fill.Value.GetProperty("type").GetString() == "image")
        {
            target.FillRgb = string.Empty;
            target.ClearFillOpacityThousandthPercent();
            target.GradientFill = null;
            target.ImageFill = BuildImagePaint(fill.Value, target.WidthEmu / 12_700d, target.HeightEmu / 12_700d,
                assets, assetDimensions, path,
                resolveOpacity: opacity => ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".opacity"),
                resolveFit: fit => ResolveGrammarStringToken(grammarRoot, fit, path + ".fit"));
            return;
        }
        if (fill.Value.GetProperty("type").GetString() != "solid")
            throw Unsupported(path, "unsupported fill");
        target.GradientFill = null;
        target.ImageFill = null;
        var resolvedColor = ResolveGrammarColorValue(grammarRoot, fill.Value.GetProperty("color"), path + ".color");
        target.FillRgb = resolvedColor.Rgb;
        var fillOpacity = resolvedColor.Alpha;
        if (fill.Value.TryGetProperty("opacity", out var opacity))
            fillOpacity = ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".opacity");
        if (fillOpacity < 1)
            target.FillOpacityThousandthPercent = Unit(fillOpacity);
        else target.ClearFillOpacityThousandthPercent();
    }

    private static void ApplyStroke(JsonElement? stroke, PresentationShape target, JsonElement grammarRoot, string path)
    {
        if (stroke is null)
        {
            target.LineRgb = string.Empty;
            target.LineScheme = string.Empty;
            target.LineStyle = "none";
            target.LineWidthEmu = 0;
            return;
        }
        var colorAlpha = ApplyStrokeColor(stroke.Value.GetProperty("color"), target, grammarRoot, path + ".color");
        target.LineWidthEmu = Emu(ResolveGrammarSizeToken(grammarRoot, stroke.Value.GetProperty("width"), path + ".width"));
        target.LineStyle = NativeDash(OptionalString(stroke.Value, "dash"));
        target.LineCap = OptionalString(stroke.Value, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke.Value, "join") ?? string.Empty;
        var strokeOpacity = colorAlpha;
        if (stroke.Value.TryGetProperty("opacity", out var opacity))
            strokeOpacity = ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".opacity");
        if (strokeOpacity < 1)
            target.LineOpacityThousandthPercent = Unit(strokeOpacity);
        else target.ClearLineOpacityThousandthPercent();
    }

    private static void ApplyConnectorStroke(JsonElement stroke, PresentationConnector target, JsonElement grammarRoot, string path)
    {
        var colorAlpha = ApplyStrokeColor(stroke.GetProperty("color"), target, grammarRoot, path + ".color");
        target.LineWidthEmu = Emu(ResolveGrammarSizeToken(grammarRoot, stroke.GetProperty("width"), path + ".width"));
        target.LineStyle = NativeDash(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
        var strokeOpacity = colorAlpha;
        if (stroke.TryGetProperty("opacity", out var opacity))
            strokeOpacity = ResolveGrammarOpacityToken(grammarRoot, opacity, path + ".opacity");
        if (strokeOpacity < 1)
            target.LineOpacityThousandthPercent = Unit(strokeOpacity);
        else target.ClearLineOpacityThousandthPercent();
    }

    private static PresentationGradientFill BuildGradientFill(
        JsonElement fill,
        string path,
        Func<JsonElement, string, (string Rgb, double Alpha)>? resolveColor = null)
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
            var colorValue = item.GetProperty("color");
            (string Rgb, double Alpha) resolvedColor = resolveColor is not null
                ? resolveColor(colorValue, $"{path}.stops[{index}].color")
                : (Rgb(colorValue, $"{path}.stops[{index}].color"), 1d);
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = Unit(item.GetProperty("offset").GetDouble()),
                ColorRgb = resolvedColor.Rgb,
            };
            var stopOpacity = resolvedColor.Alpha;
            var hasExplicitOpacity = false;
            if (item.TryGetProperty("opacity", out var opacity))
            {
                stopOpacity = opacity.GetDouble();
                hasExplicitOpacity = true;
            }
            if (hasExplicitOpacity || stopOpacity < 1)
                stop.OpacityThousandthPercent = Unit(ValidateOpacity(stopOpacity, $"{path}.stops[{index}].opacity"));
            output.Stops.Add(stop);
            index++;
        }
        PptxGradientFillCodec.Validate(output, path);
        return output;
    }

    private static void ApplyChartData(
        PpjChartElementModel before,
        PpjChartElementModel after,
        PresentationChart target,
        JsonElement grammarRoot,
        string path)
    {
        if (before.Data.Categories.Count != after.Data.Categories.Count ||
            before.Data.Series.Count != after.Data.Series.Count)
            throw Unsupported(path, "chart series or category topology change");
        var valueChanged = !before.Data.Categories.Zip(after.Data.Categories).All(pair => JsonEqual(pair.First, pair.Second)) ||
            before.Data.Series.Zip(after.Data.Series).Any(pair =>
                !pair.First.Name.Equals(pair.Second.Name, StringComparison.Ordinal) ||
                !pair.First.Values.SequenceEqual(pair.Second.Values) ||
                !pair.First.CategoryFormula.Equals(pair.Second.CategoryFormula, StringComparison.Ordinal) ||
                !pair.First.XValueFormula.Equals(pair.Second.XValueFormula, StringComparison.Ordinal) ||
                !pair.First.ValueFormula.Equals(pair.Second.ValueFormula, StringComparison.Ordinal) ||
                !pair.First.BubbleSizeFormula.Equals(pair.Second.BubbleSizeFormula, StringComparison.Ordinal));
        var fillChanged = before.Data.Series.Zip(after.Data.Series)
            .Any(pair => PropertyChanged(pair.First.Raw, pair.Second.Raw, "fill") ||
                PropertyChanged(pair.First.Raw, pair.Second.Raw, "pointStyles"));
        var labelsChanged = before.Data.Series.Zip(after.Data.Series)
            .Any(pair => PropertyChanged(pair.First.Raw, pair.Second.Raw, "dataLabels"));
        if (valueChanged) RequireCapability(after, "setChartData", path);
        if (fillChanged) RequireCapability(after, "setChartFill", path + ".series[].fill-or-pointStyles");
        if (labelsChanged) RequireCapability(after, "setChartLabels", path + ".series[].dataLabels");
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
                ApplyChartSeries(
                    before.Data.Series[index],
                    after.Data.Series[index],
                    target.ComboSeries[index].Series,
                    target.ComboSeries[index].Type,
                    after,
                    grammarRoot,
                    $"{path}.series[{index}]");
        }
        else
        {
            if (target.Series.Count != after.Data.Series.Count)
                throw Unsupported(path, "chart series topology change");
            for (var index = 0; index < after.Data.Series.Count; index++)
                ApplyChartSeries(
                    before.Data.Series[index],
                    after.Data.Series[index],
                    target.Series[index],
                    target.Type,
                    after,
                    grammarRoot,
                    $"{path}.series[{index}]");
        }
    }

    private static void ApplyChartSeries(
        PpjChartSeriesModel before,
        PpjChartSeriesModel after,
        SpreadsheetChartSeriesArtifact target,
        SpreadsheetChartType chartType,
        PpjElementModel capabilityOwner,
        JsonElement grammarRoot,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "name", "values", "categoryFormula", "xValueFormula", "valueFormula", "bubbleSizeFormula", "fill", "stroke", "marker", "pointStyles", "dataLabels", "trendlines", "errorBars");
        if (before.Id != after.Id || before.ChartType != after.ChartType || before.Axis != after.Axis || before.Values.Count != after.Values.Count)
            throw Unsupported(path, "chart-series identity or topology change");
        var formulaChanged = !before.CategoryFormula.Equals(after.CategoryFormula, StringComparison.Ordinal) ||
            !before.XValueFormula.Equals(after.XValueFormula, StringComparison.Ordinal) ||
            !before.ValueFormula.Equals(after.ValueFormula, StringComparison.Ordinal) ||
            !before.BubbleSizeFormula.Equals(after.BubbleSizeFormula, StringComparison.Ordinal);
        if (FormulaPresenceChanged(before, after))
            throw Unsupported(path, "source-bound chart formula reference topology change");
        var cacheChanged = !before.Values.SequenceEqual(after.Values) ||
            !before.XValues.SequenceEqual(after.XValues) ||
            !before.BubbleSizes.SequenceEqual(after.BubbleSizes);
        if (cacheChanged && (HasFormula(before) || HasFormula(after)))
            throw Unsupported(path, "formula-backed chart caches require an owned workbook closure");
        if (formulaChanged)
        {
            RequireCapability(capabilityOwner, "setChartData", path + ".formula");
            target.CategoryFormula = after.CategoryFormula;
            target.XValueFormula = after.XValueFormula;
            target.ValueFormula = after.ValueFormula;
            target.BubbleSizeFormula = after.BubbleSizeFormula;
        }
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
        target.XValues.Clear();
        target.XValues.Add(after.XValues);
        target.BubbleSizes.Clear();
        target.BubbleSizes.Add(after.BubbleSizes);
        if (PropertyChanged(before.Raw, after.Raw, "fill"))
        {
            target.Fill = null;
            target.SeriesFill = after.Raw.TryGetProperty("fill", out var fill)
                ? SourceBoundChartFill(fill, path + ".fill", grammarRoot)
                : null;
        }
        if (PropertyChanged(before.Raw, after.Raw, "pointStyles"))
        {
            target.PointStyles.Clear();
            if (after.Raw.TryGetProperty("pointStyles", out var points))
            {
                var index = 0;
                foreach (var point in points.EnumerateArray())
                {
                    target.PointStyles.Add(SourceBoundChartPointStyle(point, $"{path}.pointStyles[{index}]", grammarRoot));
                    index++;
                }
            }
        }
        if (PropertyChanged(before.Raw, after.Raw, "dataLabels"))
        {
            target.DataLabels = after.Raw.TryGetProperty("dataLabels", out var labels)
                ? SourceBoundSeriesDataLabels(labels, path + ".dataLabels", grammarRoot)
                : null;
        }
        if (PropertyChanged(before.Raw, after.Raw, "stroke"))
        {
            RequireCapability(capabilityOwner, "setChartSeriesStyle", path + ".stroke");
            if (chartType == SpreadsheetChartType.Scatter)
                throw Unsupported(path + ".stroke", "a scatter series line; use marker.stroke for the marker border");
            target.Line = after.Raw.TryGetProperty("stroke", out var stroke)
                ? SourceBoundChartLine(stroke, path + ".stroke", grammarRoot)
                : null;
        }
        if (PropertyChanged(before.Raw, after.Raw, "marker"))
        {
            RequireCapability(capabilityOwner, "setChartSeriesStyle", path + ".marker");
            if (chartType is not (SpreadsheetChartType.Line or SpreadsheetChartType.Scatter or SpreadsheetChartType.Radar))
                throw Unsupported(path + ".marker", $"a marker on {chartType.ToString().ToLowerInvariant()} series");
            target.Marker = after.Raw.TryGetProperty("marker", out var marker)
                ? SourceBoundChartMarker(marker, path + ".marker", grammarRoot)
                : null;
        }
        if (PropertyChanged(before.Raw, after.Raw, "trendlines"))
        {
            RequireCapability(capabilityOwner, "setChartSeriesAnalytics", path + ".trendlines");
            var oldTrendlines = before.Raw.TryGetProperty("trendlines", out var oldTrendlineValue)
                ? oldTrendlineValue
                : (JsonElement?)null;
            var newTrendlines = after.Raw.TryGetProperty("trendlines", out var newTrendlineValue)
                ? newTrendlineValue
                : (JsonElement?)null;
            if (oldTrendlines is null || newTrendlines is null ||
                oldTrendlines.Value.ValueKind != JsonValueKind.Array ||
                newTrendlines.Value.ValueKind != JsonValueKind.Array ||
                oldTrendlines.Value.GetArrayLength() != newTrendlines.Value.GetArrayLength())
                throw Unsupported(path + ".trendlines", "source-bound trendline topology change");
            var trendlines = newTrendlines.Value.EnumerateArray()
                .Select((item, index) => SourceBoundChartTrendline(item, $"{path}.trendlines[{index}]", grammarRoot))
                .ToArray();
            target.Trendlines.Clear();
            target.Trendlines.Add(trendlines);
        }
        if (PropertyChanged(before.Raw, after.Raw, "errorBars"))
        {
            RequireCapability(capabilityOwner, "setChartSeriesAnalytics", path + ".errorBars");
            var oldErrorBars = before.Raw.TryGetProperty("errorBars", out var oldErrorBarValue)
                ? oldErrorBarValue
                : (JsonElement?)null;
            var newErrorBars = after.Raw.TryGetProperty("errorBars", out var newErrorBarValue)
                ? newErrorBarValue
                : (JsonElement?)null;
            if (oldErrorBars is null || newErrorBars is null ||
                oldErrorBars.Value.ValueKind != JsonValueKind.Object ||
                newErrorBars.Value.ValueKind != JsonValueKind.Object)
                throw Unsupported(path + ".errorBars", "source-bound error-bar topology change");
            target.ErrorBars = SourceBoundChartErrorBars(newErrorBars.Value, path + ".errorBars", grammarRoot);
        }
    }

    private static SpreadsheetChartTrendlineArtifact SourceBoundChartTrendline(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var type = source.GetProperty("type").GetString() switch
        {
            "exponential" => SpreadsheetChartTrendlineType.Exponential,
            "linear" => SpreadsheetChartTrendlineType.Linear,
            "logarithmic" => SpreadsheetChartTrendlineType.Logarithmic,
            "moving-average" => SpreadsheetChartTrendlineType.MovingAverage,
            "polynomial" => SpreadsheetChartTrendlineType.Polynomial,
            "power" => SpreadsheetChartTrendlineType.Power,
            _ => throw Unsupported(path + ".type", "unsupported chart trendline type"),
        };
        var output = new SpreadsheetChartTrendlineArtifact
        {
            Type = type,
            Name = source.TryGetProperty("name", out var name) ? name.GetString()! : string.Empty,
            DisplayEquation = source.TryGetProperty("displayEquation", out var equation) && equation.GetBoolean(),
            DisplayRSquared = source.TryGetProperty("displayRSquared", out var rSquared) && rSquared.GetBoolean(),
        };
        if (source.TryGetProperty("order", out var order)) output.PolynomialOrder = checked((uint)order.GetInt32());
        if (source.TryGetProperty("period", out var period)) output.Period = checked((uint)period.GetInt32());
        if (source.TryGetProperty("forward", out var forward)) output.Forward = forward.GetDouble();
        if (source.TryGetProperty("backward", out var backward)) output.Backward = backward.GetDouble();
        if (source.TryGetProperty("intercept", out var intercept)) output.Intercept = intercept.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke))
            output.Line = SourceBoundChartLine(stroke, path + ".stroke", grammarRoot);
        return output;
    }

    private static SpreadsheetChartErrorBarsArtifact SourceBoundChartErrorBars(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var output = new SpreadsheetChartErrorBarsArtifact
        {
            Direction = source.TryGetProperty("direction", out var direction)
                ? direction.GetString() switch
                {
                    "x" => SpreadsheetChartErrorBarDirection.X,
                    "y" => SpreadsheetChartErrorBarDirection.Y,
                    _ => throw Unsupported(path + ".direction", "unsupported chart error-bar direction"),
                }
                : SpreadsheetChartErrorBarDirection.Y,
            Type = source.TryGetProperty("type", out var errorBarType)
                ? errorBarType.GetString() switch
                {
                    "both" => SpreadsheetChartErrorBarType.Both,
                    "minus" => SpreadsheetChartErrorBarType.Minus,
                    "plus" => SpreadsheetChartErrorBarType.Plus,
                    _ => throw Unsupported(path + ".type", "unsupported chart error-bar type"),
                }
                : SpreadsheetChartErrorBarType.Both,
            ValueType = source.GetProperty("valueType").GetString() switch
            {
                "fixed-value" => SpreadsheetChartErrorBarValueType.FixedValue,
                "percentage" => SpreadsheetChartErrorBarValueType.Percentage,
                "standard-deviation" => SpreadsheetChartErrorBarValueType.StandardDeviation,
                "standard-error" => SpreadsheetChartErrorBarValueType.StandardError,
                _ => throw Unsupported(path + ".valueType", "unsupported chart error-bar value type"),
            },
            NoEndCap = source.TryGetProperty("noEndCap", out var noEndCap) && noEndCap.GetBoolean(),
        };
        if (source.TryGetProperty("value", out var value)) output.Value = value.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke))
            output.Line = SourceBoundChartLine(stroke, path + ".stroke", grammarRoot);
        return output;
    }

    private static SpreadsheetChartMarkerArtifact SourceBoundChartMarker(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var symbol = source.ValueKind == JsonValueKind.String
            ? source.GetString()
            : source.TryGetProperty("symbol", out var symbolValue) && symbolValue.ValueKind == JsonValueKind.String
                ? symbolValue.GetString()
                : "circle";
        var output = new SpreadsheetChartMarkerArtifact
        {
            Symbol = symbol switch
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
                _ => throw Unsupported(path + ".symbol", "an unsupported chart marker symbol"),
            },
        };
        if (source.ValueKind != JsonValueKind.Object) return output;
        if (source.TryGetProperty("size", out var size))
            output.Size = grammarRoot is { } root
                ? ResolveGrammarIntegerToken(root, size, path + ".size", 2, 72)
                : checked((uint)size.GetInt32());
        if (source.TryGetProperty("fill", out var fill))
        {
            var resolved = grammarRoot is { } root
                ? ResolveGrammarColorValue(root, fill, path + ".fill")
                : ParseSourceBoundColor(fill, path + ".fill");
            output.Fill = new SpreadsheetColor { Rgb = resolved.Rgb };
            if (resolved.Alpha < 1)
                output.FillOpacityThousandthPercent = Unit(resolved.Alpha);
        }
        if (source.TryGetProperty("stroke", out var stroke))
            output.Line = SourceBoundChartLine(stroke, path + ".stroke", grammarRoot);
        return output;
    }

    private static SpreadsheetChartPointStyleArtifact SourceBoundChartPointStyle(JsonElement source, string path, JsonElement? grammarRoot = null)
    {
        var output = new SpreadsheetChartPointStyleArtifact
        {
            Index = checked((uint)source.GetProperty("index").GetInt32()),
        };
        if (source.TryGetProperty("fill", out var fill))
            output.Fill = SourceBoundChartFill(fill, path + ".fill", grammarRoot);
        if (source.TryGetProperty("stroke", out var stroke))
            output.Line = SourceBoundChartLine(stroke, path + ".stroke");
        if (source.TryGetProperty("explosion", out var explosion))
            output.Explosion = checked((uint)explosion.GetInt32());
        return output;
    }

    private static SpreadsheetChartSeriesDataLabelsArtifact SourceBoundSeriesDataLabels(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var output = new SpreadsheetChartSeriesDataLabelsArtifact();
        if (HasChartLabelFields(source)) output.Defaults = SourceBoundChartLabelOverride(source, path, grammarRoot);
        if (source.TryGetProperty("points", out var points))
        {
            var index = 0;
            foreach (var point in points.EnumerateArray())
            {
                output.Points.Add(new SpreadsheetChartPointDataLabelArtifact
                {
                    Index = checked((uint)point.GetProperty("index").GetInt32()),
                    Override = SourceBoundChartLabelOverride(point, $"{path}.points[{index}]", grammarRoot),
                });
                index++;
            }
        }
        return output;
    }

    private static SpreadsheetChartDataLabelOverrideArtifact SourceBoundChartLabelOverride(
        JsonElement source,
        string path,
        JsonElement? grammarRoot = null)
    {
        var output = new SpreadsheetChartDataLabelOverrideArtifact();
        if (source.TryGetProperty("text", out var text))
            output.Text = text.GetString()!;
        if (source.TryGetProperty("showValue", out var showValue))
            output.ShowValue = grammarRoot is { } showValueRoot
                ? ResolveGrammarBooleanToken(showValueRoot, showValue, path + ".showValue")
                : showValue.GetBoolean();
        if (source.TryGetProperty("showCategory", out var showCategory))
            output.ShowCategoryName = grammarRoot is { } showCategoryRoot
                ? ResolveGrammarBooleanToken(showCategoryRoot, showCategory, path + ".showCategory")
                : showCategory.GetBoolean();
        if (source.TryGetProperty("showSeries", out var showSeries))
            output.ShowSeriesName = grammarRoot is { } showSeriesRoot
                ? ResolveGrammarBooleanToken(showSeriesRoot, showSeries, path + ".showSeries")
                : showSeries.GetBoolean();
        if (source.TryGetProperty("showPercent", out var showPercent))
            output.ShowPercent = grammarRoot is { } showPercentRoot
                ? ResolveGrammarBooleanToken(showPercentRoot, showPercent, path + ".showPercent")
                : showPercent.GetBoolean();
        if (source.TryGetProperty("showBubbleSize", out var showBubbleSize))
            output.ShowBubbleSize = grammarRoot is { } showBubbleSizeRoot
                ? ResolveGrammarBooleanToken(showBubbleSizeRoot, showBubbleSize, path + ".showBubbleSize")
                : showBubbleSize.GetBoolean();
        if (source.TryGetProperty("showLeaderLines", out var showLeaderLines))
            output.ShowLeaderLines = grammarRoot is { } showLeaderLinesRoot
                ? ResolveGrammarBooleanToken(showLeaderLinesRoot, showLeaderLines, path + ".showLeaderLines")
                : showLeaderLines.GetBoolean();
        if (source.TryGetProperty("position", out var position))
            output.Position = SourceBoundDataLabelPosition(
                grammarRoot is { } positionRoot
                    ? ResolveGrammarEnumToken(
                        positionRoot,
                        position,
                        path + ".position",
                        "best-fit", "bottom", "center", "inside-base", "inside-end", "left", "outside-end", "right", "top")
                    : position.GetString()!);
        if (source.TryGetProperty("numberFormat", out var numberFormat))
            output.NumberFormatCode = grammarRoot is { } root
                ? ResolveGrammarStringToken(root, numberFormat, path + ".numberFormat")
                : numberFormat.GetString()!;
        if (source.TryGetProperty("textStyle", out var textStyle))
            output.TextStyle = SourceBoundChartTextStyle(textStyle, path + ".textStyle", grammarRoot);
        if (source.TryGetProperty("fill", out var fill))
            output.Fill = SourceBoundChartFill(fill, path + ".fill", grammarRoot);
        if (source.TryGetProperty("line", out var line))
            output.Line = SourceBoundChartLine(line, path + ".line", grammarRoot);
        return output;
    }

    private static bool HasChartLabelFields(JsonElement source) =>
        source.TryGetProperty("text", out _) || source.TryGetProperty("showValue", out _) || source.TryGetProperty("showCategory", out _) ||
        source.TryGetProperty("showSeries", out _) || source.TryGetProperty("showPercent", out _) ||
        source.TryGetProperty("showBubbleSize", out _) || source.TryGetProperty("showLeaderLines", out _) ||
        source.TryGetProperty("position", out _) || source.TryGetProperty("numberFormat", out _) ||
        source.TryGetProperty("textStyle", out _) || source.TryGetProperty("fill", out _) ||
        source.TryGetProperty("line", out _);

    private static SpreadsheetChartDataLabelPosition SourceBoundDataLabelPosition(string value) => value switch
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
        _ => throw Unsupported("chart.dataLabels.position", $"unsupported data-label position {value}"),
    };

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
                binding.Kind,
                binding.ChartData));
            changed = true;
        }
        return changed;
    }

    private static bool HasFormula(PpjChartSeriesModel series) =>
        !string.IsNullOrEmpty(series.CategoryFormula) || !string.IsNullOrEmpty(series.XValueFormula) ||
        !string.IsNullOrEmpty(series.ValueFormula) || !string.IsNullOrEmpty(series.BubbleSizeFormula);

    private static bool FormulaPresenceChanged(PpjChartSeriesModel before, PpjChartSeriesModel after) =>
        !string.IsNullOrEmpty(before.CategoryFormula) != !string.IsNullOrEmpty(after.CategoryFormula) ||
        !string.IsNullOrEmpty(before.XValueFormula) != !string.IsNullOrEmpty(after.XValueFormula) ||
        !string.IsNullOrEmpty(before.ValueFormula) != !string.IsNullOrEmpty(after.ValueFormula) ||
        !string.IsNullOrEmpty(before.BubbleSizeFormula) != !string.IsNullOrEmpty(after.BubbleSizeFormula);

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

    private static void RequireCapabilityField(PpjElementModel element, string operation, string field, string path)
    {
        RequireCapabilityField(element.NativeRef, operation, field, path);
    }

    private static void RequireCapabilityField(PpjNativeRefModel? nativeRef, string operation, string field, string path)
    {
        var reference = nativeRef ?? throw new CodecException("ppj.nativeRef.missing", "Source-bound edits require a nativeRef.", path);
        var capability = reference.Capabilities.FirstOrDefault(item =>
            item.Operation.Equals(operation, StringComparison.Ordinal) &&
            item.Fields.Contains(field, StringComparer.Ordinal));
        if (capability is null || !capability.ExpectedHash.Equals(reference.ObjectHash, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.nativeRef.capabilityMissing",
                $"The exact source object did not issue the {operation}/{field} capability required by this edit.",
                path);
    }

    private static bool AnimationStateChanged(PpjPageModel before, PpjPageModel after)
    {
        if (before.Animations.Count == 0 && after.Animations.Count == 0) return false;
        return !before.Raw.TryGetProperty("animations", out var oldAnimations) ||
            !after.Raw.TryGetProperty("animations", out var newAnimations) ||
            !JsonEqual(oldAnimations, newAnimations);
    }

    private static IReadOnlyDictionary<string, string> SourceAnimationTargetIds(
        IReadOnlyList<PpjElementModel> programElements,
        IEnumerable<PresentationElement> wireElements,
        string path)
    {
        var output = new Dictionary<string, string>(StringComparer.Ordinal);
        AddSourceAnimationTargetIds(programElements, wireElements.ToArray(), output, path);
        return output;
    }

    private static void AddSourceAnimationTargetIds(
        IReadOnlyList<PpjElementModel> programElements,
        IReadOnlyList<PresentationElement> wireElements,
        IDictionary<string, string> output,
        string path)
    {
        if (programElements.Count != wireElements.Count)
            throw Unsupported(path, "the source element tree changed while resolving animation targets");
        for (var index = 0; index < programElements.Count; index++)
        {
            var program = programElements[index];
            var wire = wireElements[index];
            if (!output.TryAdd(program.Id, wire.Id))
                throw Unsupported(path, $"duplicate animation target identity {program.Id}");
            if (program is not PpjGroupElementModel group) continue;
            if (wire.ContentCase != PresentationElement.ContentOneofCase.Group)
                throw Unsupported(path, $"group target {program.Id} no longer has a source group binding");
            AddSourceAnimationTargetIds(group.Elements, wire.Group.Children.ToArray(), output, path + "." + program.Id);
        }
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

    private static void RequireCapability(
        PpjNativeRefModel? nativeRef,
        string operation,
        string capabilityId,
        string field,
        string path)
    {
        var reference = nativeRef ?? throw new CodecException("ppj.nativeRef.missing", "Source-bound edits require a nativeRef.", path);
        var capability = reference.Capabilities.FirstOrDefault(item => item.Id.Equals(capabilityId, StringComparison.Ordinal));
        if (capability is null || !capability.Operation.Equals(operation, StringComparison.Ordinal) ||
            !capability.Fields.Contains(field, StringComparer.Ordinal) ||
            !capability.ExpectedHash.Equals(reference.ObjectHash, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.nativeRef.capabilityMissing",
                $"The exact source object did not issue capability {capabilityId} for {operation}/{field}.",
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

    private static JsonElement MaskTextValues(
        JsonElement value,
        bool maskTabStops = false,
        bool maskAlignment = false)
    {
        if (value.ValueKind == JsonValueKind.String)
        {
            using var plain = JsonDocument.Parse("\"\"");
            return plain.RootElement.Clone();
        }
        var bytes = PpjCanonicalJson.Write(value);
        var node = System.Text.Json.Nodes.JsonNode.Parse(bytes)!.AsObject();
        if (node["paragraphs"] is System.Text.Json.Nodes.JsonArray paragraphs)
        {
            foreach (var paragraph in paragraphs.OfType<System.Text.Json.Nodes.JsonObject>())
            {
                if (paragraph["runs"] is System.Text.Json.Nodes.JsonArray runs)
                    foreach (var run in runs.OfType<System.Text.Json.Nodes.JsonObject>())
                    {
                        run["text"] = string.Empty;
                        if (run["field"] is System.Text.Json.Nodes.JsonObject field)
                            field["text"] = string.Empty;
                    }
                if ((maskTabStops || maskAlignment) && paragraph["style"] is System.Text.Json.Nodes.JsonObject style)
                {
                    if (maskTabStops)
                    {
                        style.Remove("tabStops");
                        style.Remove("noTabStops");
                    }
                    if (maskAlignment) style.Remove("alignment");
                    if (style.Count == 0) paragraph.Remove("style");
                }
            }
        }
        using var document = JsonDocument.Parse(node.ToJsonString());
        return document.RootElement.Clone();
    }

    private static void ApplyTextParagraphStyleMutation(
        JsonElement afterRaw,
        PresentationShape target,
        JsonElement programRoot,
        string path)
    {
        PresentationTextBody requested;
        try
        {
            requested = PpjAuthoredPresentationCompiler.BuildSourceBoundTextBody(afterRaw, programRoot);
        }
        catch (CodecException error)
        {
            throw Unsupported(path, $"paragraph tab stops contain unsupported properties ({error.Message})");
        }
        if (requested.Paragraphs.Count != target.TextBody!.Paragraphs.Count)
            throw Unsupported(path, "paragraph topology change");
        for (var index = 0; index < requested.Paragraphs.Count; index++)
        {
            var next = requested.Paragraphs[index];
            var current = target.TextBody.Paragraphs[index];
            if (next.TabStops.Count == 0 && !next.HasNoTabStops && current.TabStops.Count > 0)
                throw Unsupported(path, "removing source tab stops requires explicit noTabStops: true");
            current.ClearAlignment();
            if (next.HasAlignment) current.Alignment = next.Alignment;
            current.TabStops.Clear();
            current.TabStops.Add(next.TabStops);
            current.ClearNoTabStops();
            if (next.HasNoTabStops && next.NoTabStops) current.NoTabStops = true;
        }
    }

    private static bool TextEqual(PpjTextContentModel left, PpjTextContentModel right)
    {
        if (left.PlainText is not null || right.PlainText is not null) return left.PlainText == right.PlainText;
        if (left.Paragraphs.Count != right.Paragraphs.Count) return false;
        for (var paragraph = 0; paragraph < left.Paragraphs.Count; paragraph++)
        {
            if (left.Paragraphs[paragraph].Runs.Count != right.Paragraphs[paragraph].Runs.Count) return false;
            for (var run = 0; run < left.Paragraphs[paragraph].Runs.Count; run++)
            {
                var leftRun = left.Paragraphs[paragraph].Runs[run];
                var rightRun = right.Paragraphs[paragraph].Runs[run];
                if (leftRun.Text != rightRun.Text || leftRun.Formula?.Syntax != rightRun.Formula?.Syntax ||
                    leftRun.Formula?.Source != rightRun.Formula?.Source ||
                    leftRun.LineBreak != rightRun.LineBreak ||
                    leftRun.Field?.Id != rightRun.Field?.Id || leftRun.Field?.Type != rightRun.Field?.Type ||
                    leftRun.Field?.Text != rightRun.Field?.Text) return false;
            }
        }
        return true;
    }

    private static bool TextBodyStyleChanged(JsonElement left, JsonElement right) =>
        left.ValueKind == JsonValueKind.Object && right.ValueKind == JsonValueKind.Object &&
        !JsonEqual(MaskTextValues(left), MaskTextValues(right));

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
                mutation.NativeLeafIndex, mutation.TextLeafIndex, mutation.Before, mutation.After,
                mutation.ChartData?.TargetPartPath ?? string.Empty,
                mutation.ChartData?.EmbeddedCellReference ?? string.Empty,
                mutation.ChartData?.ChartChannel ?? string.Empty);
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
            if (PpjNativeLeafProjection.IsChartDataLeafKind(mutation.LeafKind))
            {
                var chartData = mutation.ChartData ?? throw new CodecException(
                    "ppj.nativeRef.stale",
                    "Chart data leaf lost its source-bound ChartPart/workbook binding.",
                    mutation.ProgramElementId);
                operation.TargetPartPath = chartData.TargetPartPath;
                operation.ExpectedTargetPartSha256 = chartData.TargetPartSha256;
                operation.RelationshipId = chartData.RelationshipId;
                operation.EmbeddedPackagePartPath = chartData.EmbeddedPackagePartPath;
                operation.ExpectedEmbeddedPackageSha256 = chartData.EmbeddedPackageSha256;
                operation.EmbeddedPackageRelationshipId = chartData.EmbeddedPackageRelationshipId;
                operation.EmbeddedWorksheetPartPath = chartData.EmbeddedWorksheetPartPath;
                operation.ExpectedEmbeddedWorksheetSha256 = chartData.EmbeddedWorksheetSha256;
                operation.EmbeddedCellReference = chartData.EmbeddedCellReference;
                operation.ChartSeriesIndex = chartData.ChartSeriesIndex;
                operation.ChartPointIndex = chartData.ChartPointIndex;
                operation.ChartFormula = chartData.ChartFormula;
            }
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
            output[entry.FullName] = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
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

    private static double ApplyStrokeColor(JsonElement color, PresentationShape target, JsonElement grammarRoot, string path)
    {
        var resolved = ResolveSourceBoundStrokeColor(color, grammarRoot, path);
        target.LineRgb = resolved.Rgb;
        target.LineScheme = resolved.Scheme;
        return resolved.Alpha;
    }

    private static double ApplyStrokeColor(JsonElement color, PresentationConnector target, JsonElement grammarRoot, string path)
    {
        var resolved = ResolveSourceBoundStrokeColor(color, grammarRoot, path);
        target.LineRgb = resolved.Rgb;
        target.LineScheme = resolved.Scheme;
        return resolved.Alpha;
    }

    private static (string Rgb, string Scheme, double Alpha) ResolveSourceBoundStrokeColor(
        JsonElement color,
        JsonElement grammarRoot,
        string path)
    {
        if (color.ValueKind == JsonValueKind.Object && color.TryGetProperty("token", out var token) &&
            token.ValueKind == JsonValueKind.String)
        {
            var tokenName = token.GetString()!;
            if (TryDeclaredGrammarToken(grammarRoot, tokenName, out var definition))
            {
                if (!definition.TryGetProperty("kind", out var kind) || kind.GetString() != "color")
                    throw new CodecException("ppj.grammar.tokenKind", $"PPJ grammar token {tokenName} for {path} must declare kind color.", path);
                var resolvedGrammarColor = ResolveGrammarColorValue(grammarRoot, color, path);
                return (resolvedGrammarColor.Rgb, string.Empty, resolvedGrammarColor.Alpha);
            }
            // A token not declared by the PPJ grammar retains the historical
            // source-bound meaning of a standard DrawingML theme token.
            return (string.Empty, PptxColor.NormalizeScheme(tokenName), 1d);
        }
        var resolvedColor = ResolveGrammarColorValue(grammarRoot, color, path);
        return (resolvedColor.Rgb, string.Empty, resolvedColor.Alpha);
    }

    private static bool TryDeclaredGrammarToken(JsonElement root, string tokenName, out JsonElement definition)
    {
        definition = default;
        if (!root.TryGetProperty("design", out var design) ||
            !design.TryGetProperty("grammar", out var grammar) ||
            !grammar.TryGetProperty("tokens", out var tokens) ||
            tokens.ValueKind != JsonValueKind.Object)
            return false;
        return tokens.TryGetProperty(tokenName, out definition) && definition.ValueKind == JsonValueKind.Object;
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

    private static string ArrowValue(JsonElement owner, string property) =>
        owner.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String &&
        value.GetString() is { } arrow && arrow != "none"
            ? arrow
            : string.Empty;

    private static string? OptionalString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value.GetString() : null;

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Unsupported(string path, string detail) =>
        new("ppj.source.unsupportedMutation", $"Source-bound PPJ cannot safely compile {detail}.", path);
}
