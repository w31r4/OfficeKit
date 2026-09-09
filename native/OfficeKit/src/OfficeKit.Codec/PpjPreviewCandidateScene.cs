using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class PpjPreviewCandidateScene
{
    internal static IReadOnlyList<Diagnostic> Attach(PptxPackageSource candidate,
        PresentationProgramResult receipt, EffectiveCodecLimits limits, PpjPreviewCandidateBindings? provenance = null)
    {
        using var stage = PpjBuildProfiler.Measure("preview.candidate-import");
        // Import native candidate bytes directly. The projector can restore an
        // embedded authored program and therefore is not a visual-state oracle.
        var imported = PptxCodec.Import(candidate, limits, retainImportedAssetData: true,
            verifiedPackageSha256: receipt.OutputSha256, includeNativeBindings: provenance is not null);
        var presentation = imported.Artifact.Presentation ??
            throw new CodecException("invalid_preview_scene", "Candidate has no native presentation scene.");
        var scene = PpjPreviewSceneBuilder.Build(presentation, PresentationPreviewSceneOrigin.CandidateImport,
            receipt.ProgramSha256, receipt.OutputSha256, Bindings(presentation, imported.NativeBindings, provenance), imported.Artifact.Assets, limits);

        // Public projected asset IDs may differ from native import IDs. Resolve
        // payloads by MIME/hash; do not duplicate bytes just to create ID aliases.
        var payloads = receipt.Assets.GroupBy(AssetKey).ToDictionary(group => group.Key, group => group.ToList());
        foreach (var asset in imported.Artifact.Assets)
        {
            var key = AssetKey(asset);
            if (!payloads.TryGetValue(key, out var aliases)) payloads[key] = aliases = [];
            var existing = aliases.FirstOrDefault(item => item.Data.Equals(asset.Data));
            if (existing is not null) continue;
            var metadataOnly = aliases.FirstOrDefault(item => item.Data.IsEmpty);
            if (metadataOnly is not null) metadataOnly.Data = asset.Data;
            else
            {
                var retained = asset.Clone();
                receipt.Assets.Add(retained);
                aliases.Add(retained);
            }
        }
        receipt.PreviewScene = scene;
        return imported.Diagnostics;
    }

    private static (string MimeType, string Sha256) AssetKey(Asset asset) =>
        (asset.ContentType.ToLowerInvariant(), asset.Sha256.ToLowerInvariant());

    private static IEnumerable<PresentationPreviewNodeBinding> Bindings(PresentationArtifact presentation,
        IReadOnlyList<PptxNativeBinding> identities, PpjPreviewCandidateBindings? provenance)
    {
        var unambiguous = identities.GroupBy(identity => (identity.PartPath, identity.NativeId))
            .Where(group => group.Key.NativeId != 0 && group.Count() == 1).Select(group => group.Single())
            .GroupBy(identity => identity.ElementId, StringComparer.Ordinal).Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        foreach (var (slide, index) in presentation.Slides.Select((slide, index) => (slide, index)))
            foreach (var binding in Walk(slide.Elements, slide.Id, $"$.presentation.slides[{index}].elements",
                unambiguous, provenance, provenance?.Page(slide.Source?.PartPath ?? string.Empty)))
                yield return binding;
    }

    private static IEnumerable<PresentationPreviewNodeBinding> Walk(IEnumerable<PresentationElement> elements,
        string pageId, string path, IReadOnlyDictionary<string, PptxNativeBinding> identities,
        PpjPreviewCandidateBindings? provenance, (string Id, string Path)? page)
    {
        foreach (var (element, index) in elements.Select((element, index) => (element, index)))
        {
            var scenePath = $"{path}[{index}]";
            var owner = identities.TryGetValue(element.Id, out var identity) ? provenance?.Owner(identity) : null;
            yield return new PresentationPreviewNodeBinding
            {
                PageId = owner?.PageId ?? page?.Id ?? pageId, NativeId = element.Id,
                SemanticId = owner?.Id ?? string.Empty, SourceId = owner?.SourceId ?? string.Empty,
                ProgramPath = owner?.ProgramPath ?? page?.Path ?? "$", ScenePath = scenePath,
                Attribution = owner is null ? PresentationPreviewAttribution.Unmapped : PresentationPreviewAttribution.Direct,
                ZOrder = checked((uint)index),
            };
            if (element.Group is not null)
                foreach (var binding in Walk(element.Group.Children, pageId, scenePath + ".group.children", identities, provenance, page)) yield return binding;
            if (element.Diagram?.Drawing is not null)
                foreach (var binding in Walk(element.Diagram.Drawing.Children, pageId, scenePath + ".diagram.drawing.children", identities, provenance, page)) yield return binding;
        }
    }
}
