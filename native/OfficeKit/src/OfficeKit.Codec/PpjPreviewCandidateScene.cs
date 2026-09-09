using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class PpjPreviewCandidateScene
{
    internal static IReadOnlyList<Diagnostic> Attach(PptxPackageSource candidate,
        PresentationProgramResult receipt, EffectiveCodecLimits limits)
    {
        using var stage = PpjBuildProfiler.Measure("preview.candidate-import");
        // Import native candidate bytes directly. The projector can restore an
        // embedded authored program and therefore is not a visual-state oracle.
        var imported = PptxCodec.Import(candidate, limits, retainImportedAssetData: true,
            verifiedPackageSha256: receipt.OutputSha256);
        var presentation = imported.Artifact.Presentation ??
            throw new CodecException("invalid_preview_scene", "Candidate has no native presentation scene.");
        var scene = PpjPreviewSceneBuilder.Build(presentation, PresentationPreviewSceneOrigin.CandidateImport,
            receipt.ProgramSha256, receipt.OutputSha256, UnmappedBindings(presentation), imported.Artifact.Assets, limits);

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

    // Until an authoritative part/native-ID join is available, keep native
    // identity and exact scene paths. Ordinal import IDs must not be mistaken
    // for editable PPJ IDs after a reorder. Full semantic binding is task 2.3.
    private static IEnumerable<PresentationPreviewNodeBinding> UnmappedBindings(PresentationArtifact presentation)
    {
        foreach (var (slide, index) in presentation.Slides.Select((slide, index) => (slide, index)))
            foreach (var binding in Walk(slide.Elements, slide.Id, $"$.presentation.slides[{index}].elements"))
                yield return binding;
    }

    private static IEnumerable<PresentationPreviewNodeBinding> Walk(IEnumerable<PresentationElement> elements,
        string pageId, string path)
    {
        foreach (var (element, index) in elements.Select((element, index) => (element, index)))
        {
            var scenePath = $"{path}[{index}]";
            yield return new PresentationPreviewNodeBinding
            {
                PageId = pageId, NativeId = element.Id, ProgramPath = "$", ScenePath = scenePath,
                Attribution = PresentationPreviewAttribution.Unmapped, ZOrder = checked((uint)index),
            };
            if (element.Group is not null)
                foreach (var binding in Walk(element.Group.Children, pageId, scenePath + ".group.children")) yield return binding;
            if (element.Diagram?.Drawing is not null)
                foreach (var binding in Walk(element.Diagram.Drawing.Children, pageId, scenePath + ".diagram.drawing.children")) yield return binding;
        }
    }
}
