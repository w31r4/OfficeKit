using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// An opt-in observer of the existing writer plan, not a second lowering pass.
// The writer's exact slide reaches RecordNativeBindings after slide validation.
internal sealed class PpjPreviewSourceFreeBuildPlan : IPptxSourceFreeBuildPlan
{
    private readonly IPptxSourceFreeBuildPlan inner;
    private readonly PpjPreviewSceneBuilder.Collector collector;
    private readonly Dictionary<(string Page, string Id), PpjExpandedNodeModel> nodes;
    private readonly IReadOnlyDictionary<(string PageId, string Id), string>? origins;
    private readonly List<PresentationPreviewNodeBinding> bindings = [];

    internal PpjPreviewSourceFreeBuildPlan(IPptxSourceFreeBuildPlan inner,
        IEnumerable<PpjExpandedNodeModel> expandedNodes, EffectiveCodecLimits limits,
        IReadOnlyDictionary<(string PageId, string Id), string>? origins = null)
    {
        this.inner = inner;
        collector = new(inner.Presentation, limits);
        nodes = expandedNodes.ToDictionary(node => (node.PageId, node.Id));
        this.origins = origins;
    }

    public PresentationArtifact Presentation => inner.Presentation;
    public bool RequiresPreviousSlide(int index) => inner.RequiresPreviousSlide(index);
    public PresentationSlide MaterializeSlide(int index, PresentationSlide? previousSlide) =>
        inner.MaterializeSlide(index, previousSlide);

    public void RecordNativeBindings(int index, PresentationSlide slide, IReadOnlyList<PresentationElement> flattenedElements)
    {
        inner.RecordNativeBindings(index, slide, flattenedElements);
        collector.CaptureSlide(index, slide);
        Bind(slide.Elements, slide.Id, $"$.presentation.slides[{index}].elements", $"$.pages[{index}]", null);
    }

    internal PresentationPreviewScene Complete(string programSha256, string candidateSha256, IReadOnlyList<Asset> assets) =>
        collector.Complete(programSha256, candidateSha256, bindings, assets);

    private void Bind(IEnumerable<PresentationElement> elements, string pageId, string scenePath,
        string pagePath, PpjExpandedNodeModel? ancestor)
    {
        foreach (var (element, index) in elements.Select((element, index) => (element, index)))
        {
            var direct = nodes.TryGetValue((pageId, element.Id), out var node);
            var owner = direct ? node : ancestor;
            var ownerPath = owner is null ? null : origins?.GetValueOrDefault((pageId, owner.Id)) ??
                (owner.ComponentId is null ? owner.ProgramPath : null);
            var path = $"{scenePath}[{index}]";
            bindings.Add(new PresentationPreviewNodeBinding
            {
                PageId = pageId,
                SemanticId = owner?.Id ?? string.Empty,
                // Native IR identity; numeric OOXML cNvPr IDs remain writer-owned.
                NativeId = element.Id,
                ProgramPath = ownerPath ?? pagePath,
                ScenePath = path,
                Attribution = ownerPath is null ? PresentationPreviewAttribution.Unmapped :
                    direct && owner!.ComponentId is null ? PresentationPreviewAttribution.Direct : PresentationPreviewAttribution.Generated,
                SourceId = owner?.SourceId ?? string.Empty,
                ComponentId = owner?.ComponentId ?? string.Empty,
                InstanceId = owner?.InstanceId ?? string.Empty,
                RepeatKey = owner?.RepeatKey ?? string.Empty,
                ZOrder = checked((uint)index),
            });
            if (element.Group is not null)
                Bind(element.Group.Children, pageId, path + ".group.children", pagePath, owner);
            if (element.Diagram?.Drawing is not null)
                Bind(element.Diagram.Drawing.Children, pageId, path + ".diagram.drawing.children", pagePath, owner);
        }
    }
}
