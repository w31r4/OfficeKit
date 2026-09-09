using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeKit.Codec;

// Opt-in provenance for the existing expansion pass. Track actual model/clone
// identities, not local IDs: different component instances may reuse local IDs.
// No annotations are written into authored JSON or the existing node map.
internal sealed class PpjPreviewOrigins
{
    private readonly Dictionary<PpjElementModel, string> sources = new(ReferenceEqualityComparer.Instance);
    internal Dictionary<(string PageId, string Id), string> Expanded { get; } = [];

    internal PpjPreviewOrigins(PpjProgramModel program)
    {
        for (var i = 0; i < program.Pages.Count; i++)
            Index(program.Pages[i].Elements, $"$.pages[{i}].elements");
        for (var i = 0; i < program.Components.Count; i++)
            Index(program.Components[i].Elements, $"$.components[{i}].elements");
    }

    private void Index(IReadOnlyList<PpjElementModel> elements, string path)
    {
        for (var i = 0; i < elements.Count; i++)
        {
            var element = elements[i];
            var owner = $"{path}[{i}]";
            sources.Add(element, owner);
            if (element is PpjGroupElementModel group) Index(group.Elements, owner + ".elements");
            if (element is PpjComponentElementModel component)
                foreach (var slot in component.Slots)
                    Index(slot.Value, owner + ".slots[" + JsonSerializer.Serialize(slot.Key) + "]");
        }
    }

    internal void Record(string pageId, string id, PpjElementModel source)
    {
        if (sources.TryGetValue(source, out var path)) Expanded.Add((pageId, id), path);
    }

    internal Dictionary<JsonNode, string> TrackClone(PpjElementModel source, JsonObject clone)
    {
        var paths = new Dictionary<JsonNode, string>(ReferenceEqualityComparer.Instance);
        TrackClone(source, clone, paths);
        return paths;
    }

    internal void TrackClone(PpjElementModel source, JsonNode clone, Dictionary<JsonNode, string> paths) =>
        Walk(source, clone, (model, node) =>
        {
            if (sources.TryGetValue(model, out var path)) paths.Add(node, path);
        });

    internal void RecordParsed(PpjElementModel parsed, JsonObject clone, Dictionary<JsonNode, string> paths) =>
        Walk(parsed, clone, (model, node) =>
        {
            if (paths.TryGetValue(node, out var path)) sources.Add(model, path);
        });

    internal void ReleaseParsed(PpjElementModel parsed)
    {
        sources.Remove(parsed);
        if (parsed is PpjGroupElementModel group)
            foreach (var child in group.Elements) ReleaseParsed(child);
        if (parsed is PpjComponentElementModel component)
            foreach (var slot in component.Slots.Values)
                foreach (var child in slot) ReleaseParsed(child);
    }

    // Pair typed elements with their exact JSON clone before or after the
    // compiler's slot replacement. Array indexes here describe that same tree,
    // never an inferred correspondence between different expansion results.
    private static void Walk(PpjElementModel element, JsonNode node, Action<PpjElementModel, JsonNode> visit)
    {
        visit(element, node);
        if (element is PpjGroupElementModel group)
            for (var i = 0; i < group.Elements.Count; i++)
                Walk(group.Elements[i], node["elements"]![i]!, visit);
        if (element is PpjComponentElementModel component)
            foreach (var slot in component.Slots)
                for (var i = 0; i < slot.Value.Count; i++)
                    Walk(slot.Value[i], node["slots"]![slot.Key]![i]!, visit);
    }
}
