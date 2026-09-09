namespace OfficeKit.Codec;

// Joins only identities issued by the exact source projection and actual native
// candidate import. Neither visual proximity nor ordinal IDs establish ownership.
internal sealed class PpjPreviewCandidateBindings
{
    private readonly Dictionary<(string PartPath, uint NativeId), PpjExpandedNodeModel> owners = [];
    private readonly Dictionary<string, (string Id, string Path)> pages = new(StringComparer.Ordinal);

    internal PpjPreviewCandidateBindings(IReadOnlyList<PptxNativeBinding> source, PpjExpansionResult requested)
    {
        var requestedNodes = requested.Nodes.ToDictionary(node => (node.PageId, node.Id));
        var requestedPages = requested.Pages.Select((page, index) => (page.Id, Path: $"$.pages[{index}]"))
            .ToDictionary(page => page.Id, StringComparer.Ordinal);
        foreach (var entries in source.GroupBy(binding => binding.PartPath, StringComparer.Ordinal))
        {
            var pageIds = entries.Select(binding => binding.PageId).Distinct(StringComparer.Ordinal).ToArray();
            if (pageIds.Length == 1 && requestedPages.TryGetValue(pageIds[0], out var page)) pages[entries.Key] = page;
        }
        foreach (var entries in source.GroupBy(binding => (binding.PartPath, binding.NativeId)))
        {
            // Native IDs must be nonzero and unique in their source part.
            if (entries.Key.NativeId == 0 || entries.Count() != 1) continue;
            var entry = entries.Single();
            if (requestedNodes.TryGetValue((entry.PageId, entry.ElementId), out var node)) owners[entries.Key] = node;
        }
    }

    internal PpjExpandedNodeModel? Owner(PptxNativeBinding identity) =>
        owners.GetValueOrDefault((identity.PartPath, identity.NativeId));

    internal (string Id, string Path)? Page(string partPath) => pages.TryGetValue(partPath, out var page) ? page : null;
}
