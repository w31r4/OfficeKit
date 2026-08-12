using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PptxSlideDeletionPlan(
    bool Supported,
    string BlockedReason,
    IReadOnlySet<string> RemovedPackagePartPaths,
    uint OwnedPartCount);

// Slide deletion is an OPC ownership operation, not a catalog of
// PresentationML object types. Build the complete graph reachable from every
// requested SlidePart, retain every subgraph that is still reachable from an
// outside parent, and let Open XML SDK delete the remaining exclusive closure. This
// keeps charts, notes, comments, OLE, diagrams, media, and future closed leaves
// on one audited path while shared layout/master/theme/media parts survive.
internal static class PptxSlideDeletionCodec
{
    internal static PptxSlideDeletionPlan Analyze(
        PresentationPart presentationPart,
        PptxSourceSlideEntry source,
        OpaqueOpcGraph opaque) => AnalyzeTransaction(presentationPart, [source], opaque);

    internal static PptxSlideDeletionPlan AnalyzeTransaction(
        PresentationPart presentationPart,
        IReadOnlyList<PptxSourceSlideEntry> sources,
        OpaqueOpcGraph opaque)
    {
        var root = presentationPart.Presentation;
        if (root is null)
            return Blocked("the package has no Presentation root");
        if (root.ChildElements.Any(element => element.LocalName is "custShowLst" or "sectionLst" or "extLst"))
            return Blocked("presentation-level custom shows, sections, or extension data may retain slide identity");

        foreach (var source in sources)
        {
            var blockedReason = SourceBlockedReason(presentationPart, root, source, opaque);
            if (blockedReason is not null) return Blocked($"slide {source.Index + 1}: {blockedReason}");
        }

        var sourceParts = sources.Select(source => source.Part).ToHashSet();
        var reachable = new HashSet<OpenXmlPart>();
        foreach (var sourcePart in sourceParts) reachable.UnionWith(ReachableParts(sourcePart));
        var retained = new HashSet<OpenXmlPart>();
        var queue = new Queue<OpenXmlPart>();
        foreach (var part in reachable)
        {
            if (sourceParts.Contains(part)) continue;
            if (!part.GetParentParts().Any(parent => !reachable.Contains(parent))) continue;
            retained.Add(part);
            queue.Enqueue(part);
        }
        while (queue.TryDequeue(out var retainedPart))
        {
            foreach (var child in retainedPart.Parts.Select(pair => pair.OpenXmlPart))
            {
                if (reachable.Contains(child) && retained.Add(child)) queue.Enqueue(child);
            }
        }

        var removedParts = reachable.Where(part => !retained.Contains(part)).ToHashSet();
        var removedDataParts = removedParts
            .SelectMany(part => part.DataPartReferenceRelationships.Select(relationship => relationship.DataPart))
            .Distinct()
            .Where(dataPart => dataPart.GetDataPartReferenceRelationships().All(relationship =>
                relationship.Container is OpenXmlPart owner && removedParts.Contains(owner)))
            .ToArray();
        var removedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in removedParts)
        {
            removedPaths.Add(PartPath(part));
            removedPaths.Add(RelationshipPartPath(part));
        }
        foreach (var dataPart in removedDataParts) removedPaths.Add(DataPartPath(dataPart));
        return new PptxSlideDeletionPlan(
            true,
            string.Empty,
            removedPaths,
            checked((uint)(removedParts.Count + removedDataParts.Length)));
    }

    private static string? SourceBlockedReason(
        PresentationPart presentationPart,
        DocumentFormat.OpenXml.Presentation.Presentation root,
        PptxSourceSlideEntry source,
        OpaqueOpcGraph opaque)
    {
        var slidePath = PartPath(source.Part);
        var presentationEdges = presentationPart.Parts
            .Where(pair => ReferenceEquals(pair.OpenXmlPart, source.Part))
            .ToArray();
        if (presentationEdges.Length != 1 ||
            !presentationEdges[0].RelationshipId.Equals(source.RelationshipId, StringComparison.Ordinal))
            return "its PresentationPart relationship is missing or ambiguous";

        var otherParents = source.Part.GetParentParts()
            .Where(parent => !ReferenceEquals(parent, presentationPart))
            .Select(PartPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(4)
            .ToArray();
        if (otherParents.Length > 0)
            return $"it is referenced by {string.Join(", ", otherParents)}";

        var opaqueInbound = opaque.PackageRelationships
            .Where(relationship => !relationship.TargetMode.Equals("External", StringComparison.OrdinalIgnoreCase))
            .Where(relationship => ResolveTarget(relationship.SourcePath, relationship.Target)
                .Equals(slidePath, StringComparison.OrdinalIgnoreCase))
            .Where(relationship => !relationship.SourcePath.Equals(PartPath(presentationPart), StringComparison.OrdinalIgnoreCase))
            .Select(relationship => relationship.SourcePath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(4)
            .ToArray();
        if (opaqueInbound.Length > 0)
            return $"an opaque package relationship references it from {string.Join(", ", opaqueInbound)}";

        var relationshipReferencesOutsideSlideIds = root
            .Descendants()
            .Where(element => element is not DocumentFormat.OpenXml.Presentation.SlideId)
            .SelectMany(element => element.GetAttributes())
            .Any(attribute =>
                (attribute.NamespaceUri is "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
                    "http://purl.oclc.org/ooxml/officeDocument/relationships") &&
                string.Equals(attribute.Value, source.RelationshipId, StringComparison.Ordinal));
        return relationshipReferencesOutsideSlideIds
            ? "a presentation-level relationship outside p:sldIdLst still references it"
            : null;
    }

    private static HashSet<OpenXmlPart> ReachableParts(OpenXmlPart root)
    {
        var reachable = new HashSet<OpenXmlPart> { root };
        var queue = new Queue<OpenXmlPart>();
        queue.Enqueue(root);
        while (queue.TryDequeue(out var part))
        {
            foreach (var child in part.Parts.Select(pair => pair.OpenXmlPart))
            {
                if (reachable.Add(child)) queue.Enqueue(child);
            }
        }
        return reachable;
    }

    private static PptxSlideDeletionPlan Blocked(string reason) =>
        new(false, reason, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0);

    private static string ResolveTarget(string sourcePath, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var sourceDirectory = sourcePath.Contains('/')
            ? sourcePath[..sourcePath.LastIndexOf('/')]
            : string.Empty;
        var baseUri = new Uri($"http://office-kit.invalid/{(sourceDirectory.Length > 0 ? $"{sourceDirectory}/" : string.Empty)}");
        return Uri.UnescapeDataString(new Uri(baseUri, target).AbsolutePath).TrimStart('/');
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
}
