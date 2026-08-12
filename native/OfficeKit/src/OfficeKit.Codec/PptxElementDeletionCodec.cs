using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxElementDeletionPlan(
    bool Supported,
    string BlockedReason,
    uint NativeId,
    IReadOnlySet<string> RelationshipIds,
    IReadOnlySet<string> RemovedPackagePartPaths);

// A shape-tree deletion is safe only when removing the XML subtree cannot
// leave package relationships or native-identity consumers behind. Keep the
// eligibility proof separate from semantic editing: the caller may omit an
// element only after this module proves the concrete source graph again.
internal static class PptxElementDeletionCodec
{
    internal static PptxElementDeletionPlan Analyze(
        SlidePart slidePart,
        OpenXmlElement source,
        IReadOnlyList<OpenXmlElement> siblings)
    {
        if (source is not (P.Shape or P.Picture))
            return Blocked("only a top-level PresentationML shape or embedded picture is in the bounded deletion profile");

        var slide = slidePart.Slide;
        var common = slide?.CommonSlideData;
        if (slide is null || common?.ShapeTree is null)
            return Blocked("the slide has no canonical shape tree");
        if (!ReferenceEquals(source.Parent, common.ShapeTree))
            return Blocked("nested group children are not top-level shape-tree elements");
        var nativeId = NativeId(source);
        if (nativeId is null)
            return Blocked("the element has no unique native drawing ID");
        if (siblings.Count(element => NativeId(element) == nativeId) != 1)
            return Blocked($"native drawing ID {nativeId} is ambiguous", nativeId.Value);
        if (HasIdentitySensitiveSlideGraph(slide, common))
            return Blocked("slide timing or extension data may retain native element identity", nativeId.Value);
        if (PptxLegacyCommentsCodec.CommentPartPresent(slidePart))
            return Blocked("a slide comment graph may retain native element identity", nativeId.Value);
        if (siblings.OfType<P.ConnectionShape>().Any(connector => References(connector, nativeId.Value)))
            return Blocked($"connector topology references native drawing ID {nativeId}", nativeId.Value);

        return source switch
        {
            P.Shape when HasRelationshipReference(source) =>
                Blocked("the shape owns one or more package relationship references", nativeId.Value),
            P.Shape => Supported(nativeId.Value),
            P.Picture picture => AnalyzePicture(slidePart, picture, nativeId.Value),
            _ => Blocked("the element type is outside the bounded deletion profile", nativeId.Value),
        };
    }

    internal static void Apply(SlidePart slidePart, OpenXmlElement source, PptxElementDeletionPlan plan)
    {
        if (!plan.Supported)
            throw new InvalidOperationException("Cannot apply an unsupported presentation element deletion plan.");
        source.Remove();
        foreach (var relationshipId in plan.RelationshipIds)
        {
            if (!slidePart.DeletePart(relationshipId))
                throw new CodecException(
                    "presentation_element_delete_relationship_missing",
                    $"Presentation element deletion could not remove source relationship {relationshipId}.",
                    PartPath(slidePart));
        }
    }

    private static PptxElementDeletionPlan AnalyzePicture(SlidePart slidePart, P.Picture picture, uint nativeId)
    {
        var relationshipAttributes = RelationshipAttributes(picture).ToArray();
        var relationshipId = picture.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0 || relationshipAttributes.Length != 1 ||
            !string.Equals(relationshipAttributes[0].Value, relationshipId, StringComparison.Ordinal))
            return Blocked("the picture is not one canonical embedded-image relationship", nativeId);

        var slide = slidePart.Slide;
        if (slide is null || RelationshipAttributes(slide)
                .Where(attribute => string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal))
                .Count() != 1)
            return Blocked($"picture relationship {relationshipId} is referenced outside the element", nativeId);

        var edges = slidePart.Parts
            .Where(pair => pair.RelationshipId.Equals(relationshipId, StringComparison.Ordinal))
            .ToArray();
        if (edges.Length != 1 || edges[0].OpenXmlPart is not ImagePart imagePart)
            return Blocked($"picture relationship {relationshipId} does not resolve uniquely to an embedded image part", nativeId);
        if (slidePart.Parts.Count(pair => ReferenceEquals(pair.OpenXmlPart, imagePart)) != 1)
            return Blocked("the picture image part has another relationship from the same slide", nativeId);

        var removedParts = ExclusiveClosure(slidePart, imagePart);
        var removedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in removedParts)
        {
            removedPaths.Add(PartPath(part));
            removedPaths.Add(RelationshipPartPath(part));
        }
        var removedDataParts = removedParts
            .SelectMany(part => part.DataPartReferenceRelationships.Select(relationship => relationship.DataPart))
            .Distinct()
            .Where(dataPart => dataPart.GetDataPartReferenceRelationships().All(relationship =>
                relationship.Container is OpenXmlPart owner && removedParts.Contains(owner)))
            .ToArray();
        foreach (var dataPart in removedDataParts) removedPaths.Add(DataPartPath(dataPart));
        return new PptxElementDeletionPlan(
            true,
            string.Empty,
            nativeId,
            new HashSet<string>([relationshipId], StringComparer.Ordinal),
            removedPaths);
    }

    private static HashSet<OpenXmlPart> ExclusiveClosure(SlidePart owner, OpenXmlPart root)
    {
        var reachable = ReachableParts(root);
        var retained = new HashSet<OpenXmlPart>();
        var queue = new Queue<OpenXmlPart>();
        foreach (var part in reachable)
        {
            var outsideParent = part.GetParentParts().Any(parent =>
                !reachable.Contains(parent) &&
                !(ReferenceEquals(part, root) && ReferenceEquals(parent, owner)));
            if (!outsideParent) continue;
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
        return reachable.Where(part => !retained.Contains(part)).ToHashSet();
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

    private static bool HasRelationshipReference(OpenXmlElement source) => RelationshipAttributes(source).Any();

    private static IEnumerable<OpenXmlAttribute> RelationshipAttributes(OpenXmlElement source) =>
        Attributes(source).Where(attribute =>
            attribute.NamespaceUri is
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships");

    private static bool HasIdentitySensitiveSlideGraph(P.Slide slide, P.CommonSlideData common) =>
        slide.ChildElements.Any(element => element.LocalName is "timing" or "extLst") ||
        common.ChildElements.Any(element => element.LocalName is "extLst");

    private static IEnumerable<OpenXmlAttribute> Attributes(OpenXmlElement source)
    {
        foreach (var attribute in source.GetAttributes()) yield return attribute;
        foreach (var descendant in source.Descendants())
            foreach (var attribute in descendant.GetAttributes()) yield return attribute;
    }

    internal static uint? NativeId(OpenXmlElement source) => source switch
    {
        P.Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        P.Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
        P.ConnectionShape connector => connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
        P.GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        _ => null,
    };

    private static bool References(P.ConnectionShape connector, uint nativeId) =>
        connector.Descendants<A.StartConnection>().Any(connection => connection.Id?.Value == nativeId) ||
        connector.Descendants<A.EndConnection>().Any(connection => connection.Id?.Value == nativeId);

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

    private static PptxElementDeletionPlan Supported(uint nativeId) => new(
        true,
        string.Empty,
        nativeId,
        new HashSet<string>(StringComparer.Ordinal),
        new HashSet<string>(StringComparer.OrdinalIgnoreCase));

    private static PptxElementDeletionPlan Blocked(string reason, uint nativeId = 0) => new(
        false,
        reason,
        nativeId,
        new HashSet<string>(StringComparer.Ordinal),
        new HashSet<string>(StringComparer.OrdinalIgnoreCase));
}
