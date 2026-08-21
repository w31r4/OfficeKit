using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxElementDeletionPlan(
    bool Supported,
    string BlockedReason,
    uint NativeId,
    IReadOnlySet<string> RelationshipIds,
    IReadOnlySet<string> ReferenceRelationshipIds,
    IReadOnlySet<string> RemovedPackagePartPaths);

// A shape-tree deletion is safe only when removing the XML subtree cannot
// leave package relationships or native-identity consumers behind. Keep the
// eligibility proof separate from semantic editing: the caller may omit an
// element only after this module proves the concrete source graph again.
internal static class PptxElementDeletionCodec
{
    private const string SlideCreationIdExtensionUri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";
    private const string PowerPoint2010MainNamespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    internal static PptxElementDeletionPlan Analyze(
        SlidePart slidePart,
        OpenXmlElement source,
        IReadOnlyList<OpenXmlElement> siblings,
        bool allowDuplicateNativeIds = false)
    {
        if (source is not (P.Shape or P.Picture or P.ConnectionShape or P.GraphicFrame or P.GroupShape))
            return Blocked("only a bounded top-level PresentationML shape, picture, connector, table, chart, or group is in the deletion profile");

        var slide = slidePart.Slide;
        var common = slide?.CommonSlideData;
        if (slide is null || common?.ShapeTree is null)
            return Blocked("the slide has no canonical shape tree");
        if (!ReferenceEquals(source.Parent, common.ShapeTree))
            return Blocked("nested group children are not top-level shape-tree elements");
        var nativeId = NativeId(source);
        if (nativeId is null)
            return Blocked("the element has no unique native drawing ID");
        var ownedIds = NativeIds(source);
        var slideIds = NativeIdOccurrences(common.ShapeTree);
        if (!allowDuplicateNativeIds && ownedIds.Any(id => slideIds.Count(candidate => candidate == id) != 1))
            return Blocked($"native drawing ID {nativeId} or one of its descendants is ambiguous", nativeId.Value);
        if (HasIdentitySensitiveSlideGraph(slide, common))
            return Blocked("slide timing or extension data may retain native element identity", nativeId.Value);
        if (PptxLegacyCommentsCodec.CommentPartPresent(slidePart))
            return Blocked("a slide comment graph may retain native element identity", nativeId.Value);
        if (slidePart.Parts.Any(pair => pair.OpenXmlPart is PowerPointCommentPart))
            return Blocked("a slide comment graph may retain native element identity", nativeId.Value);
        var ownedElements = source.Descendants().Prepend(source).ToHashSet();
        if (common.ShapeTree.Descendants<P.ConnectionShape>()
                .Where(connector => !ownedElements.Contains(connector))
                .Any(connector => ownedIds.Any(id => References(connector, id))))
            return Blocked($"connector topology references native drawing ID {nativeId} or one of its descendants", nativeId.Value);

        return source switch
        {
            P.Shape when HasRelationshipReference(source) =>
                AnalyzeRelationshipClosure(slidePart, source, nativeId.Value),
            P.Shape => Supported(nativeId.Value),
            P.Picture picture => AnalyzePicture(slidePart, picture, nativeId.Value),
            P.ConnectionShape when HasRelationshipReference(source) =>
                AnalyzeRelationshipClosure(slidePart, source, nativeId.Value),
            P.ConnectionShape => Supported(nativeId.Value),
            P.GraphicFrame frame => AnalyzeGraphicFrame(slidePart, frame, nativeId.Value),
            P.GroupShape group => AnalyzeGroup(slidePart, group, nativeId.Value),
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
            if (plan.ReferenceRelationshipIds.Contains(relationshipId))
            {
                slidePart.DeleteReferenceRelationship(relationshipId);
                continue;
            }
            if (!slidePart.DeletePart(relationshipId))
                throw new CodecException(
                    "presentation_element_delete_relationship_missing",
                    $"Presentation element deletion could not remove source relationship {relationshipId}.",
                    PartPath(slidePart));
        }
    }

    private static PptxElementDeletionPlan AnalyzeGroup(SlidePart slidePart, P.GroupShape group, uint nativeId)
    {
        var relationshipIds = RelationshipAttributes(group)
            .Select(attribute => attribute.Value)
            .Where(value => !string.IsNullOrEmpty(value))
            .Select(value => value!)
            .ToHashSet(StringComparer.Ordinal);
        if (relationshipIds.Count == 0) return Supported(nativeId);

        var slide = slidePart.Slide!;
        var rootParts = new HashSet<OpenXmlPart>();
        var referenceRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationshipId in relationshipIds)
        {
            var ownedReferences = RelationshipAttributes(group).Count(attribute =>
                string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal));
            var slideReferences = RelationshipAttributes(slide).Count(attribute =>
                string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal));
            if (ownedReferences == 0 || ownedReferences != slideReferences)
                return Blocked($"group relationship {relationshipId} is referenced outside the group", nativeId);

            var partEdges = slidePart.Parts.Where(pair => pair.RelationshipId.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var hyperlinks = slidePart.HyperlinkRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var externals = slidePart.ExternalRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var dataParts = slidePart.DataPartReferenceRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            if (partEdges.Length + hyperlinks.Length + externals.Length + dataParts.Length != 1)
                return Blocked($"group relationship {relationshipId} does not resolve uniquely", nativeId);
            if (dataParts.Length != 0)
                return Blocked($"group relationship {relationshipId} is an unsupported data-part reference", nativeId);
            if (hyperlinks.Length != 0 || externals.Length != 0)
            {
                referenceRelationshipIds.Add(relationshipId);
                continue;
            }
            rootParts.Add(partEdges[0].OpenXmlPart);
        }

        foreach (var rootPart in rootParts)
        {
            var ownerRelationshipIds = slidePart.Parts
                .Where(pair => ReferenceEquals(pair.OpenXmlPart, rootPart))
                .Select(pair => pair.RelationshipId)
                .ToArray();
            if (ownerRelationshipIds.Any(relationshipId => !relationshipIds.Contains(relationshipId)))
                return Blocked("a group-owned package part has another relationship from the same slide", nativeId);
        }

        var removedParts = ExclusiveClosure(slidePart, rootParts);
        var removedPaths = RemovedPackagePaths(removedParts);
        return new PptxElementDeletionPlan(
            true,
            string.Empty,
            nativeId,
            relationshipIds,
            referenceRelationshipIds,
            removedPaths);
    }

    private static PptxElementDeletionPlan AnalyzePicture(SlidePart slidePart, P.Picture picture, uint nativeId)
    {
        var relationshipId = picture.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0)
            return Blocked("the picture is not one canonical embedded-image relationship", nativeId);

        var canonical = AnalyzeOwnedPartRelationship<ImagePart>(slidePart, picture, nativeId, relationshipId, "picture", "embedded image part");
        if (canonical.Supported) return canonical;

        // Some Office producers retain a second, independently owned image
        // relationship (for example a WDP/HD Photo fallback) on the same
        // picture.  The canonical single-image profile above is intentionally
        // strict, but a complete relationship-closure proof can still show
        // that removing the picture owns every referenced part.  Do not
        // infer this from content type or discard the fallback: only the
        // generic closure proof may authorize the deletion.
        return AnalyzeRelationshipClosure(slidePart, picture, nativeId);
    }

    private static PptxElementDeletionPlan AnalyzeGraphicFrame(SlidePart slidePart, P.GraphicFrame frame, uint nativeId)
    {
        var graphicData = frame.Graphic?.GraphicData;
        if (graphicData is null)
            return Blocked("the graphic frame has no canonical DrawingML payload", nativeId);
        if (string.Equals(graphicData.Uri?.Value, "http://schemas.openxmlformats.org/drawingml/2006/table", StringComparison.Ordinal))
        {
            if (!PptxTableCodec.TryRead(frame, out _))
                return Blocked("the graphic frame is outside the bounded DrawingML table profile", nativeId);
            return HasRelationshipReference(frame)
                ? Blocked("the table owns one or more package relationship references", nativeId)
                : Supported(nativeId);
        }
        if (!string.Equals(graphicData.Uri?.Value, "http://schemas.openxmlformats.org/drawingml/2006/chart", StringComparison.Ordinal))
            return AnalyzeRelationshipClosure(slidePart, frame, nativeId);
        var references = graphicData.Elements<C.ChartReference>().ToArray();
        if (references.Length != 1 || references[0].Id?.Value is not { Length: > 0 } relationshipId)
            return Blocked("the chart is not one canonical internal ChartPart relationship", nativeId);
        return AnalyzeOwnedPartRelationship<ChartPart>(slidePart, frame, nativeId, relationshipId, "chart", "ChartPart");
    }

    // An unmodeled top-level object may still be safely removed from a cloned
    // slide when its relationship closure is independently owned by that
    // element. This is a deletion proof only: the object remains opaque and
    // no semantic parser or serializer is introduced for its native graph.
    private static PptxElementDeletionPlan AnalyzeRelationshipClosure(
        SlidePart slidePart,
        OpenXmlElement source,
        uint nativeId)
    {
        var relationshipIds = RelationshipAttributes(source)
            .Select(attribute => attribute.Value)
            .Where(value => !string.IsNullOrEmpty(value))
            .Select(value => value!)
            .ToHashSet(StringComparer.Ordinal);
        if (relationshipIds.Count == 0) return Supported(nativeId);
        var slide = slidePart.Slide;
        if (slide is null) return Blocked("the slide root is unavailable for relationship-closure proof", nativeId);
        var rootParts = new HashSet<OpenXmlPart>();
        var referenceRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationshipId in relationshipIds)
        {
            var ownedReferences = RelationshipAttributes(source).Count(attribute =>
                string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal));
            var slideReferences = RelationshipAttributes(slide).Count(attribute =>
                string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal));
            if (ownedReferences == 0 || ownedReferences != slideReferences)
                return Blocked($"relationship {relationshipId} is referenced outside the element", nativeId);

            var partEdges = slidePart.Parts.Where(pair => pair.RelationshipId.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var hyperlinks = slidePart.HyperlinkRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var externals = slidePart.ExternalRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            var dataParts = slidePart.DataPartReferenceRelationships.Where(relationship => relationship.Id.Equals(relationshipId, StringComparison.Ordinal)).ToArray();
            if (partEdges.Length + hyperlinks.Length + externals.Length + dataParts.Length != 1)
                return Blocked($"relationship {relationshipId} does not resolve uniquely", nativeId);
            if (dataParts.Length != 0)
                return Blocked($"relationship {relationshipId} is an unsupported data-part reference", nativeId);
            if (hyperlinks.Length != 0 || externals.Length != 0)
            {
                referenceRelationshipIds.Add(relationshipId);
                continue;
            }
            rootParts.Add(partEdges[0].OpenXmlPart);
        }

        foreach (var rootPart in rootParts)
        {
            var ownerRelationshipIds = slidePart.Parts
                .Where(pair => ReferenceEquals(pair.OpenXmlPart, rootPart))
                .Select(pair => pair.RelationshipId)
                .ToArray();
            if (ownerRelationshipIds.Any(relationshipId => !relationshipIds.Contains(relationshipId)))
                return Blocked("a relationship-owned package part has another relationship from the same slide", nativeId);
        }

        var removedParts = ExclusiveClosure(slidePart, rootParts);
        return new PptxElementDeletionPlan(
            true,
            string.Empty,
            nativeId,
            relationshipIds,
            referenceRelationshipIds,
            RemovedPackagePaths(removedParts));
    }

    private static PptxElementDeletionPlan AnalyzeOwnedPartRelationship<TPart>(
        SlidePart slidePart,
        OpenXmlElement source,
        uint nativeId,
        string relationshipId,
        string elementKind,
        string partKind)
        where TPart : OpenXmlPart
    {
        var relationshipAttributes = RelationshipAttributes(source).ToArray();
        if (relationshipAttributes.Length != 1 ||
            !string.Equals(relationshipAttributes[0].Value, relationshipId, StringComparison.Ordinal))
            return Blocked($"the {elementKind} is not one canonical internal {partKind} relationship", nativeId);

        var slide = slidePart.Slide;
        if (slide is null || RelationshipAttributes(slide)
                .Where(attribute => string.Equals(attribute.Value, relationshipId, StringComparison.Ordinal))
                .Count() != 1)
            return Blocked($"{elementKind} relationship {relationshipId} is referenced outside the element", nativeId);

        var edges = slidePart.Parts
            .Where(pair => pair.RelationshipId.Equals(relationshipId, StringComparison.Ordinal))
            .ToArray();
        if (edges.Length != 1 || edges[0].OpenXmlPart is not TPart rootPart)
            return Blocked($"{elementKind} relationship {relationshipId} does not resolve uniquely to an internal {partKind}", nativeId);
        if (slidePart.Parts.Count(pair => ReferenceEquals(pair.OpenXmlPart, rootPart)) != 1)
            return Blocked($"the {elementKind} {partKind} has another relationship from the same slide", nativeId);

        var removedParts = ExclusiveClosure(slidePart, new HashSet<OpenXmlPart> { rootPart });
        var removedPaths = RemovedPackagePaths(removedParts);
        return new PptxElementDeletionPlan(
            true,
            string.Empty,
            nativeId,
            new HashSet<string>([relationshipId], StringComparer.Ordinal),
            new HashSet<string>(StringComparer.Ordinal),
            removedPaths);
    }

    private static HashSet<OpenXmlPart> ExclusiveClosure(SlidePart owner, IReadOnlySet<OpenXmlPart> roots)
    {
        var reachable = new HashSet<OpenXmlPart>();
        foreach (var root in roots) reachable.UnionWith(ReachableParts(root));
        var retained = new HashSet<OpenXmlPart>();
        var queue = new Queue<OpenXmlPart>();
        foreach (var part in reachable)
        {
            var outsideParent = part.GetParentParts().Any(parent =>
                !reachable.Contains(parent) &&
                !(roots.Contains(part) && ReferenceEquals(parent, owner)));
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

    private static HashSet<string> RemovedPackagePaths(IReadOnlySet<OpenXmlPart> removedParts)
    {
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
        return removedPaths;
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
        common.ChildElements.Any(element => element.LocalName == "extLst" && !IsSafeSlideCreationIdExtensionList(element));

    // PowerPoint writes this slide-level creationId extension on otherwise
    // ordinary slides. It identifies the slide itself, not a shape-tree
    // child, so deleting a proven sibling leaves the extension byte-for-byte
    // intact and cannot create a dangling object reference. Every other
    // extension remains identity-sensitive until its schema is understood.
    private static bool IsSafeSlideCreationIdExtensionList(OpenXmlElement extensionList)
    {
        if (extensionList.NamespaceUri != "http://schemas.openxmlformats.org/presentationml/2006/main" ||
            extensionList.ChildElements.Count != 1 ||
            extensionList.ChildElements[0].LocalName != "ext" ||
            extensionList.ChildElements[0].NamespaceUri != "http://schemas.openxmlformats.org/presentationml/2006/main")
            return false;
        var extension = extensionList.ChildElements[0];
        var extensionAttributes = extension.GetAttributes().ToArray();
        if (extensionAttributes.Length != 1 || extensionAttributes[0].LocalName != "uri" || extensionAttributes[0].NamespaceUri.Length != 0 ||
            !string.Equals(extensionAttributes[0].Value, SlideCreationIdExtensionUri, StringComparison.OrdinalIgnoreCase) ||
            extension.ChildElements.Count != 1)
            return false;
        var creationId = extension.ChildElements[0];
        if (creationId.LocalName != "creationId" || creationId.NamespaceUri != PowerPoint2010MainNamespace || creationId.ChildElements.Count != 0)
            return false;
        var creationAttributes = creationId.GetAttributes().ToArray();
        return creationAttributes.Length == 1 && creationAttributes[0].LocalName == "val" && creationAttributes[0].NamespaceUri.Length == 0 &&
            uint.TryParse(creationAttributes[0].Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out _);
    }

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

    internal static IReadOnlySet<uint> NativeIds(OpenXmlElement source) => source
        .Descendants()
        .Prepend(source)
        .Select(NativeId)
        .Where(id => id is not null)
        .Select(id => id!.Value)
        .ToHashSet();

    private static IReadOnlyList<uint> NativeIdOccurrences(OpenXmlElement source) => source
        .Descendants()
        .Prepend(source)
        .Select(NativeId)
        .Where(id => id is not null)
        .Select(id => id!.Value)
        .ToArray();

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
        new HashSet<string>(StringComparer.Ordinal),
        new HashSet<string>(StringComparer.OrdinalIgnoreCase));

    private static PptxElementDeletionPlan Blocked(string reason, uint nativeId = 0) => new(
        false,
        reason,
        nativeId,
        new HashSet<string>(StringComparer.Ordinal),
        new HashSet<string>(StringComparer.Ordinal),
        new HashSet<string>(StringComparer.OrdinalIgnoreCase));
}
