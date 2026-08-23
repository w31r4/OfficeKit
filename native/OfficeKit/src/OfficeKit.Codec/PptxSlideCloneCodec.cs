using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxSlideClonePlan(
    bool Supported,
    string BlockedReason,
    uint ClonedPartCount,
    uint SharedPartCount);

internal sealed record PptxSlideCloneResult(
    SlidePart Part,
    IReadOnlySet<string> ChangedPackagePaths,
    IReadOnlySet<string> AddedOpaquePartPaths,
    IReadOnlySet<string> AddedOpaqueRelationshipKeys,
    IReadOnlyDictionary<string, string> CopiedPartSourcePaths);

// Slide cloning is an OPC graph-copy operation. Open XML SDK already owns the
// difficult cross-package copy algorithm: it preserves relationship IDs,
// repeated graph nodes, external relationships, and DataParts while allocating
// collision-free package URIs. This module adds the Presentation-specific ownership
// policy around that primitive. Mutable slide-owned descendants are copied;
// proven immutable or identity-bearing resources are rebound to their source
// package objects; connected unknown topology fails closed.
internal static class PptxSlideCloneCodec
{
    private const int MaxClonedParts = 2_048;
    private const int MaxDataParts = 512;

    internal static PptxSlideClonePlan Analyze(
        PresentationPart presentationPart,
        PptxSourceSlideEntry source,
        IReadOnlySet<SlidePart> retainedSlideParts,
        IReadOnlySet<int>? omittedShapeTreeIndices = null)
    {
        if (PptxSectionCodec.HasSectionGraph(presentationPart))
            return Blocked("the presentation contains a PowerPoint section graph whose slide identity cannot yet be extended safely");
        if (source.Part.Parts.Any(pair => pair.OpenXmlPart is PowerPointCommentPart))
            return Blocked("its modern comment part embeds native slide identity that requires a dedicated rewrite");

        var cloneOnlyParts = CloneOnlyParts(source.Part, omittedShapeTreeIndices);
        var owned = new HashSet<OpenXmlPart> { source.Part };
        var shared = new HashSet<OpenXmlPart>();
        var queue = new Queue<OpenXmlPart>();
        queue.Enqueue(source.Part);
        while (queue.TryDequeue(out var owner))
        {
            foreach (var pair in owner.Parts)
            {
                var child = pair.OpenXmlPart;
                if (SamePart(child, source.Part)) continue;
                if (ShouldShare(child))
                {
                    if (child is SlidePart targetSlide && !retainedSlideParts.Contains(targetSlide))
                        return Blocked($"slide-jump relationship {pair.RelationshipId} targets a slide that is not retained");
                    shared.Add(child);
                    continue;
                }
                if (owned.Add(child)) queue.Enqueue(child);
                if (owned.Count > MaxClonedParts)
                    return Blocked($"its owned descendant graph exceeds the {MaxClonedParts}-part clone budget");
            }
        }

        foreach (var part in owned)
        {
            if (SamePart(part, source.Part)) continue;
            var outsideParents = part.GetParentParts()
                .Where(parent => !owned.Contains(parent))
                .Select(PartPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(4)
                .ToArray();
            if (outsideParents.Length > 0 && !cloneOnlyParts.Contains(part))
                return Blocked($"owned part {PartPath(part)} is also referenced from {string.Join(", ", outsideParents)}");
        }

        var dataParts = owned
            .SelectMany(part => part.DataPartReferenceRelationships.Select(relationship => relationship.DataPart))
            .Distinct()
            .Take(MaxDataParts + 1)
            .Count();
        if (dataParts > MaxDataParts)
            return Blocked($"its owned graph exceeds the {MaxDataParts}-DataPart clone budget");

        return new PptxSlideClonePlan(
            true,
            string.Empty,
            checked((uint)owned.Count),
            checked((uint)shared.Count));
    }

    internal static PptxSlideCloneResult Clone(
        PresentationPart presentationPart,
        PptxSourceSlideEntry source,
        IReadOnlySet<SlidePart> retainedSlideParts,
        IReadOnlySet<int>? omittedShapeTreeIndices = null)
    {
        var plan = Analyze(presentationPart, source, retainedSlideParts, omittedShapeTreeIndices);
        if (!plan.Supported) throw Unsupported(source, plan.BlockedReason);

        using var scratchStream = new MemoryStream();
        using var scratch = PresentationDocument.Create(
            scratchStream,
            PresentationDocumentType.Presentation,
            autoSave: false);
        var scratchPresentation = scratch.AddPresentationPart();
        var scratchSlide = scratchPresentation.AddPart(source.Part, "rIdCloneSource");

        // The second cross-package AddPart is the single graph-copy primitive.
        // Reconciliation below replaces copied identity/immutable boundaries
        // with the corresponding original-package resource.
        // Open XML SDK's parameterless AddPart allocates a process-random
        // relationship ID.  A clone is a source-bound operation, so derive
        // the presentation relationship from the source part and the stable
        // current slide ordinal instead; this keeps repeated exports
        // byte-stable at the OPC content level without changing source IDs.
        var clone = presentationPart.AddPart(scratchSlide, CloneRelationshipId(presentationPart, source));
        var sourceToClone = new Dictionary<OpenXmlPart, OpenXmlPart>
        {
            [source.Part] = clone,
        };
        Reconcile(source.Part, clone, source.Part, clone, retainedSlideParts, sourceToClone);
        ValidateMappedGraph(source.Part, clone, source.Part, clone, retainedSlideParts, sourceToClone);

        var changedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var opaquePartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var opaqueRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        var cloneParts = sourceToClone.Values.Distinct().ToArray();
        foreach (var part in cloneParts)
        {
            var partPath = PartPath(part);
            changedPaths.Add(partPath);
            if (!OpcPackageProfile.Pptx.OwnsPath(partPath)) opaquePartPaths.Add(partPath);
            if (HasRelationships(part))
            {
                var relationshipPath = RelationshipPartPath(part);
                changedPaths.Add(relationshipPath);
                if (!OpcPackageProfile.Pptx.OwnsPath(relationshipPath)) opaquePartPaths.Add(relationshipPath);
            }
            foreach (var relationshipId in RelationshipIds(part))
                opaqueRelationshipKeys.Add($"{partPath}\0{relationshipId}");
        }
        foreach (var dataPart in cloneParts
                     .SelectMany(part => part.DataPartReferenceRelationships.Select(relationship => relationship.DataPart))
                     .Distinct())
        {
            var dataPath = DataPartPath(dataPart);
            changedPaths.Add(dataPath);
            if (!OpcPackageProfile.Pptx.OwnsPath(dataPath)) opaquePartPaths.Add(dataPath);
        }

        var copiedPartSourcePaths = sourceToClone
            .Where(pair => !SamePart(pair.Key, pair.Value))
            .ToDictionary(
                pair => PartPath(pair.Value),
                pair => PartPath(pair.Key),
                StringComparer.OrdinalIgnoreCase);

        return new PptxSlideCloneResult(clone, changedPaths, opaquePartPaths, opaqueRelationshipKeys, copiedPartSourcePaths);
    }

    internal static void Validate(
        PptxSourceSlideEntry source,
        SlidePart clone,
        IReadOnlySet<SlidePart> retainedSlideParts)
    {
        var sourceToClone = new Dictionary<OpenXmlPart, OpenXmlPart>
        {
            [source.Part] = clone,
        };
        ValidateMappedGraph(source.Part, clone, source.Part, clone, retainedSlideParts, sourceToClone);
    }

    internal static CodecException Unsupported(PptxSourceSlideEntry source, string reason) =>
        new(
            "unsupported_presentation_slide_clone",
            $"Source-preserving PPTX cloning requires a closed, uniquely owned OPC descendant graph with only proven shared identity resources; slide {source.Index + 1} cannot be cloned because {reason}.",
            PartPath(source.Part));

    private static void Reconcile(
        OpenXmlPart sourceOwner,
        OpenXmlPart cloneOwner,
        SlidePart sourceRoot,
        SlidePart cloneRoot,
        IReadOnlySet<SlidePart> retainedSlideParts,
        IDictionary<OpenXmlPart, OpenXmlPart> sourceToClone)
    {
        AssertRelationshipShellEqual(sourceOwner, cloneOwner);
        var cloneChildren = cloneOwner.Parts.ToDictionary(pair => pair.RelationshipId, StringComparer.Ordinal);
        foreach (var sourcePair in sourceOwner.Parts)
        {
            if (!cloneChildren.TryGetValue(sourcePair.RelationshipId, out var clonePair))
                throw Mismatch(sourceRoot, $"relationship {sourcePair.RelationshipId} disappeared during graph copy");
            var sourceChild = sourcePair.OpenXmlPart;
            var cloneChild = clonePair.OpenXmlPart;
            if (SamePart(sourceChild, sourceRoot))
            {
                if (!SamePart(cloneChild, cloneRoot))
                    Rebind(cloneOwner, cloneChild, cloneRoot, sourcePair.RelationshipId);
                continue;
            }
            if (ShouldShare(sourceChild))
            {
                if (sourceChild is SlidePart targetSlide && !retainedSlideParts.Contains(targetSlide))
                    throw new CodecException(
                        "unsupported_presentation_slide_clone",
                        $"Source-preserving PPTX cloning cannot retain slide-jump relationship {sourcePair.RelationshipId} because its target slide is not retained.",
                        PartPath(sourceRoot));
                if (!SamePart(cloneChild, sourceChild))
                    Rebind(cloneOwner, cloneChild, sourceChild, sourcePair.RelationshipId);
                continue;
            }
            if (sourceToClone.TryGetValue(sourceChild, out var mapped))
            {
                if (!SamePart(cloneChild, mapped))
                    Rebind(cloneOwner, cloneChild, mapped, sourcePair.RelationshipId);
                continue;
            }
            sourceToClone[sourceChild] = cloneChild;
            Reconcile(sourceChild, cloneChild, sourceRoot, cloneRoot, retainedSlideParts, sourceToClone);
        }
    }

    private static void ValidateMappedGraph(
        OpenXmlPart sourceOwner,
        OpenXmlPart cloneOwner,
        SlidePart sourceRoot,
        SlidePart cloneRoot,
        IReadOnlySet<SlidePart> retainedSlideParts,
        IDictionary<OpenXmlPart, OpenXmlPart> sourceToClone)
    {
        if (!sourceOwner.ContentType.Equals(cloneOwner.ContentType, StringComparison.OrdinalIgnoreCase) ||
            !PartBytes(sourceOwner).SequenceEqual(PartBytes(cloneOwner)))
            throw Mismatch(sourceRoot, $"part {PartPath(sourceOwner)} did not retain exact content bytes");
        AssertRelationshipShellEqual(sourceOwner, cloneOwner);
        AssertDataRelationshipsEqual(sourceOwner, cloneOwner, sourceRoot);
        var cloneChildren = cloneOwner.Parts.ToDictionary(pair => pair.RelationshipId, StringComparer.Ordinal);
        foreach (var sourcePair in sourceOwner.Parts)
        {
            if (!cloneChildren.TryGetValue(sourcePair.RelationshipId, out var clonePair) ||
                !sourcePair.OpenXmlPart.RelationshipType.Equals(clonePair.OpenXmlPart.RelationshipType, StringComparison.Ordinal))
                throw Mismatch(sourceRoot, $"relationship {sourcePair.RelationshipId} changed during graph copy");
            var sourceChild = sourcePair.OpenXmlPart;
            var cloneChild = clonePair.OpenXmlPart;
            if (SamePart(sourceChild, sourceRoot))
            {
                if (!SamePart(cloneChild, cloneRoot))
                    throw Mismatch(sourceRoot, $"back-reference {sourcePair.RelationshipId} does not target the clone SlidePart");
                continue;
            }
            if (ShouldShare(sourceChild))
            {
                if (sourceChild is SlidePart targetSlide && !retainedSlideParts.Contains(targetSlide))
                    throw Mismatch(sourceRoot, $"slide-jump relationship {sourcePair.RelationshipId} targets a removed slide");
                if (!SamePart(sourceChild, cloneChild))
                    throw Mismatch(sourceRoot, $"shared relationship {sourcePair.RelationshipId} was duplicated instead of rebound");
                continue;
            }
            if (sourceToClone.TryGetValue(sourceChild, out var mapped))
            {
                if (!SamePart(mapped, cloneChild))
                    throw Mismatch(sourceRoot, $"shared owned node {PartPath(sourceChild)} was copied more than once");
                continue;
            }
            if (SamePart(sourceChild, cloneChild))
                throw Mismatch(sourceRoot, $"mutable owned node {PartPath(sourceChild)} was shared with its clone");
            sourceToClone[sourceChild] = cloneChild;
            ValidateMappedGraph(sourceChild, cloneChild, sourceRoot, cloneRoot, retainedSlideParts, sourceToClone);
        }
    }

    private static bool ShouldShare(OpenXmlPart part) =>
        part is SlideLayoutPart or NotesMasterPart or ImagePart or SlidePart ||
        IsSharedReadOnlyOleObjectPart(part) ||
        // Some producers emit legacy/extended raster assets (for example WDP)
        // through an OpenXmlPart type that is not ImagePart.  The payload is
        // immutable presentation media, so sharing it is safe when its
        // content type is explicitly image/*; treating it as owned would
        // incorrectly reject an otherwise closed slide graph merely because
        // the same asset is used by another slide.
        part.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);

    // Some real producers bind the same legacy OLE binary to many slides. The
    // binary is not an editable Office package: OfficeKit exposes payload
    // replacement only for a uniquely inbound EmbeddedPackagePart. Sharing is
    // therefore safe only for a closed EmbeddedObjectPart leaf whose complete
    // inbound set consists of at least two SlideParts. A package part, a part
    // with descendants or external/data relationships, or a non-slide parent
    // remains mutable/ambiguous and continues to fail closed.
    private static bool IsSharedReadOnlyOleObjectPart(OpenXmlPart part)
    {
        if (part is not EmbeddedObjectPart || HasRelationships(part)) return false;
        var parents = part.GetParentParts().ToArray();
        return parents.Length >= 2 && parents.All(parent => parent is SlidePart);
    }

    // A component clone may copy a mutable package part that is shared by a
    // sibling which will be removed from the clone immediately afterwards.
    // Such a part is safe only when it is reachable exclusively through the
    // omitted shape-tree elements; a part also reachable from a retained
    // element remains blocked. The source package is never mutated by this
    // path, and the clone-side deletion proof removes the temporary copy.
    private static IReadOnlySet<OpenXmlPart> CloneOnlyParts(
        SlidePart source,
        IReadOnlySet<int>? omittedShapeTreeIndices)
    {
        if (omittedShapeTreeIndices is null || omittedShapeTreeIndices.Count == 0) return new HashSet<OpenXmlPart>();
        var shapeTree = source.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return new HashSet<OpenXmlPart>();
        var elements = ShapeTreeElements(shapeTree);
        var omittedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var retainedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < elements.Length; index++)
        {
            var relationshipIds = RelationshipIds(elements[index]);
            if (omittedShapeTreeIndices.Contains(index)) omittedRelationshipIds.UnionWith(relationshipIds);
            else retainedRelationshipIds.UnionWith(relationshipIds);
        }
        omittedRelationshipIds.ExceptWith(retainedRelationshipIds);
        var omittedRoots = source.Parts
            .Where(pair => omittedRelationshipIds.Contains(pair.RelationshipId))
            .Select(pair => pair.OpenXmlPart)
            .ToHashSet();
        if (omittedRoots.Count == 0) return new HashSet<OpenXmlPart>();
        var retainedRoots = source.Parts
            .Where(pair => !omittedRelationshipIds.Contains(pair.RelationshipId))
            .Select(pair => pair.OpenXmlPart)
            .ToHashSet();
        var omittedReachable = omittedRoots.SelectMany(ReachableParts).ToHashSet();
        var retainedReachable = retainedRoots.SelectMany(ReachableParts).ToHashSet();
        omittedReachable.ExceptWith(retainedReachable);
        return omittedReachable;
    }

    private static OpenXmlElement[] ShapeTreeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements
            .Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties)
            .ToArray();

    private static IReadOnlySet<string> RelationshipIds(OpenXmlElement source) =>
        new[] { source }
            .Concat(source.Descendants())
            .SelectMany(element => element.GetAttributes())
            .Where(attribute => attribute.NamespaceUri is
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships")
            .Select(attribute => attribute.Value)
            .Where(value => !string.IsNullOrEmpty(value))
            .Select(value => value!)
            .ToHashSet(StringComparer.Ordinal);

    private static IReadOnlySet<OpenXmlPart> ReachableParts(OpenXmlPart root)
    {
        var reachable = new HashSet<OpenXmlPart> { root };
        var queue = new Queue<OpenXmlPart>();
        queue.Enqueue(root);
        while (queue.TryDequeue(out var part))
        {
            foreach (var child in part.Parts.Select(pair => pair.OpenXmlPart))
                if (reachable.Add(child)) queue.Enqueue(child);
        }
        return reachable;
    }

    // Open XML SDK may materialize different wrapper instances for the same
    // package part after reopen. Package URI, not CLR object identity, is the
    // durable identity at this boundary.
    private static bool SamePart(OpenXmlPart left, OpenXmlPart right) =>
        left.Uri.Equals(right.Uri);

    private static void Rebind(OpenXmlPart owner, OpenXmlPart copied, OpenXmlPart target, string relationshipId)
    {
        if (!owner.DeletePart(relationshipId))
            throw new CodecException("presentation_slide_clone_graph_mismatch", $"Could not remove copied relationship {relationshipId} from {PartPath(owner)}.", PartPath(owner));
        owner.AddPart(target, relationshipId);
    }

    private static void AssertRelationshipShellEqual(OpenXmlPart source, OpenXmlPart clone)
    {
        var sourceExternal = source.ExternalRelationships
            .Select(item => $"{item.Id}\0{item.RelationshipType}\0{item.Uri}")
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var cloneExternal = clone.ExternalRelationships
            .Select(item => $"{item.Id}\0{item.RelationshipType}\0{item.Uri}")
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var sourceHyperlinks = source.HyperlinkRelationships
            .Select(item => $"{item.Id}\0{item.RelationshipType}\0{item.IsExternal}\0{item.Uri}")
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        var cloneHyperlinks = clone.HyperlinkRelationships
            .Select(item => $"{item.Id}\0{item.RelationshipType}\0{item.IsExternal}\0{item.Uri}")
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        if (!sourceExternal.SequenceEqual(cloneExternal, StringComparer.Ordinal) ||
            !sourceHyperlinks.SequenceEqual(cloneHyperlinks, StringComparer.Ordinal))
            throw new CodecException("presentation_slide_clone_graph_mismatch", $"External or hyperlink relationships changed while copying {PartPath(source)}.", PartPath(source));
    }

    private static void AssertDataRelationshipsEqual(OpenXmlPart source, OpenXmlPart clone, SlidePart sourceRoot)
    {
        var sourceRelationships = source.DataPartReferenceRelationships.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var cloneRelationships = clone.DataPartReferenceRelationships.ToDictionary(item => item.Id, StringComparer.Ordinal);
        if (!sourceRelationships.Keys.ToHashSet(StringComparer.Ordinal).SetEquals(cloneRelationships.Keys))
            throw Mismatch(sourceRoot, $"DataPart relationship inventory changed for {PartPath(source)}");
        foreach (var (relationshipId, sourceRelationship) in sourceRelationships)
        {
            var cloneRelationship = cloneRelationships[relationshipId];
            if (sourceRelationship.GetType() != cloneRelationship.GetType() ||
                !sourceRelationship.RelationshipType.Equals(cloneRelationship.RelationshipType, StringComparison.Ordinal) ||
                sourceRelationship.DataPart.Uri.Equals(cloneRelationship.DataPart.Uri) ||
                !sourceRelationship.DataPart.ContentType.Equals(cloneRelationship.DataPart.ContentType, StringComparison.OrdinalIgnoreCase) ||
                !DataPartBytes(sourceRelationship.DataPart).SequenceEqual(DataPartBytes(cloneRelationship.DataPart)))
                throw Mismatch(sourceRoot, $"DataPart relationship {relationshipId} changed during graph copy");
        }
    }

    private static IEnumerable<string> RelationshipIds(OpenXmlPart part) =>
        part.Parts.Select(pair => pair.RelationshipId)
            .Concat(part.ExternalRelationships.Select(item => item.Id))
            .Concat(part.HyperlinkRelationships.Select(item => item.Id))
            .Concat(part.DataPartReferenceRelationships.Select(item => item.Id))
            .Distinct(StringComparer.Ordinal);

    private static bool HasRelationships(OpenXmlPart part) =>
        part.Parts.Any() || part.ExternalRelationships.Any() ||
        part.HyperlinkRelationships.Any() || part.DataPartReferenceRelationships.Any();

    private static byte[] PartBytes(OpenXmlPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        return copy.ToArray();
    }

    private static byte[] DataPartBytes(DataPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        return copy.ToArray();
    }

    private static PptxSlideClonePlan Blocked(string reason) => new(false, reason, 0, 0);
    private static CodecException Mismatch(SlidePart source, string reason) =>
        new("presentation_slide_clone_graph_mismatch", reason, PartPath(source));
    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static string CloneRelationshipId(PresentationPart presentationPart, PptxSourceSlideEntry source)
    {
        var used = presentationPart.Parts
            .Select(pair => pair.RelationshipId)
            .ToHashSet(StringComparer.Ordinal);
        var slideOrdinal = presentationPart.Parts.Count(pair => pair.OpenXmlPart is SlidePart);
        var seed = $"{PartPath(source.Part)}\0{slideOrdinal}";
        var suffix = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(seed)))
            .ToLowerInvariant()[..16];
        var baseId = $"rIdOfficeKitClone{suffix}";
        var candidate = baseId;
        var collision = 0;
        while (!used.Add(candidate)) candidate = $"{baseId}_{++collision}";
        return candidate;
    }

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
