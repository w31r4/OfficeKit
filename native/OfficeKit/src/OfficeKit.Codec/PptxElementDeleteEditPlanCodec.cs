using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Lowers one already-issued top-level deletion capability into token-local
// SlidePart and relationship patches. The existing deletion analyzer remains
// the authority for native identity and relationship-closure ownership.
internal static partial class PptxEditPlanCodec
{
    private static PptxElementDeletionPlan ProveElementDeletion(
        SlidePart slidePart,
        P.ShapeTree tree,
        OpenXmlElement element,
        PresentationElement projectedElement,
        PresentationEditOperation operation)
    {
        var expectedNativeId = operation.ElementDeletion?.ExpectedNativeId ?? 0;
        var issued = projectedElement.Source?.DeletionCapability;
        if (issued?.Supported != true || issued.NativeId == 0 || issued.NativeId != expectedNativeId)
            throw new CodecException("unsupported_presentation_element_delete", $"PPTX element deletion {operation.OperationId} has no current codec-issued deletion capability.", operation.SlidePartPath);
        var plan = PptxElementDeletionCodec.Analyze(slidePart, element, tree.ChildElements.ToArray());
        if (!plan.Supported || plan.NativeId != expectedNativeId)
            throw new CodecException("unsupported_presentation_element_delete", $"PPTX element deletion {operation.OperationId} no longer satisfies its source-graph proof{(plan.BlockedReason.Length == 0 ? string.Empty : $": {plan.BlockedReason}")}.", operation.SlidePartPath);
        return plan;
    }

    private static void ApplyElementDeletionPackagePatches(
        IReadOnlyDictionary<string, byte[]> sourceParts,
        IReadOnlyList<PptxEditPlanProof> proofs,
        IDictionary<string, byte[]> patchedParts,
        ISet<string> removedParts)
    {
        var deletions = proofs.Where(proof => proof.Deletion is not null).ToArray();
        foreach (var group in deletions
                     .Where(proof => proof.Deletion!.RelationshipIds.Count > 0)
                     .GroupBy(proof => RelationshipPartPathFor(proof.Operation.SlidePartPath), StringComparer.OrdinalIgnoreCase))
        {
            var path = group.Key;
            if (patchedParts.ContainsKey(path))
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX element deletion overlaps another mutation of {path}.", path);
            if (!sourceParts.TryGetValue(path, out var sourceRelationships))
                throw new CodecException("presentation_edit_target_missing", $"PPTX element deletion relationship part {path} is missing.", path);
            var relationshipIds = group.SelectMany(proof => proof.Deletion!.RelationshipIds).ToArray();
            if (relationshipIds.Distinct(StringComparer.Ordinal).Count() != relationshipIds.Length)
                throw new CodecException("presentation_edit_plan_scope_violation", $"PPTX element deletions overlap relationship ownership in {path}.", path);
            patchedParts.Add(path, RemoveElementRelationships(sourceRelationships, relationshipIds, path));
        }
        foreach (var path in deletions.SelectMany(proof => proof.Deletion!.RemovedPackagePartPaths).Distinct(StringComparer.OrdinalIgnoreCase))
            if (sourceParts.ContainsKey(path)) removedParts.Add(path);
    }

    private static byte[] RemoveElementRelationships(byte[] sourcePart, IReadOnlyList<string> relationshipIds, string partPath)
    {
        var (xml, bomBytes) = DecodeXml(sourcePart);
        var tokens = XmlTokenPattern().Matches(xml).Cast<Match>().ToArray();
        var ranges = relationshipIds.Select(relationshipId => RelationshipToken(tokens, relationshipId, partPath))
            .Select(token => (token.Index, End: token.Index + token.Length))
            .OrderByDescending(range => range.Index)
            .ToArray();
        var output = xml;
        foreach (var range in ranges) output = output[..range.Index] + output[range.End..];
        var encoded = StrictUtf8.GetBytes(output);
        return bomBytes == 0 ? encoded : StrictUtf8.GetPreamble().Concat(encoded).ToArray();
    }
}
