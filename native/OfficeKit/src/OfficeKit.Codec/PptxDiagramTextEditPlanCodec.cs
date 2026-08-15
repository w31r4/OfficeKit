using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static partial class PptxEditPlanCodec
{
    private static PptxXmlPatch[] CompileDiagramTextXmlPatches(
        byte[] partBytes,
        IReadOnlyList<PptxEditPlanProof> proofs)
    {
        var (xml, _) = DecodeXml(partBytes);
        var drawingPrefixes = NamespacePrefixes(xml, DrawingNamespace);
        if (drawingPrefixes.Count == 0)
            throw new CodecException("presentation_edit_target_missing", "PPTX DiagramDataPart does not declare the DrawingML namespace.", proofs[0].MutationPartPath);
        var textRanges = NamespaceElementRanges(xml, "t", drawingPrefixes);
        var patches = new List<PptxXmlPatch>();
        foreach (var proof in proofs)
        {
            var operation = proof.Operation;
            if (proof.RawTextOrdinal is not uint rawTextOrdinal || rawTextOrdinal >= (uint)textRanges.Count)
                throw new CodecException("presentation_edit_target_missing", $"PPTX edit operation {operation.OperationId} raw SmartArt text ordinal is out of range.", proof.MutationPartPath);
            if (NeedsPreserve(operation.Value))
                throw new CodecException("invalid_presentation_edit_operation", $"PPTX edit operation {operation.OperationId} diagramText v1 cannot introduce leading or trailing whitespace.", proof.MutationPartPath);
            var range = textRanges[(int)rawTextOrdinal];
            var (start, end) = ElementTextSpan(xml, range);
            if (DecodePlainXmlText(xml[start..end]) != operation.ExpectedValue)
                throw new CodecException("presentation_leaf_precondition_failed", $"PPTX edit operation {operation.OperationId} raw SmartArt text does not match the expected value.", proof.MutationPartPath);
            patches.Add(new PptxXmlPatch(
                operation,
                start,
                end,
                EscapeText(operation.Value),
                proof.SourceElementSha256,
                proof.MutationPartPath));
        }
        return OrderedNonOverlapping(patches, proofs[0].MutationPartPath);
    }

    private static string DecodePlainXmlText(string value)
    {
        try
        {
            return XElement.Parse($"<root>{value}</root>", LoadOptions.PreserveWhitespace).Value;
        }
        catch (System.Xml.XmlException exception)
        {
            throw new CodecException("invalid_presentation_text_leaf", "PPTX SmartArt text leaf is not safe XML.", innerException: exception);
        }
    }
}
