using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Shared direct value contract. Effect-list composition belongs to each owner.
internal static class PptxSoftEdgeValueCodec
{
    private const long MaxRadiusEmu = 12_700_000L;

    internal static void Validate(PresentationSoftEdge? softEdge, string elementId, string subject = "shape")
    {
        if (softEdge is null) return;
        if (!softEdge.HasRadiusEmu || softEdge.RadiusEmu is < 0 or > MaxRadiusEmu)
            throw new CodecException(
                "invalid_presentation_soft_edge",
                $"Presentation {subject} {elementId} has invalid soft-edge radius.");
    }

    internal static bool TryRead(A.SoftEdge source, out PresentationSoftEdge? softEdge)
    {
        softEdge = null;
        if (!HasOnlyAttributes(source, "rad") || source.ChildElements.Count != 0 ||
            source.Radius?.Value is not { } radius || radius > MaxRadiusEmu)
            return false;
        softEdge = new PresentationSoftEdge { RadiusEmu = radius };
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
