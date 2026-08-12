using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxElementDeletionPlan(bool Supported, string BlockedReason, uint NativeId);

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
        if (source is not P.Shape)
            return Blocked("only a top-level PresentationML shape is in the bounded deletion profile");

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
        if (HasRelationshipReference(source))
            return Blocked("the element owns one or more package relationship references", nativeId.Value);
        if (HasIdentitySensitiveSlideGraph(slide, common))
            return Blocked("slide timing or extension data may retain native element identity", nativeId.Value);
        if (PptxLegacyCommentsCodec.CommentPartPresent(slidePart))
            return Blocked("a slide comment graph may retain native element identity", nativeId.Value);
        if (siblings.OfType<P.ConnectionShape>().Any(connector => References(connector, nativeId.Value)))
            return Blocked($"connector topology references native drawing ID {nativeId}", nativeId.Value);

        return new PptxElementDeletionPlan(true, string.Empty, nativeId.Value);
    }

    internal static void Apply(OpenXmlElement source, PptxElementDeletionPlan plan)
    {
        if (!plan.Supported)
            throw new InvalidOperationException("Cannot apply an unsupported presentation element deletion plan.");
        source.Remove();
    }

    private static bool HasRelationshipReference(OpenXmlElement source) =>
        Attributes(source).Any(attribute =>
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

    private static PptxElementDeletionPlan Blocked(string reason, uint nativeId = 0) => new(false, reason, nativeId);
}
