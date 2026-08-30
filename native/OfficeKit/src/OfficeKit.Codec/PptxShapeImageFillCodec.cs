using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the deliberately narrow image-filled-shape projection.  A shape-level
// blipFill is not treated as a general image authoring surface: only the
// common direct embedded-image profile is projected, and its asset remains
// bound to the source relationship.  Placement and already-modeled custom
// geometry may then be edited without reconstructing the native fill.
internal static class PptxShapeImageFillCodec
{
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictOfficeRelationshipsNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    internal static bool TryRead(A.BlipFill? source, PptxPartContext? context, out Asset asset)
    {
        asset = null!;
        if (source is null || context is null || context.Assets is null ||
            !HasOnlyUnqualifiedAttributes(source, "rotWithShape") || source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Blip blip || source.ChildElements[1] is not A.Stretch stretch ||
            !HasOnlyUnqualifiedAttributes(stretch) || stretch.ChildElements.Count != 1 ||
            stretch.FirstChild is not A.FillRectangle fillRectangle ||
            !HasZeroFillRectangle(fillRectangle) ||
            blip.Link is not null || blip.Embed?.Value is not { Length: > 0 } embed ||
            blip.CompressionState is not null || !HasOnlyEmbedAttribute(blip) ||
            blip.ChildElements.Count > 1 ||
            blip.ChildElements.Any(child => child is not A.AlphaModulationFixed ||
                child.GetAttributes().Count != 0 || child.ChildElements.Count != 0))
            return false;

        try
        {
            asset = context.ReadEmbeddedPicture(embed);
            return true;
        }
        catch (CodecException)
        {
            asset = null!;
            return false;
        }
    }

    private static bool HasOnlyEmbedAttribute(A.Blip blip)
    {
        var attributes = blip.GetAttributes();
        if (attributes.Count != 1) return false;
        var attribute = attributes[0];
        return attribute.LocalName == "embed" &&
               attribute.NamespaceUri is OfficeRelationshipsNamespace or StrictOfficeRelationshipsNamespace;
    }

    private static bool HasZeroFillRectangle(A.FillRectangle fillRectangle)
    {
        var attributes = fillRectangle.GetAttributes();
        if (attributes.Count != 4) return false;
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
        {
            if (attribute.NamespaceUri.Length != 0 || attribute.Value != "0" ||
                attribute.LocalName is not ("b" or "l" or "r" or "t") ||
                !names.Add(attribute.LocalName)) return false;
        }
        return names.Count == 4;
    }

    private static bool HasOnlyUnqualifiedAttributes(OpenXmlElement element, params string[] allowedNames)
    {
        var allowed = allowedNames.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute =>
            attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
