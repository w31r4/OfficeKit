using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded p:cNvPr title/description contract. Other element codecs
// can reuse this native boundary without acquiring its parsing policy.
internal static class PptxNonVisualAccessibilityCodec
{
    private const int MaxTextLength = 1_024;

    internal static bool Supports(P.NonVisualDrawingProperties? source) => TryRead(source, out _);

    internal static PresentationNonVisualAccessibility? Read(P.NonVisualDrawingProperties? source) =>
        TryRead(source, out var value) ? value : null;

    internal static void Validate(PresentationNonVisualAccessibility? value, string elementId, string elementKind = "shape")
    {
        if (value is null) return;
        if (!value.HasTitle && !value.HasDescription)
            throw Invalid(elementId, elementKind, "must contain title and/or description");
        if (value.HasTitle && !IsValidValue(value.Title)) throw Invalid(elementId, elementKind, "title must contain 1 through 1024 XML-safe characters");
        if (value.HasDescription && !IsValidValue(value.Description)) throw Invalid(elementId, elementKind, "description must contain 1 through 1024 XML-safe characters");
    }

    internal static void ApplyAuthored(P.NonVisualDrawingProperties target, PresentationNonVisualAccessibility? value)
    {
        target.RemoveAttribute("title", string.Empty);
        target.RemoveAttribute("descr", string.Empty);
        if (value?.HasTitle == true) target.SetAttribute(new OpenXmlAttribute("title", string.Empty, value.Title));
        if (value?.HasDescription == true) target.SetAttribute(new OpenXmlAttribute("descr", string.Empty, value.Description));
    }

    internal static void ApplyBound(P.NonVisualDrawingProperties? source, PresentationNonVisualAccessibility? requested, string elementKind = "shape")
    {
        if (!TryRead(source, out _))
        {
            if (requested is null) return;
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} alternative text is not a canonical p:cNvPr profile.");
        }
        ApplyAuthored(source!, requested);
        if (!TryRead(source, out var actual) || !Equal(actual, requested))
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} alternative text did not round trip.");
    }

    internal static void ScrubModeledContent(P.NonVisualDrawingProperties? source)
    {
        if (!TryRead(source, out _)) return;
        source!.RemoveAttribute("title", string.Empty);
        source.RemoveAttribute("descr", string.Empty);
    }

    private static bool TryRead(P.NonVisualDrawingProperties? source, out PresentationNonVisualAccessibility? value)
    {
        value = null;
        if (source is null || source.HasChildren) return false;
        var hasId = false;
        var hasName = false;
        var hasHidden = false;
        var hasTitle = false;
        var hasDescription = false;
        string? title = null;
        string? description = null;

        foreach (var attribute in source.GetAttributes())
        {
            if (attribute.NamespaceUri.Length != 0) return false;
            switch (attribute.LocalName)
            {
                case "id" when !hasId && uint.TryParse(attribute.Value, out var id) && id != 0:
                    hasId = true;
                    break;
                case "name" when !hasName && IsValidName(attribute.Value):
                    hasName = true;
                    break;
                case "hidden" when !hasHidden && attribute.Value is "0" or "1" or "false" or "true":
                    hasHidden = true;
                    break;
                case "title" when !hasTitle && IsValidValue(attribute.Value):
                    hasTitle = true;
                    title = attribute.Value;
                    break;
                case "descr" when !hasDescription && IsValidValue(attribute.Value):
                    hasDescription = true;
                    description = attribute.Value;
                    break;
                default:
                    return false;
            }
        }
        if (!hasId || !hasName) return false;
        if (!hasTitle && !hasDescription) return true;
        value = new PresentationNonVisualAccessibility();
        if (hasTitle) value.Title = title!;
        if (hasDescription) value.Description = description!;
        return true;
    }

    private static bool IsValidName(string? value) =>
        value is not null && value.Length <= MaxTextLength && !value.Contains('\u007f') && IsXmlSafe(value);

    private static bool IsValidValue(string? value) =>
        !string.IsNullOrEmpty(value) && value.Length <= MaxTextLength && !value.Contains('\u007f') && IsXmlSafe(value);

    private static bool IsXmlSafe(string value)
    {
        try { XmlConvert.VerifyXmlChars(value); return true; }
        catch (XmlException) { return false; }
    }

    private static bool Equal(PresentationNonVisualAccessibility? left, PresentationNonVisualAccessibility? right) =>
        left?.HasTitle == right?.HasTitle && left?.Title == right?.Title &&
        left?.HasDescription == right?.HasDescription && left?.Description == right?.Description;

    private static CodecException Invalid(string elementId, string elementKind, string message) =>
        new($"invalid_presentation_{elementKind}", $"Presentation {elementKind} {elementId} accessibility {message}.");
}
