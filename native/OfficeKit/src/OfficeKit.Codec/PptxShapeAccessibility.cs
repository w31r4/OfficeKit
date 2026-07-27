using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the narrow p:sp alternative-text profile. PowerPoint stores this
// metadata on p:nvSpPr/p:cNvPr rather than in the visible text frame. It is
// deliberately separate from PptxTextCodec and shape formatting so an
// accessibility-only edit cannot rewrite the text body, geometry, fill, or
// relationship graph.
internal static class PptxShapeAccessibility
{
    private const int MaxTextLength = 1_024;

    internal static void Validate(PresentationShape shape, string elementId)
    {
        ValidateValue(shape.HasAccessibilityTitle, shape.AccessibilityTitle, elementId, "accessibility_title");
        ValidateValue(shape.HasAccessibilityDescription, shape.AccessibilityDescription, elementId, "accessibility_description");
    }

    internal static bool Read(P.Shape source, PresentationShape artifact)
    {
        artifact.ClearAccessibilityTitle();
        artifact.ClearAccessibilityDescription();
        if (!TryRead(source, out var title, out var description)) return false;
        if (title is not null) artifact.AccessibilityTitle = title;
        if (description is not null) artifact.AccessibilityDescription = description;
        return true;
    }

    internal static bool Supports(P.Shape source) => TryRead(source, out _, out _);

    internal static void ApplyAuthored(P.NonVisualDrawingProperties nonVisual, PresentationShape requested)
    {
        Write(nonVisual,
            requested.HasAccessibilityTitle ? requested.AccessibilityTitle : null,
            requested.HasAccessibilityDescription ? requested.AccessibilityDescription : null);
    }

    internal static void Apply(P.Shape source, PresentationShape requested)
    {
        if (!TryRead(source, out _, out _))
            throw Unsupported("Source-preserving PPTX shape alternative text requires one child-free p:cNvPr with only canonical id/name/title/descr/hidden attributes and non-empty XML-safe title/description values.");
        var nonVisual = source.NonVisualShapeProperties?.NonVisualDrawingProperties ??
            throw Unsupported("Source-preserving PPTX shape alternative text cannot create a missing p:cNvPr container.");
        Write(nonVisual,
            requested.HasAccessibilityTitle ? requested.AccessibilityTitle : null,
            requested.HasAccessibilityDescription ? requested.AccessibilityDescription : null);
        if (!TryRead(source, out var title, out var description) || !Same(title, description, requested))
            throw Unsupported("Source-preserving PPTX shape alternative text did not round trip through the bounded profile.");
    }

    internal static void MaskModeled(P.Shape source)
    {
        if (!TryRead(source, out _, out _)) return;
        var nonVisual = source.NonVisualShapeProperties?.NonVisualDrawingProperties;
        if (nonVisual is null) return;
        Write(nonVisual, null, null);
    }

    private static bool TryRead(P.Shape source, out string? title, out string? description)
    {
        title = null;
        description = null;
        var nonVisual = source.NonVisualShapeProperties?.NonVisualDrawingProperties;
        if (nonVisual is null || nonVisual.ChildElements.Count != 0) return false;
        var attributes = nonVisual.GetAttributes();
        if (attributes.Any(attribute => attribute.NamespaceUri.Length != 0 ||
            attribute.LocalName is not ("id" or "name" or "title" or "descr" or "hidden"))) return false;
        var ids = attributes.Where(attribute => attribute.LocalName == "id").ToArray();
        if (ids.Length != 1 || !uint.TryParse(ids[0].Value, out var nativeId) || nativeId == 0) return false;
        if (!TryReadRequiredName(attributes.Where(attribute => attribute.LocalName == "name").ToArray())) return false;
        if (!TryReadHidden(attributes.Where(attribute => attribute.LocalName == "hidden").ToArray())) return false;
        if (!TryReadValue(attributes.Where(attribute => attribute.LocalName == "title").ToArray(), out title) ||
            !TryReadValue(attributes.Where(attribute => attribute.LocalName == "descr").ToArray(), out description)) return false;
        return true;
    }

    private static bool TryReadRequiredName(IReadOnlyList<OpenXmlAttribute> attributes)
    {
        if (attributes.Count != 1) return false;
        var value = attributes[0].Value;
        if (string.IsNullOrEmpty(value) || value.Contains('\u007f')) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool TryReadHidden(IReadOnlyList<OpenXmlAttribute> attributes) =>
        attributes.Count == 0 || (attributes.Count == 1 && attributes[0].Value is "0" or "1" or "false" or "true");

    private static bool TryReadValue(IReadOnlyList<OpenXmlAttribute> attributes, out string? value)
    {
        value = null;
        if (attributes.Count == 0) return true;
        if (attributes.Count != 1) return false;
        value = attributes[0].Value;
        return IsValidValue(value);
    }

    private static bool Same(string? title, string? description, PresentationShape requested) =>
        (title is not null) == requested.HasAccessibilityTitle &&
        (title is null || string.Equals(title, requested.AccessibilityTitle, StringComparison.Ordinal)) &&
        (description is not null) == requested.HasAccessibilityDescription &&
        (description is null || string.Equals(description, requested.AccessibilityDescription, StringComparison.Ordinal));

    private static void Write(P.NonVisualDrawingProperties nonVisual, string? title, string? description)
    {
        nonVisual.RemoveAttribute("title", string.Empty);
        nonVisual.RemoveAttribute("descr", string.Empty);
        if (title is not null) nonVisual.SetAttribute(new OpenXmlAttribute("title", string.Empty, title));
        if (description is not null) nonVisual.SetAttribute(new OpenXmlAttribute("descr", string.Empty, description));
    }

    private static void ValidateValue(bool present, string value, string elementId, string field)
    {
        if (present && !IsValidValue(value))
            throw new CodecException("invalid_presentation_shape", $"Presentation shape {elementId} {field} must contain 1 through {MaxTextLength} XML-safe characters.");
    }

    private static bool IsValidValue(string? value)
    {
        if (string.IsNullOrEmpty(value) || value.Length > MaxTextLength || value.Contains('\u007f')) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
