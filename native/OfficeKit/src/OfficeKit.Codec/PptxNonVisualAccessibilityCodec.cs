using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using AD = DocumentFormat.OpenXml.Office2019.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded p:cNvPr title/description/decorative contract. Other
// element codecs can reuse this native boundary without acquiring its parsing
// or extension-list policy.
internal static class PptxNonVisualAccessibilityCodec
{
    private const int MaxTextLength = 1_024;
    private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

    internal static bool Supports(P.NonVisualDrawingProperties? source) => TryRead(source, out _);

    // Pictures already have a residual-protected cNvPr profile. Unknown
    // children stay byte-owned by that profile; only a canonical decorative
    // extension, when present, becomes modeled state.
    internal static bool SupportsResidual(P.NonVisualDrawingProperties? source) => TryReadResidual(source, out _);

    internal static PresentationNonVisualAccessibility? Read(P.NonVisualDrawingProperties? source) =>
        TryRead(source, out var value) ? value : null;

    internal static bool TryReadResidual(P.NonVisualDrawingProperties? source, out PresentationNonVisualAccessibility? value)
    {
        value = null;
        if (source is null || !TryReadDecorative(source, allowOtherChildren: true, out var decorative)) return false;
        if (!TryReadAttributeValue(source, "title", out var title) ||
            !TryReadAttributeValue(source, "descr", out var description)) return false;
        if (decorative == true && (title is not null || description is not null)) return false;
        value = Value(title, description, decorative);
        return true;
    }

    internal static void Validate(PresentationNonVisualAccessibility? value, string elementId, string elementKind = "shape")
    {
        if (value is null) return;
        if (!value.HasTitle && !value.HasDescription && !value.HasDecorative)
            throw Invalid(elementId, elementKind, "must contain title, description, and/or decorative");
        if (value.HasTitle && !IsValidValue(value.Title)) throw Invalid(elementId, elementKind, "title must contain 1 through 1024 XML-safe characters");
        if (value.HasDescription && !IsValidValue(value.Description)) throw Invalid(elementId, elementKind, "description must contain 1 through 1024 XML-safe characters");
        if (value.HasDecorative && value.Decorative && (value.HasTitle || value.HasDescription))
            throw Invalid(elementId, elementKind, "cannot combine decorative true with title or description");
    }

    internal static void ApplyAuthored(P.NonVisualDrawingProperties target, PresentationNonVisualAccessibility? value)
    {
        ApplyText(target, value);
        ApplyDecorative(target, value?.HasDecorative == true ? value.Decorative : null, allowOtherChildren: false);
    }

    internal static void ApplyBound(P.NonVisualDrawingProperties? source, PresentationNonVisualAccessibility? requested, string elementKind = "shape")
    {
        if (!TryRead(source, out _))
        {
            if (requested is null) return;
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} accessibility metadata is not a canonical p:cNvPr profile.");
        }
        ApplyAuthored(source!, requested);
        if (!TryRead(source, out var actual) || !Equal(actual, requested))
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} accessibility metadata did not round trip.");
    }

    internal static void ApplyResidualBound(P.NonVisualDrawingProperties? source, PresentationNonVisualAccessibility? requested, string elementKind)
    {
        if (!TryReadResidual(source, out _))
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} accessibility metadata has an ambiguous decorative extension graph.");
        ApplyText(source!, requested);
        ApplyDecorative(source!, requested?.HasDecorative == true ? requested.Decorative : null, allowOtherChildren: true);
        if (!TryReadResidual(source, out var actual) || !Equal(actual, requested))
            throw new CodecException("unsupported_presentation_edit", $"Source {elementKind} accessibility metadata did not round trip.");
    }

    internal static void ScrubModeledContent(P.NonVisualDrawingProperties? source)
    {
        if (!TryRead(source, out _)) return;
        source!.RemoveAttribute("title", string.Empty);
        source.RemoveAttribute("descr", string.Empty);
        ApplyDecorative(source, null, allowOtherChildren: false);
    }

    internal static void ScrubResidualModeledContent(P.NonVisualDrawingProperties? source)
    {
        if (!TryReadResidual(source, out _)) return;
        source!.RemoveAttribute("title", string.Empty);
        source.RemoveAttribute("descr", string.Empty);
        ApplyDecorative(source, null, allowOtherChildren: true);
    }

    private static bool TryRead(P.NonVisualDrawingProperties? source, out PresentationNonVisualAccessibility? value)
    {
        value = null;
        if (source is null || !TryReadDecorative(source, allowOtherChildren: false, out var decorative)) return false;
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
        if (decorative == true && (title is not null || description is not null)) return false;
        value = Value(title, description, decorative);
        return true;
    }

    private static PresentationNonVisualAccessibility? Value(string? title, string? description, bool? decorative)
    {
        if (title is null && description is null && decorative is null) return null;
        var value = new PresentationNonVisualAccessibility();
        if (title is not null) value.Title = title;
        if (description is not null) value.Description = description;
        if (decorative is not null) value.Decorative = decorative.Value;
        return value;
    }

    private static bool TryReadAttributeValue(P.NonVisualDrawingProperties source, string localName, out string? value)
    {
        value = null;
        var matches = source.GetAttributes()
            .Where(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == localName)
            .ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1 || !IsValidValue(matches[0].Value)) return false;
        value = matches[0].Value;
        return true;
    }

    private static bool TryReadDecorative(P.NonVisualDrawingProperties source, bool allowOtherChildren, out bool? decorative)
    {
        decorative = null;
        var lists = source.Elements<A.NonVisualDrawingPropertiesExtensionList>().ToArray();
        if (lists.Length > 1 ||
            (!allowOtherChildren && source.ChildElements.Any(child => child is not A.NonVisualDrawingPropertiesExtensionList))) return false;
        if (lists.Length == 0) return allowOtherChildren || !source.HasChildren;
        var list = lists[0];
        if (list.HasAttributes || list.ChildElements.Any(child => child is not A.NonVisualDrawingPropertiesExtension)) return false;
        var extensions = list.Elements<A.NonVisualDrawingPropertiesExtension>().ToArray();
        var modeled = extensions.Where(extension => extension.Uri?.Value == DecorativeExtensionUri).ToArray();
        if (modeled.Length > 1 || (!allowOtherChildren && extensions.Length != modeled.Length)) return false;
        if (modeled.Length == 0) return allowOtherChildren;
        var extension = modeled[0];
        if (extension.GetAttributes().Count != 1 || extension.ChildElements.Count != 1 ||
            extension.GetFirstChild<AD.Decorative>() is not { } native || native.HasChildren ||
            native.Val is null || native.GetAttributes().Count != 1) return false;
        decorative = native.Val.Value;
        return true;
    }

    private static void ApplyText(P.NonVisualDrawingProperties target, PresentationNonVisualAccessibility? value)
    {
        target.RemoveAttribute("title", string.Empty);
        target.RemoveAttribute("descr", string.Empty);
        if (value?.HasTitle == true) target.SetAttribute(new OpenXmlAttribute("title", string.Empty, value.Title));
        if (value?.HasDescription == true) target.SetAttribute(new OpenXmlAttribute("descr", string.Empty, value.Description));
    }

    private static void ApplyDecorative(P.NonVisualDrawingProperties target, bool? requested, bool allowOtherChildren)
    {
        if (!TryReadDecorative(target, allowOtherChildren, out _))
            throw new CodecException("unsupported_presentation_edit", "Presentation accessibility metadata has an ambiguous decorative extension graph.");
        var list = target.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
        var extension = list?.Elements<A.NonVisualDrawingPropertiesExtension>()
            .SingleOrDefault(item => item.Uri?.Value == DecorativeExtensionUri);
        if (requested is null)
        {
            extension?.Remove();
            if (list is not null && !list.HasChildren) list.Remove();
            return;
        }
        if (list is null)
        {
            list = new A.NonVisualDrawingPropertiesExtensionList();
            target.Append(list);
        }
        if (extension is null)
        {
            extension = new A.NonVisualDrawingPropertiesExtension { Uri = DecorativeExtensionUri };
            list.Append(extension);
        }
        extension.RemoveAllChildren();
        extension.Append(new AD.Decorative { Val = requested.Value });
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
        left?.HasDescription == right?.HasDescription && left?.Description == right?.Description &&
        left?.HasDecorative == right?.HasDecorative && left?.Decorative == right?.Decorative;

    private static CodecException Invalid(string elementId, string elementKind, string message) =>
        new($"invalid_presentation_{elementKind}", $"Presentation {elementKind} {elementId} accessibility {message}.");
}
