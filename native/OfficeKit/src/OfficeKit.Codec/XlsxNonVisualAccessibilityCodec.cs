using System.Xml;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using AD = DocumentFormat.OpenXml.Office2019.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeKit.Codec;

// Owns the residual-protected xdr:cNvPr title/description/decorative leaf.
// ChartSpace and picture geometry keep independent capability decisions.
internal static class XlsxNonVisualAccessibilityCodec
{
    private const int MaxTextLength = 1_024;
    private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

    internal sealed record Value(string? Title, string? Description, bool? Decorative);

    internal static bool TryRead(Xdr.NonVisualDrawingProperties? source, out Value? value)
    {
        value = null;
        if (source is null || !TryReadDecorative(source, out var decorative) ||
            !TryReadAttributeValue(source, "title", out var title) ||
            !TryReadAttributeValue(source, "descr", out var description)) return false;
        if (decorative == true && (title is not null || description is not null)) return false;
        value = title is null && description is null && decorative is null ? null : new Value(title, description, decorative);
        return true;
    }

    internal static void Validate(string? title, string? description, bool? decorative, string worksheetId, string elementId, string kind)
    {
        if (title is not null && !IsValidValue(title)) throw Invalid(worksheetId, elementId, kind, "title must contain 1 through 1024 XML-safe characters");
        if (description is not null && !IsValidValue(description)) throw Invalid(worksheetId, elementId, kind, "description must contain 1 through 1024 XML-safe characters");
        if (decorative == true && (title is not null || description is not null)) throw Invalid(worksheetId, elementId, kind, "cannot combine decorative true with title or description");
    }

    internal static void ApplyAuthored(Xdr.NonVisualDrawingProperties target, string? title, string? description, bool? decorative)
    {
        ApplyText(target, title, description);
        ApplyDecorative(target, decorative);
    }

    internal static void ApplyBound(Xdr.NonVisualDrawingProperties? target, string? title, string? description, bool? decorative, string elementId, string kind)
    {
        if (!TryRead(target, out _)) throw new CodecException($"unsupported_spreadsheet_{kind}_edit", $"Worksheet {kind} {elementId} accessibility metadata has an ambiguous xdr:cNvPr extension graph.");
        ApplyAuthored(target!, title, description, decorative);
        if (!TryRead(target, out var actual) || !Equal(actual, title, description, decorative))
            throw new CodecException($"unsupported_spreadsheet_{kind}_edit", $"Worksheet {kind} {elementId} accessibility metadata did not round trip.");
    }

    internal static bool Equal(Value? value, string? title, string? description, bool? decorative) =>
        value?.Title == title && value?.Description == description && value?.Decorative == decorative;

    internal static string Semantics(string? title, string? description, bool? decorative) =>
        string.Join('\0', title is null ? "absent" : $"title:{title}", description is null ? "absent" : $"description:{description}", decorative is null ? "absent" : decorative.Value ? "decorative:true" : "decorative:false");

    private static bool TryReadAttributeValue(Xdr.NonVisualDrawingProperties source, string localName, out string? value)
    {
        value = null;
        var matches = source.GetAttributes().Where(attribute => attribute.NamespaceUri.Length == 0 && attribute.LocalName == localName).ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1 || !IsValidValue(matches[0].Value)) return false;
        value = matches[0].Value;
        return true;
    }

    private static bool TryReadDecorative(Xdr.NonVisualDrawingProperties source, out bool? decorative)
    {
        decorative = null;
        var lists = source.Elements<A.NonVisualDrawingPropertiesExtensionList>().ToArray();
        if (lists.Length > 1) return false;
        if (lists.Length == 0) return true;
        var list = lists[0];
        if (list.HasAttributes || list.ChildElements.Any(child => child is not A.NonVisualDrawingPropertiesExtension)) return false;
        var modeled = list.Elements<A.NonVisualDrawingPropertiesExtension>().Where(extension => extension.Uri?.Value == DecorativeExtensionUri).ToArray();
        if (modeled.Length > 1) return false;
        if (modeled.Length == 0) return true;
        var extension = modeled[0];
        if (extension.GetAttributes().Count != 1 || extension.ChildElements.Count != 1 ||
            extension.GetFirstChild<AD.Decorative>() is not { } native || native.HasChildren || native.Val is null || native.GetAttributes().Count != 1) return false;
        decorative = native.Val.Value;
        return true;
    }

    private static void ApplyText(Xdr.NonVisualDrawingProperties target, string? title, string? description)
    {
        target.RemoveAttribute("title", string.Empty);
        target.RemoveAttribute("descr", string.Empty);
        if (title is not null) target.SetAttribute(new OpenXmlAttribute("title", string.Empty, title));
        if (description is not null) target.SetAttribute(new OpenXmlAttribute("descr", string.Empty, description));
    }

    private static void ApplyDecorative(Xdr.NonVisualDrawingProperties target, bool? requested)
    {
        if (!TryReadDecorative(target, out _)) throw new CodecException("unsupported_spreadsheet_accessibility_edit", "Worksheet drawing accessibility metadata has an ambiguous xdr:cNvPr extension graph.");
        var list = target.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
        var extension = list?.Elements<A.NonVisualDrawingPropertiesExtension>().SingleOrDefault(item => item.Uri?.Value == DecorativeExtensionUri);
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

    private static bool IsValidValue(string? value) => !string.IsNullOrEmpty(value) && value.Length <= MaxTextLength && !value.Contains('\u007f') && IsXmlSafe(value);
    private static bool IsXmlSafe(string value) { try { XmlConvert.VerifyXmlChars(value); return true; } catch (XmlException) { return false; } }
    private static CodecException Invalid(string worksheetId, string elementId, string kind, string message) =>
        new($"invalid_spreadsheet_{kind}", $"Worksheet {worksheetId} {kind} {elementId} accessibility {message}.");
}
