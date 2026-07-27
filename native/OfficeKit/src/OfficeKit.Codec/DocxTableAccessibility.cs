using System.Xml;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeKit.Codec;

// Owns the narrow, non-visible table alternative-text profile. Word stores a
// table title and description in tblPr rather than in visible caption
// paragraphs. Keeping those leaves separate from DocxTableFormatting makes
// their accessibility semantics independently source-bound and prevents a
// visual-formatting edit from rewriting them.
internal static class DocxTableAccessibility
{
    private const int MaxTextLength = 32_767;
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static void Validate(DocumentTable table)
    {
        ValidateValue(table.HasAccessibilityTitle, table.AccessibilityTitle, "accessibility_title");
        ValidateValue(table.HasAccessibilityDescription, table.AccessibilityDescription, "accessibility_description");
    }

    internal static bool Read(W.Table table, DocumentTable artifact)
    {
        artifact.ClearAccessibilityTitle();
        artifact.ClearAccessibilityDescription();
        if (!TryRead(table, out var title, out var description)) return false;
        if (title is not null) artifact.AccessibilityTitle = title;
        if (description is not null) artifact.AccessibilityDescription = description;
        return true;
    }

    internal static bool Same(DocumentTable left, DocumentTable right) =>
        left.HasAccessibilityTitle == right.HasAccessibilityTitle &&
        (!left.HasAccessibilityTitle || string.Equals(left.AccessibilityTitle, right.AccessibilityTitle, StringComparison.Ordinal)) &&
        left.HasAccessibilityDescription == right.HasAccessibilityDescription &&
        (!left.HasAccessibilityDescription || string.Equals(left.AccessibilityDescription, right.AccessibilityDescription, StringComparison.Ordinal));

    internal static void AppendAuthored(W.TableProperties properties, DocumentTable table) =>
        Write(properties, Value(table.HasAccessibilityTitle, table.AccessibilityTitle), Value(table.HasAccessibilityDescription, table.AccessibilityDescription));

    internal static void Apply(W.Table table, DocumentTable requested)
    {
        if (!TryRead(table, out _, out _))
            throw Unsupported("Source-preserving DOCX table accessibility metadata requires zero or one canonical w:tblCaption and w:tblDescription leaf with only a non-empty w:val attribute.");
        var properties = table.TableProperties;
        if (properties is null)
            throw Unsupported("Source-preserving DOCX table accessibility metadata cannot create a missing w:tblPr container.");
        Write(properties, Value(requested.HasAccessibilityTitle, requested.AccessibilityTitle), Value(requested.HasAccessibilityDescription, requested.AccessibilityDescription));
        if (!TryRead(table, out var title, out var description) ||
            !Same(title, description, requested))
            throw Unsupported("Source-preserving DOCX table accessibility metadata did not round trip through the bounded profile.");
    }

    internal static void MaskModeled(W.Table table)
    {
        if (!TryRead(table, out _, out _)) return;
        var properties = table.TableProperties;
        if (properties is null) return;
        foreach (var caption in properties.Elements<W.TableCaption>().ToArray()) caption.Remove();
        foreach (var description in properties.Elements<W.TableDescription>().ToArray()) description.Remove();
    }

    private static bool TryRead(W.Table table, out string? title, out string? description)
    {
        title = null;
        description = null;
        var properties = table.TableProperties;
        if (properties is null) return true;
        var captions = properties.Elements<W.TableCaption>().ToArray();
        var descriptions = properties.Elements<W.TableDescription>().ToArray();
        if (captions.Length > 1 || descriptions.Length > 1 ||
            !TryReadValue(captions.SingleOrDefault(), out title) ||
            !TryReadValue(descriptions.SingleOrDefault(), out description))
            return false;
        return true;
    }

    private static bool Same(string? title, string? description, DocumentTable requested) =>
        (title is not null) == requested.HasAccessibilityTitle &&
        (title is null || string.Equals(title, requested.AccessibilityTitle, StringComparison.Ordinal)) &&
        (description is not null) == requested.HasAccessibilityDescription &&
        (description is null || string.Equals(description, requested.AccessibilityDescription, StringComparison.Ordinal));

    private static string? Value(bool present, string value) => present ? value : null;

    private static void Write(W.TableProperties properties, string? title, string? description)
    {
        foreach (var caption in properties.Elements<W.TableCaption>().ToArray()) caption.Remove();
        foreach (var currentDescription in properties.Elements<W.TableDescription>().ToArray()) currentDescription.Remove();
        if (title is not null) InsertBeforeChangeOrAppend(properties, new W.TableCaption { Val = title });
        if (description is not null) InsertBeforeChangeOrAppend(properties, new W.TableDescription { Val = description });
    }

    private static void InsertBeforeChangeOrAppend(W.TableProperties properties, OpenXmlElement element)
    {
        var change = properties.GetFirstChild<W.TablePropertiesChange>();
        if (change is null) properties.Append(element);
        else properties.InsertBefore(element, change);
    }

    private static bool TryReadValue(OpenXmlElement? element, out string? value)
    {
        value = null;
        if (element is null) return true;
        var attributes = element.GetAttributes();
        if (element.ChildElements.Count != 0 || attributes.Count != 1 ||
            !attributes[0].LocalName.Equals("val", StringComparison.Ordinal) ||
            !attributes[0].NamespaceUri.Equals(WordprocessingNamespace, StringComparison.Ordinal))
            return false;
        value = attributes[0].Value;
        return IsValidValue(value);
    }

    private static void ValidateValue(bool present, string value, string field)
    {
        if (present && !IsValidValue(value))
            throw Invalid($"Document table {field} must contain 1 through {MaxTextLength} XML-safe characters.");
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

    private static CodecException Invalid(string message) => new("invalid_document_table", message);
    private static CodecException Unsupported(string message) => new("unsupported_document_edit", message, "word/document.xml");
}
