using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeKit.Codec;

// Owns one deliberately small WordprocessingML accessibility primitive: a
// contiguous prefix of native w:tblHeader row markers. It does not infer a
// semantic header from bold/fill styling and it does not normalize arbitrary
// row properties. That keeps the source-bound edit limited to the native
// repeat-header leaves that Office hosts use at page boundaries.
internal static class DocxTableHeaderRows
{
    internal static bool TryRead(W.Table table, out uint headerRowCount)
    {
        headerRowCount = 0;
        var seenNonHeader = false;
        foreach (var row in table.Elements<W.TableRow>())
        {
            if (!TryReadRow(row, out var isHeader)) return false;
            if (isHeader)
            {
                if (seenNonHeader) return false;
                headerRowCount = checked(headerRowCount + 1);
            }
            else
            {
                seenNonHeader = true;
            }
        }
        return table.Elements<W.TableRow>().Any();
    }

    internal static void Apply(W.Table table, uint headerRowCount)
    {
        var rows = table.Elements<W.TableRow>().ToArray();
        if (rows.Length == 0 || headerRowCount > rows.Length || !TryRead(table, out _))
            throw new CodecException(
                "unsupported_document_edit",
                "Source-preserving DOCX repeat headers require a non-empty table with only canonical grid-offset and w:tblHeader row properties.",
                "word/document.xml");

        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var properties = rows[rowIndex].TableRowProperties;
            var existing = properties?.Elements<W.TableHeader>().SingleOrDefault();
            if (rowIndex < headerRowCount)
            {
                if (existing is null)
                {
                    properties ??= EnsureProperties(rows[rowIndex]);
                    // w:tblHeader follows w:gridBefore/w:gridAfter in CT_TrPr.
                    properties.Append(new W.TableHeader());
                }
            }
            else
            {
                if (existing is not null)
                {
                    existing.Remove();
                    if (!properties!.ChildElements.Any()) properties.Remove();
                }
            }
        }

        if (!TryRead(table, out var actual) || actual != headerRowCount)
            throw new CodecException(
                "document_semantics_not_applied",
                "DOCX repeat-header rows did not round trip through the bounded native profile.",
                "word/document.xml");
    }

    internal static void MaskModeled(W.Table table)
    {
        if (!TryRead(table, out _)) return;
        foreach (var properties in table.Elements<W.TableRow>()
                     .Select(row => row.TableRowProperties)
                     .Where(properties => properties is not null)
                     .Cast<W.TableRowProperties>()
                     .ToArray())
        {
            foreach (var header in properties.Elements<W.TableHeader>().ToArray()) header.Remove();
            if (!properties.ChildElements.Any()) properties.Remove();
        }
    }

    private static bool TryReadRow(W.TableRow row, out bool isHeader)
    {
        isHeader = false;
        var properties = row.TableRowProperties;
        if (properties is null) return true;
        var precedingOrder = -1;
        foreach (var child in properties.ChildElements)
        {
            var order = child switch
            {
                W.GridBefore => 0,
                W.GridAfter => 1,
                W.TableHeader => 2,
                _ => -1,
            };
            if (order < 0 || order < precedingOrder) return false;
            precedingOrder = order;
        }
        if (properties.Elements<W.GridBefore>().Skip(1).Any() || properties.Elements<W.GridAfter>().Skip(1).Any())
            return false;
        var headers = properties.Elements<W.TableHeader>().ToArray();
        if (headers.Length > 1) return false;
        if (headers.Length == 0) return true;

        // The canonical authoring form has no w:val. Explicit on/off values
        // can carry producer-specific intent, so preserve them opaque.
        if (headers[0].Val is not null)
            return false;
        isHeader = true;
        return true;
    }

    private static W.TableRowProperties EnsureProperties(W.TableRow row)
    {
        var properties = new W.TableRowProperties();
        row.PrependChild(properties);
        return properties;
    }
}
