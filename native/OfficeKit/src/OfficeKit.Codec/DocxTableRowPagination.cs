using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeKit.Codec;

// Owns the deliberately small pagination subset of WordprocessingML row
// properties: w:cantSplit and the contiguous w:tblHeader prefix. Both affect
// page flow rather than visual table formatting, and both must share one
// canonical trPr profile so an imported edit never becomes a general row-
// property normalizer.
internal static class DocxTableRowPagination
{
    internal static bool TryRead(W.Table table, out uint headerRowCount, out uint[] keepTogetherRows)
    {
        headerRowCount = 0;
        var keepTogether = new List<uint>();
        var seenNonHeader = false;
        var rows = table.Elements<W.TableRow>().ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            if (!TryReadRow(rows[rowIndex], out var isHeader, out var isKeepTogether))
            {
                keepTogetherRows = [];
                return false;
            }
            if (isHeader)
            {
                if (seenNonHeader)
                {
                    keepTogetherRows = [];
                    return false;
                }
                headerRowCount = checked(headerRowCount + 1);
            }
            else
            {
                seenNonHeader = true;
            }
            if (isKeepTogether) keepTogether.Add(checked((uint)rowIndex));
        }
        keepTogetherRows = keepTogether.ToArray();
        return rows.Length > 0;
    }

    internal static void Apply(W.Table table, uint headerRowCount, IEnumerable<uint> keepTogetherRows)
    {
        var rows = table.Elements<W.TableRow>().ToArray();
        var requestedKeepTogetherRows = keepTogetherRows.ToArray();
        if (rows.Length == 0 || headerRowCount > rows.Length ||
            !HasCanonicalKeepTogetherRows(requestedKeepTogetherRows, rows.Length) ||
            !TryRead(table, out _, out _))
            throw new CodecException(
                "unsupported_document_edit",
                "Source-preserving DOCX table pagination requires a non-empty table with only canonical grid-offset, w:cantSplit, and w:tblHeader row properties.",
                "word/document.xml");

        var requestedKeepTogether = requestedKeepTogetherRows.ToHashSet();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var properties = rows[rowIndex].TableRowProperties;
            if (properties is not null)
            {
                foreach (var cantSplit in properties.Elements<W.CantSplit>().ToArray()) cantSplit.Remove();
                foreach (var header in properties.Elements<W.TableHeader>().ToArray()) header.Remove();
            }

            var wantsKeepTogether = requestedKeepTogether.Contains(checked((uint)rowIndex));
            var wantsHeader = rowIndex < headerRowCount;
            if (wantsKeepTogether || wantsHeader)
            {
                properties ??= EnsureProperties(rows[rowIndex]);
                // CT_TrPr orders the modeled leaves after w:gridBefore/After:
                // w:cantSplit precedes w:tblHeader.
                if (wantsKeepTogether) properties.Append(new W.CantSplit());
                if (wantsHeader) properties.Append(new W.TableHeader());
            }
            if (properties is not null && !properties.ChildElements.Any()) properties.Remove();
        }

        if (!TryRead(table, out var actualHeaderRowCount, out var actualKeepTogetherRows) ||
            actualHeaderRowCount != headerRowCount ||
            !actualKeepTogetherRows.SequenceEqual(requestedKeepTogetherRows))
            throw new CodecException(
                "document_semantics_not_applied",
                "DOCX table pagination rows did not round trip through the bounded native profile.",
                "word/document.xml");
    }

    internal static void MaskModeled(W.Table table)
    {
        if (!TryRead(table, out _, out _)) return;
        foreach (var properties in table.Elements<W.TableRow>()
                     .Select(row => row.TableRowProperties)
                     .Where(properties => properties is not null)
                     .Cast<W.TableRowProperties>()
                     .ToArray())
        {
            foreach (var cantSplit in properties.Elements<W.CantSplit>().ToArray()) cantSplit.Remove();
            foreach (var header in properties.Elements<W.TableHeader>().ToArray()) header.Remove();
            if (!properties.ChildElements.Any()) properties.Remove();
        }
    }

    private static bool TryReadRow(W.TableRow row, out bool isHeader, out bool isKeepTogether)
    {
        isHeader = false;
        isKeepTogether = false;
        var properties = row.TableRowProperties;
        if (properties is null) return true;
        var precedingOrder = -1;
        foreach (var child in properties.ChildElements)
        {
            var order = child switch
            {
                W.GridBefore => 0,
                W.GridAfter => 1,
                W.CantSplit => 2,
                W.TableHeader => 3,
                _ => -1,
            };
            if (order < 0 || order < precedingOrder) return false;
            precedingOrder = order;
        }
        if (properties.Elements<W.GridBefore>().Skip(1).Any() || properties.Elements<W.GridAfter>().Skip(1).Any())
            return false;

        var cantSplits = properties.Elements<W.CantSplit>().ToArray();
        if (cantSplits.Length > 1 || (cantSplits.Length == 1 && cantSplits[0].Val is not null)) return false;
        var headers = properties.Elements<W.TableHeader>().ToArray();
        if (headers.Length > 1 || (headers.Length == 1 && headers[0].Val is not null)) return false;
        isKeepTogether = cantSplits.Length == 1;
        isHeader = headers.Length == 1;
        return true;
    }

    internal static bool HasCanonicalKeepTogetherRows(IEnumerable<uint> values, int rowCount)
    {
        uint previous = 0;
        var hasPrevious = false;
        foreach (var value in values)
        {
            if (value >= rowCount || (hasPrevious && value <= previous)) return false;
            previous = value;
            hasPrevious = true;
        }
        return true;
    }

    private static W.TableRowProperties EnsureProperties(W.TableRow row)
    {
        var properties = new W.TableRowProperties();
        row.PrependChild(properties);
        return properties;
    }
}
