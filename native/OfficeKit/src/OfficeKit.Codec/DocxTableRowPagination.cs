using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeKit.Codec;

// Owns the deliberately small row-layout subset of WordprocessingML row
// properties: a non-clipping w:trHeight/@w:hRule="atLeast", w:cantSplit, and
// the contiguous w:tblHeader prefix. These must share one canonical trPr
// profile so an imported edit never becomes a general row-property normalizer.
internal static class DocxTableRowPagination
{
    internal static bool TryRead(
        W.Table table,
        out uint headerRowCount,
        out uint[] keepTogetherRows,
        out uint[] minimumRowHeightsDxa)
    {
        headerRowCount = 0;
        var keepTogether = new List<uint>();
        var minimumRowHeights = new List<uint>();
        var seenNonHeader = false;
        var rows = table.Elements<W.TableRow>().ToArray();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            if (!TryReadRow(rows[rowIndex], out var isHeader, out var isKeepTogether, out var minimumRowHeight))
            {
                keepTogetherRows = [];
                minimumRowHeightsDxa = [];
                return false;
            }
            if (isHeader)
            {
                if (seenNonHeader)
                {
                    keepTogetherRows = [];
                    minimumRowHeightsDxa = [];
                    return false;
                }
                headerRowCount = checked(headerRowCount + 1);
            }
            else
            {
                seenNonHeader = true;
            }
            if (isKeepTogether) keepTogether.Add(checked((uint)rowIndex));
            minimumRowHeights.Add(minimumRowHeight);
        }
        keepTogetherRows = keepTogether.ToArray();
        minimumRowHeightsDxa = minimumRowHeights.ToArray();
        return rows.Length > 0;
    }

    internal static void Apply(
        W.Table table,
        uint headerRowCount,
        IEnumerable<uint> keepTogetherRows,
        IEnumerable<uint> minimumRowHeightsDxa)
    {
        var rows = table.Elements<W.TableRow>().ToArray();
        var requestedKeepTogetherRows = keepTogetherRows.ToArray();
        var requestedMinimumRowHeights = minimumRowHeightsDxa.ToArray();
        if (rows.Length == 0 || headerRowCount > rows.Length ||
            !HasCanonicalKeepTogetherRows(requestedKeepTogetherRows, rows.Length) ||
            !HasCanonicalMinimumRowHeights(requestedMinimumRowHeights, rows.Length) ||
            !TryRead(table, out _, out _, out _))
            throw new CodecException(
                "unsupported_document_edit",
                "Source-preserving DOCX table row layout requires a non-empty table with only canonical grid-offset, w:trHeight, w:cantSplit, and w:tblHeader row properties.",
                "word/document.xml");

        var requestedKeepTogether = requestedKeepTogetherRows.ToHashSet();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var properties = rows[rowIndex].TableRowProperties;
            if (properties is not null)
            {
                foreach (var minimumHeight in properties.Elements<W.TableRowHeight>().ToArray()) minimumHeight.Remove();
                foreach (var cantSplit in properties.Elements<W.CantSplit>().ToArray()) cantSplit.Remove();
                foreach (var header in properties.Elements<W.TableHeader>().ToArray()) header.Remove();
            }

            var minimumRowHeight = requestedMinimumRowHeights[rowIndex];
            var wantsKeepTogether = requestedKeepTogether.Contains(checked((uint)rowIndex));
            var wantsHeader = rowIndex < headerRowCount;
            if (minimumRowHeight != 0 || wantsKeepTogether || wantsHeader)
            {
                properties ??= EnsureProperties(rows[rowIndex]);
                // CT_TrPr orders these modeled leaves after w:gridBefore/After:
                // w:trHeight, w:cantSplit, then w:tblHeader.
                if (minimumRowHeight != 0)
                {
                    properties.Append(new W.TableRowHeight
                    {
                        Val = minimumRowHeight,
                        HeightType = W.HeightRuleValues.AtLeast,
                    });
                }
                if (wantsKeepTogether) properties.Append(new W.CantSplit());
                if (wantsHeader) properties.Append(new W.TableHeader());
            }
            if (properties is not null && !properties.ChildElements.Any()) properties.Remove();
        }

        if (!TryRead(table, out var actualHeaderRowCount, out var actualKeepTogetherRows, out var actualMinimumRowHeights) ||
            actualHeaderRowCount != headerRowCount ||
            !actualKeepTogetherRows.SequenceEqual(requestedKeepTogetherRows) ||
            !actualMinimumRowHeights.SequenceEqual(requestedMinimumRowHeights))
            throw new CodecException(
                "document_semantics_not_applied",
                "DOCX table row layout did not round trip through the bounded native profile.",
                "word/document.xml");
    }

    internal static void MaskModeled(W.Table table)
    {
        if (!TryRead(table, out _, out _, out _)) return;
        foreach (var properties in table.Elements<W.TableRow>()
                     .Select(row => row.TableRowProperties)
                     .Where(properties => properties is not null)
                     .Cast<W.TableRowProperties>()
                     .ToArray())
        {
            foreach (var minimumHeight in properties.Elements<W.TableRowHeight>().ToArray()) minimumHeight.Remove();
            foreach (var cantSplit in properties.Elements<W.CantSplit>().ToArray()) cantSplit.Remove();
            foreach (var header in properties.Elements<W.TableHeader>().ToArray()) header.Remove();
            if (!properties.ChildElements.Any()) properties.Remove();
        }
    }

    private static bool TryReadRow(W.TableRow row, out bool isHeader, out bool isKeepTogether, out uint minimumRowHeight)
    {
        isHeader = false;
        isKeepTogether = false;
        minimumRowHeight = 0;
        var properties = row.TableRowProperties;
        if (properties is null) return true;
        var precedingOrder = -1;
        foreach (var child in properties.ChildElements)
        {
            var order = child switch
            {
                W.GridBefore => 0,
                W.GridAfter => 1,
                W.TableRowHeight => 2,
                W.CantSplit => 3,
                W.TableHeader => 4,
                _ => -1,
            };
            if (order < 0 || order < precedingOrder) return false;
            precedingOrder = order;
        }
        if (properties.Elements<W.GridBefore>().Skip(1).Any() || properties.Elements<W.GridAfter>().Skip(1).Any())
            return false;

        var heights = properties.Elements<W.TableRowHeight>().ToArray();
        if (heights.Length > 1) return false;
        if (heights.Length == 1)
        {
            if (heights[0].Val?.Value is not uint value || value == 0 ||
                value > 1_000_000 || heights[0].HeightType?.Value != W.HeightRuleValues.AtLeast)
                return false;
            minimumRowHeight = value;
        }

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

    internal static bool HasCanonicalMinimumRowHeights(IEnumerable<uint> values, int rowCount)
    {
        var count = 0;
        foreach (var value in values)
        {
            if (value > 1_000_000) return false;
            count++;
        }
        return count == rowCount;
    }

    private static W.TableRowProperties EnsureProperties(W.TableRow row)
    {
        var properties = new W.TableRowProperties();
        row.PrependChild(properties);
        return properties;
    }
}
