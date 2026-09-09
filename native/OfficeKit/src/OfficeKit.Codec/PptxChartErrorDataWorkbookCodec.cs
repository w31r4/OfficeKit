using System.Globalization;
using System.Security.Cryptography;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeKit.Codec;

internal static class PptxChartErrorDataWorkbookCodec
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static bool TryParseRange(string formula, out PptxNativeChartLeafCodec.PptxCellRange range) =>
        PptxNativeChartLeafCodec.TryParseRange(formula, out range, maxPoints: 1_048_576);

    // Only opt typed PPTX charts into embedded workbooks for custom error data.
    // Other workbook-backed chart profiles retain the existing opaque path.
    internal static bool SupportsExternalData(XElement root)
    {
        var bindings = root.Elements(C + "externalData").ToArray();
        if (bindings.Length == 0) return true;
        return bindings.Length == 1 && bindings[0].Attribute(R + "id") is { Value.Length: > 0 } &&
            bindings[0].Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name == R + "id") &&
            bindings[0].Elements().All(element => element.Name == C + "autoUpdate" && !element.HasElements &&
                element.Attributes().All(attribute => attribute.Name == "val") && (string?)element.Attribute("val") is "0" or "1" or "false" or "true") &&
            root.Descendants(C + "errBars").Elements().Where(side => side.Name == C + "plus" || side.Name == C + "minus")
                .Elements(C + "numRef").Elements(C + "f").Any(formula => TryParseRange(formula.Value, out _));
    }
    internal sealed record Replacement(EmbeddedPackagePart Part, byte[] Bytes)
    {
        internal string Path => Part.Uri.OriginalString.TrimStart('/');
        internal string Sha256 => Convert.ToHexString(SHA256.HashData(Bytes)).ToLowerInvariant();
    }

    internal static Replacement? Prepare(ChartPart part, XDocument chartXml, PresentationChart original, PresentationChart requested, EffectiveCodecLimits limits)
    {
        var oldSeries = Series(original);
        var newSeries = Series(requested);
        if (chartXml.Root!.Element(C + "externalData") is not null)
        {
            for (var index = 0; index < oldSeries.Length; index++)
            {
                var before = oldSeries[index];
                var after = newSeries[index];
                if (before.CategoryFormula != after.CategoryFormula || before.ValueFormula != after.ValueFormula ||
                    before.XValueFormula != after.XValueFormula || before.BubbleSizeFormula != after.BubbleSizeFormula ||
                    before.CategoryFormula.Length > 0 && !original.Categories.SequenceEqual(requested.Categories) ||
                    before.ValueFormula.Length > 0 && (!before.Values.SequenceEqual(after.Values) || !before.MissingValueIndexes.SequenceEqual(after.MissingValueIndexes)) ||
                    before.XValueFormula.Length > 0 && !before.XValues.SequenceEqual(after.XValues) ||
                    before.BubbleSizeFormula.Length > 0 && !before.BubbleSizes.SequenceEqual(after.BubbleSizes))
                    throw Rejected("workbook-backed channels outside error data cannot change through this operation");
            }
        }
        var nativeSeries = chartXml.Descendants(C + "ser").ToArray();
        if (original.Type == SpreadsheetChartType.Combo)
            nativeSeries = nativeSeries.OrderBy(series => (uint?)series.Element(C + "order")?.Attribute("val") ?? uint.MaxValue).ToArray();
        var sides = new List<(XElement Formula, SpreadsheetChartErrorBarDataArtifact Before, SpreadsheetChartErrorBarDataArtifact After)>();
        for (var index = 0; index < oldSeries.Length; index++)
        foreach (var side in new[] { "plus", "minus" })
        {
            var before = side == "plus" ? oldSeries[index].ErrorBars?.Plus : oldSeries[index].ErrorBars?.Minus;
            var after = side == "plus" ? newSeries[index].ErrorBars?.Plus : newSeries[index].ErrorBars?.Minus;
            if ((before?.Formula ?? "") != (after?.Formula ?? "")) throw Rejected("error-data formula binding topology change");
            if (before is null || before.Formula.Length == 0) continue;
            if (after is null || before.Values.Count != after.Values.Count || nativeSeries.Length != oldSeries.Length)
                throw Rejected("error-data point topology change");
            var formula = nativeSeries[index].Element(C + "errBars")?.Element(C + side)?.Element(C + "numRef")?.Element(C + "f");
            if (formula?.Value != before.Formula) throw Rejected("error-data formula owner does not match its series");
            sides.Add((formula, before, after));
        }
        if (sides.All(side => side.Before.Values.SequenceEqual(side.After.Values))) return null;

        var presentation = FindPresentation(part);
        if (presentation is null || !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentation, part))
            throw Rejected("ChartPart must have one owning relationship");
        var external = chartXml.Root!.Elements(C + "externalData").ToArray();
        if (external.Length != 1 || external[0].Attribute(R + "id")?.Value is not { Length: > 0 } relationshipId ||
            part.Parts.SingleOrDefault(item => item.RelationshipId == relationshipId).OpenXmlPart is not EmbeddedPackagePart embedded ||
            embedded.ContentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
            !PptxNativeObjectCatalog.HasUniqueInboundRelationship(presentation, embedded))
            throw Rejected("error data requires a uniquely owned embedded XLSX workbook");
        var bytes = Read(embedded);
        _ = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Xlsx, includeSourcePackage: false);
        using var memory = new MemoryStream(bytes, writable: false);
        using var workbook = SpreadsheetDocument.Open(memory, isEditable: false);
        var book = workbook.WorkbookPart ?? throw Rejected("embedded workbook root is missing");
        if (book.Workbook.Descendants<S.DefinedName>().Any() || book.Workbook.Descendants().Any(node => node.LocalName == "extLst") || book.ExternalRelationships.Any() ||
            book.Parts.Any(item => item.OpenXmlPart is not (WorksheetPart or WorkbookStylesPart or SharedStringTablePart or ThemePart)))
            throw Rejected("embedded workbook dependency graph is outside numeric-cell replacement");
        var sheets = book.Workbook.Sheets?.Elements<S.Sheet>().ToArray() ?? [];
        var worksheetOwners = new HashSet<OpenXmlPart>();
        foreach (var sheet in sheets)
            if (sheet.Id?.Value is not { } id || book.GetPartById(id) is not WorksheetPart worksheet || !worksheetOwners.Add(worksheet))
                throw Rejected("embedded worksheet must have one sheet owner");
        if ((uint)sheets.Length > limits.MaxSheets ||
            book.WorksheetParts.Aggregate(0UL, (count, worksheet) => count + (ulong)worksheet.Worksheet.Descendants<S.Cell>().LongCount()) > limits.MaxCells)
            throw Rejected("embedded workbook exceeds worksheet or cell limits");
        foreach (var worksheet in book.WorksheetParts)
            if (worksheet.Parts.Any() || worksheet.ExternalRelationships.Any() ||
                worksheet.Worksheet.Descendants().Any(node => node.LocalName is "f" or "formula" or "formula1" or "formula2" or "mergeCell" or "extLst"))
                throw Rejected("embedded worksheet contains formulas or dependent content");

        var formulas = new List<(XElement Owner, PptxNativeChartLeafCodec.PptxCellRange Range)>();
        foreach (var formula in chartXml.Descendants(C + "f"))
        {
            if (!TryParseRange(formula.Value, out var range))
                throw Rejected("chart contains an unresolved formula consumer");
            formulas.Add((formula, range));
        }
        var operations = new List<PresentationEditOperation>();
        foreach (var side in sides)
        {
            var range = formulas.Single(item => ReferenceEquals(item.Owner, side.Formula)).Range;
            if (range.Length != side.Before.Values.Count) throw Rejected("error-data cache must cover the entire formula range");
            var sheet = sheets.SingleOrDefault(item => string.Equals(item.Name?.Value, range.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet?.Id?.Value is not { } sheetId || book.GetPartById(sheetId) is not WorksheetPart worksheet)
                throw Rejected("error-data worksheet is missing");
            var worksheetCells = worksheet.Worksheet.Descendants<S.Cell>().ToLookup(cell => cell.CellReference?.Value ?? "", StringComparer.OrdinalIgnoreCase);
            for (var point = 0; point < side.Before.Values.Count; point++)
            {
                var address = range.CellAt((uint)point);
                var cells = worksheetCells[address].Take(2).ToArray();
                if (cells.Length != 1 || cells[0].CellFormula is not null || cells[0].ChildElements.Count != 1 ||
                    cells[0].CellValue is not { } value || cells[0].DataType?.Value is { } type && type != S.CellValues.Number ||
                    !double.TryParse(value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number) || number != side.Before.Values[point])
                    throw Rejected("error-data cache and numeric worksheet cell disagree");
                if (side.Before.Values[point] == side.After.Values[point]) continue;
                var column = range.StartColumn + (range.StartRow == range.EndRow ? point : 0);
                var row = range.StartRow + (range.StartColumn == range.EndColumn ? (uint)point : 0);
                if (formulas.Any(item => !ReferenceEquals(item.Owner, side.Formula) &&
                    string.Equals(item.Range.SheetName, range.SheetName, StringComparison.OrdinalIgnoreCase) &&
                    column >= item.Range.StartColumn && column <= item.Range.EndColumn && row >= item.Range.StartRow && row <= item.Range.EndRow))
                    throw Rejected("error-data cell is shared by another chart formula consumer");
                operations.Add(new PresentationEditOperation
                {
                    OperationId = $"error-data-{operations.Count}", LeafKind = "chartDataValue",
                    EmbeddedWorksheetPartPath = worksheet.Uri.OriginalString.TrimStart('/'),
                    EmbeddedCellReference = address, ExpectedValue = value.Text,
                    Value = side.After.Values[point].ToString("R", CultureInfo.InvariantCulture),
                });
            }
        }
        return new(embedded, PptxEditPlanCodec.PatchEmbeddedNumericCells(bytes, operations));
    }

    private static SpreadsheetChartSeriesArtifact[] Series(PresentationChart chart) => chart.Type == SpreadsheetChartType.Combo
        ? chart.ComboSeries.Select(series => series.Series).ToArray() : chart.Series.ToArray();

    private static PresentationPart? FindPresentation(OpenXmlPart part)
    {
        var visited = new HashSet<OpenXmlPart>();
        var pending = new Queue<OpenXmlPart>();
        pending.Enqueue(part);
        while (pending.Count > 0)
        {
            var current = pending.Dequeue();
            if (!visited.Add(current)) continue;
            if (current is PresentationPart presentation) return presentation;
            foreach (var parent in current.GetParentParts()) pending.Enqueue(parent);
        }
        return null;
    }

    private static byte[] Read(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static CodecException Rejected(string reason) => new("presentation_chart_error_data_binding", reason);
}
