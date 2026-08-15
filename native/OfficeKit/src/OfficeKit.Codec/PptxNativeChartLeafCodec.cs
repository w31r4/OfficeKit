using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeKit.Codec;

internal sealed record PptxNativeChartTitleLeaf(uint Index, string Text, XElement Element);

internal sealed record PptxNativeChartDataPointResolution(
    PresentationNativeChartDataPoint Binding,
    XElement CacheValue,
    WorksheetPart WorksheetPart,
    S.Cell Cell);

internal sealed record PptxNativeChartDataResolution(
    EmbeddedPackagePart Part,
    byte[] PackageBytes,
    IReadOnlyList<PptxNativeChartDataPointResolution> Points);

internal sealed record PptxNativeChartResolution(
    PresentationNativeChart Binding,
    ChartPart Part,
    XDocument Document,
    IReadOnlyList<PptxNativeChartTitleLeaf> TitleLeaves,
    PptxNativeChartDataResolution? Data);

// Describes and re-proves a deliberately tiny source-owned chart surface.
// A native chart remains opaque; only direct rich-title a:r/a:t leaves and
// uniquely bound direct numeric bar-cache points are projected. Styles,
// extensions, plot topology, and every other ChartSpace token stay owned by
// the original package.
internal static partial class PptxNativeChartLeafCodec
{
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const int MaxTitleLeaves = 256;
    private const int MaxDataPointLeaves = 256;
    private const int MaxLeafLength = 32_767;

    private static readonly XNamespace ChartNs = ChartNamespace;
    private static readonly XNamespace DrawingNs = DrawingNamespace;
    private static readonly XNamespace RelationshipsNs = RelationshipsNamespace;

    [GeneratedRegex("^(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!\\[\\]]+))!\\$(?<startColumn>[A-Za-z]{1,3})\\$(?<startRow>[1-9][0-9]*)(?::\\$(?<endColumn>[A-Za-z]{1,3})\\$(?<endRow>[1-9][0-9]*))?$", RegexOptions.CultureInvariant)]
    private static partial Regex CellRangeFormulaPattern();

    internal static bool TryDescribe(OpenXmlElement source, OpenXmlPart owner, EffectiveCodecLimits limits, out PresentationNativeChart binding)
    {
        if (!TryResolve(source, owner, limits, out var resolved))
        {
            binding = null!;
            return false;
        }
        binding = resolved.Binding;
        return true;
    }

    internal static bool TryResolve(OpenXmlElement source, OpenXmlPart owner, EffectiveCodecLimits limits, out PptxNativeChartResolution resolved)
    {
        resolved = null!;
        if (source is not P.GraphicFrame frame || owner is not SlidePart || source.Parent is not P.ShapeTree ||
            PptxNativeObjectCatalog.Classify(source) != "graphicFrame")
            return false;
        var graphicData = frame.Graphic?.GraphicData;
        if (graphicData is null || graphicData.Uri?.Value != ChartNamespace) return false;
        var references = graphicData.Elements<C.ChartReference>().ToArray();
        if (references.Length != 1 || string.IsNullOrWhiteSpace(references[0].Id?.Value)) return false;
        var relationshipId = references[0].Id!.Value!;

        ChartPart part;
        try
        {
            part = owner.GetPartById(relationshipId) as ChartPart ?? null!;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
        if (part is null || !part.ContentType.Equals(ChartContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/chart", StringComparison.Ordinal)) return false;

        byte[] bytes;
        XDocument document;
        try
        {
            bytes = ReadPart(part);
            using var memory = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(memory, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                IgnoreWhitespace = false,
            });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is XmlException or IOException or UnauthorizedAccessException)
        {
            return false;
        }

        var root = document.Root;
        if (root?.Name != ChartNs + "chartSpace") return false;
        var charts = root.Elements(ChartNs + "chart").ToArray();
        if (charts.Length != 1) return false;
        var textElements = ResolveTitleLeaves(charts[0]);

        var binding = new PresentationNativeChart
        {
            PartPath = PartPath(part),
            ContentType = part.ContentType,
            SourceSha256 = Hash(bytes),
            RelationshipId = relationshipId,
        };
        var leaves = textElements.Select((element, index) =>
            new PptxNativeChartTitleLeaf(checked((uint)index), element.Value, element)).ToArray();
        binding.TitleLeaves.Add(leaves.Select(leaf => new PresentationNativeChartTitleLeaf
        {
            TextLeafIndex = leaf.Index,
            Text = leaf.Text,
        }));
        PptxNativeChartDataResolution? data = null;
        if (TryResolveData(part, document, limits, out var resolvedData))
        {
            data = resolvedData;
            binding.EmbeddedPackagePartPath = PartPath(data.Part);
            binding.EmbeddedPackageSourceSha256 = Hash(data.PackageBytes);
            binding.EmbeddedPackageRelationshipId = part.GetIdOfPart(data.Part);
            binding.DataPoints.Add(data.Points.Select(point => point.Binding));
        }
        if (leaves.Length == 0 && data is null) return false;
        resolved = new PptxNativeChartResolution(binding, part, document, leaves, data);
        return true;
    }

    private static XElement[] ResolveTitleLeaves(XElement chart)
    {
        var titles = chart.Elements(ChartNs + "title").ToArray();
        if (titles.Length != 1) return [];
        var textOwners = titles[0].Elements(ChartNs + "tx").ToArray();
        if (textOwners.Length != 1) return [];
        var richBodies = textOwners[0].Elements(ChartNs + "rich").ToArray();
        if (richBodies.Length != 1 || richBodies[0].Descendants(DrawingNs + "fld").Any()) return [];
        var textElements = richBodies[0].Descendants(DrawingNs + "t").ToArray();
        return textElements.Length is > 0 and <= MaxTitleLeaves && textElements.All(element =>
            element.Parent?.Name == DrawingNs + "r" && ValidText(element.Value)) ? textElements : [];
    }

    internal static bool SameBinding(PresentationNativeChart expected, PresentationNativeChart actual) =>
        SameTitleBinding(expected, actual) &&
        expected.EmbeddedPackagePartPath.Equals(actual.EmbeddedPackagePartPath, StringComparison.OrdinalIgnoreCase) &&
        expected.EmbeddedPackageSourceSha256.Equals(actual.EmbeddedPackageSourceSha256, StringComparison.OrdinalIgnoreCase) &&
        expected.EmbeddedPackageRelationshipId == actual.EmbeddedPackageRelationshipId &&
        expected.DataPoints.SequenceEqual(actual.DataPoints);

    internal static bool SameTitleBinding(PresentationNativeChart expected, PresentationNativeChart actual) =>
        expected.PartPath.Equals(actual.PartPath, StringComparison.OrdinalIgnoreCase) &&
        expected.ContentType.Equals(actual.ContentType, StringComparison.OrdinalIgnoreCase) &&
        expected.SourceSha256.Equals(actual.SourceSha256, StringComparison.OrdinalIgnoreCase) &&
        expected.RelationshipId == actual.RelationshipId &&
        expected.TitleLeaves.SequenceEqual(actual.TitleLeaves);

    private static bool TryResolveData(
        ChartPart chartPart,
        XDocument chartDocument,
        EffectiveCodecLimits limits,
        out PptxNativeChartDataResolution resolved)
    {
        resolved = null!;
        try
        {
            var root = chartDocument.Root!;
            var externalData = root.Elements(ChartNs + "externalData").ToArray();
            if (externalData.Length != 1) return false;
            var relationshipId = externalData[0].Attribute(RelationshipsNs + "id")?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) return false;
            var packagePart = chartPart.GetPartById(relationshipId) as EmbeddedPackagePart;
            if (packagePart is null || !packagePart.ContentType.Equals(SpreadsheetContentType, StringComparison.OrdinalIgnoreCase) ||
                !packagePart.RelationshipType.EndsWith("/package", StringComparison.Ordinal)) return false;
            var packageBytes = ReadPart(packagePart);
            _ = PackageGuards.ValidateAndCollectOpaque(packageBytes, limits, OpcPackageProfile.Xlsx, includeSourcePackage: false);

            using var stream = new MemoryStream(packageBytes, writable: false);
            using var workbook = SpreadsheetDocument.Open(stream, isEditable: false);
            var workbookPart = workbook.WorkbookPart;
            var sheets = workbookPart?.Workbook?.Sheets?.Elements<S.Sheet>().ToArray() ?? [];
            if (workbookPart is null || sheets.Length == 0 || (uint)sheets.Length > limits.MaxSheets) return false;
            var chart = AssertSingle(root.Elements(ChartNs + "chart"));
            var plotArea = AssertSingle(chart?.Elements(ChartNs + "plotArea"));
            if (plotArea is null) return false;
            var series = plotArea.Elements()
                .Where(element => element.Name == ChartNs + "barChart")
                .SelectMany(element => element.Elements(ChartNs + "ser"))
                .ToArray();
            if (series.Length == 0 || series.Length > 64) return false;

            var points = new List<PptxNativeChartDataPointResolution>();
            for (var seriesIndex = 0; seriesIndex < series.Length; seriesIndex++)
            {
                var valueOwner = AssertSingle(series[seriesIndex].Elements(ChartNs + "val"));
                var numberReference = AssertSingle(valueOwner?.Elements(ChartNs + "numRef"));
                var formula = AssertSingle(numberReference?.Elements(ChartNs + "f"))?.Value;
                var cache = AssertSingle(numberReference?.Elements(ChartNs + "numCache"));
                if (string.IsNullOrWhiteSpace(formula) || cache is null || !TryParseRange(formula, out var range)) continue;
                var sheet = sheets.SingleOrDefault(candidate => string.Equals(candidate.Name?.Value, range.SheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet?.Id?.Value is not string sheetRelationshipId || workbookPart.GetPartById(sheetRelationshipId) is not WorksheetPart worksheetPart) continue;
                var worksheetBytes = ReadPart(worksheetPart);
                var worksheetHash = Hash(worksheetBytes);
                var cachePoints = cache.Elements(ChartNs + "pt").ToArray();
                foreach (var cachePoint in cachePoints)
                {
                    if (!uint.TryParse(cachePoint.Attribute("idx")?.Value, out var pointIndex) || pointIndex >= range.Length) continue;
                    var cacheValue = AssertSingle(cachePoint.Elements(ChartNs + "v"));
                    if (cacheValue is null || !ValidNumber(cacheValue.Value)) continue;
                    var cellReference = range.CellAt(pointIndex);
                    var cells = worksheetPart.Worksheet?.Descendants<S.Cell>()
                        .Where(cell => string.Equals(cell.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase))
                        .ToArray() ?? [];
                    var dataType = cells.Length == 1 ? cells[0].DataType?.Value : null;
                    if (cells.Length != 1 || cells[0].CellFormula is not null || cells[0].InlineString is not null ||
                        (dataType is not null && dataType != S.CellValues.Number) || cells[0].CellValue is null ||
                        cells[0].CellValue!.Text != cacheValue.Value) continue;
                    points.Add(new PptxNativeChartDataPointResolution(
                        new PresentationNativeChartDataPoint
                        {
                            SeriesIndex = checked((uint)seriesIndex),
                            PointIndex = pointIndex,
                            Value = cacheValue.Value,
                            Formula = formula,
                            WorksheetPartPath = PartPath(worksheetPart),
                            WorksheetSourceSha256 = worksheetHash,
                            WorksheetName = range.SheetName,
                            CellReference = cellReference,
                        },
                        cacheValue,
                        worksheetPart,
                        cells[0]));
                    if (points.Count > MaxDataPointLeaves) return false;
                }
            }
            if (points.Count == 0 || points
                    .GroupBy(point => (point.Binding.SeriesIndex, point.Binding.PointIndex))
                    .Any(group => group.Count() != 1))
                return false;
            resolved = new PptxNativeChartDataResolution(packagePart, packageBytes, points);
            return true;
        }
        catch (Exception exception) when (exception is CodecException or OpenXmlPackageException or InvalidDataException or XmlException or IOException or UnauthorizedAccessException or ArgumentOutOfRangeException or InvalidOperationException)
        {
            return false;
        }
    }

    private static XElement? AssertSingle(IEnumerable<XElement>? elements)
    {
        if (elements is null) return null;
        using var enumerator = elements.GetEnumerator();
        if (!enumerator.MoveNext()) return null;
        var value = enumerator.Current;
        return enumerator.MoveNext() ? null : value;
    }

    private static bool ValidNumber(string value) =>
        value.Length is > 0 and <= 128 &&
        double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number) &&
        double.IsFinite(number);

    private static bool TryParseRange(string formula, out PptxCellRange range)
    {
        range = default;
        var match = CellRangeFormulaPattern().Match(formula);
        if (!match.Success) return false;
        var sheetName = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
            : match.Groups["plain"].Value;
        var startColumn = ColumnNumber(match.Groups["startColumn"].Value);
        var endColumn = match.Groups["endColumn"].Success ? ColumnNumber(match.Groups["endColumn"].Value) : startColumn;
        if (!uint.TryParse(match.Groups["startRow"].Value, out var startRow)) return false;
        var endRow = match.Groups["endRow"].Success && uint.TryParse(match.Groups["endRow"].Value, out var parsedEndRow) ? parsedEndRow : startRow;
        if (startColumn <= 0 || endColumn <= 0 || startRow == 0 || endRow == 0 ||
            (startColumn != endColumn && startRow != endRow) || endColumn < startColumn || endRow < startRow) return false;
        var length = checked((uint)(endColumn - startColumn) + (endRow - startRow) + 1);
        if (length is 0 or > MaxDataPointLeaves) return false;
        range = new PptxCellRange(sheetName, startColumn, startRow, endColumn, endRow, length);
        return true;
    }

    private static int ColumnNumber(string value)
    {
        var result = 0;
        foreach (var character in value.ToUpperInvariant())
        {
            if (character is < 'A' or > 'Z') return 0;
            result = checked(result * 26 + character - 'A' + 1);
        }
        return result is > 0 and <= 16_384 ? result : 0;
    }

    private static string ColumnName(int value)
    {
        var result = string.Empty;
        for (var current = value; current > 0; current = (current - 1) / 26)
            result = (char)('A' + ((current - 1) % 26)) + result;
        return result;
    }

    private readonly record struct PptxCellRange(
        string SheetName,
        int StartColumn,
        uint StartRow,
        int EndColumn,
        uint EndRow,
        uint Length)
    {
        internal string CellAt(uint index) => StartColumn == EndColumn
            ? $"{ColumnName(StartColumn)}{StartRow + index}"
            : $"{ColumnName(StartColumn + checked((int)index))}{StartRow}";
    }

    private static bool ValidText(string value) =>
        value.Length <= MaxLeafLength &&
        !value.Any(character => character is >= '\u0000' and <= '\u0008' or '\u000B' or '\u000C' or >= '\u000E' and <= '\u001F');

    private static byte[] ReadPart(OpenXmlPart part)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
