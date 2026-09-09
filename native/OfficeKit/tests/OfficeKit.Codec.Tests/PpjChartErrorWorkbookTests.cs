using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("column", 0, false)]
    [InlineData("bar", 0, false)]
    [InlineData("line", 0, false)]
    [InlineData("combo", 1, false)]
    [InlineData("line", 0, true)]
    public void PpjErrorWorkbookSynchronizesCacheAndCells(string type, int index, bool grouped)
    {
        var (source, chartPath, workbookPath) = ErrorWorkbookSource(type, index, grouped);
        var projected = ProjectTrendlineList(source);
        Assert.Equal("'Sheet1'!$D$2:$D$5", ErrorWorkbookSeries(projected, index)["errorBars"]!["plus"]!["formula"]!.GetValue<string>());
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);
        Assert.Equal(source, noOp.File.ToByteArray());
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        for (var round = 0; round < 2; round++)
        {
            var errors = ErrorWorkbookSeries(projected, index)["errorBars"]!;
            errors["plus"]!["values"]![0] = round == 0 ? 0.75 : 0;
            errors["minus"]!["values"]![3] = round == 0 ? 0 : 2.5;
            var expected = errors.DeepClone();
            var result = CompileTrendlineList(projected, source);
            Assert.True(result.Ok, Diagnostics(result));
            Assert.Equal(new[] { chartPath, workbookPath }.Order(), result.PresentationProgram.ChangedParts.Select(path => path.TrimStart('/')).Order());
            var output = RemoveEmbeddedPpj(result.File.ToByteArray());
            foreach (var part in ZipPartPaths(source).Where(path => path != chartPath && path != workbookPath))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var oldChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, chartPath)));
            var newChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, chartPath)));
            Assert.Equal(oldChart.Descendants(c + "f").Select(node => node.Value), newChart.Descendants(c + "f").Select(node => node.Value));
            oldChart.Descendants(c + "errBars").Remove();
            newChart.Descendants(c + "errBars").Remove();
            Assert.True(XNode.DeepEquals(oldChart, newChart));
            var oldBook = ZipBytes(source, workbookPath);
            var newBook = ZipBytes(output, workbookPath);
            Assert.Equal(ZipPartPaths(oldBook), ZipPartPaths(newBook));
            var worksheetPath = SingleZipEntryPath(oldBook, path => path.StartsWith("xl/worksheets/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
            foreach (var part in ZipPartPaths(oldBook).Where(path => path != worksheetPath))
                Assert.Equal(ZipBytes(oldBook, part), ZipBytes(newBook, part));
            var oldSheet = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(oldBook, worksheetPath)));
            var newSheet = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(newBook, worksheetPath)));
            foreach (var (address, value) in new[] { ("D2", round == 0 ? "0.75" : "0"), ("E5", round == 0 ? "0" : "2.5") })
            {
                oldSheet.Descendants(s + "c").Single(cell => (string?)cell.Attribute("r") == address).Element(s + "v")!.Value = value;
                Assert.Equal(value, newSheet.Descendants(s + "c").Single(cell => (string?)cell.Attribute("r") == address).Element(s + "v")!.Value);
            }
            Assert.True(XNode.DeepEquals(oldSheet, newSheet));
            source = output;
            projected = ProjectTrendlineList(source);
            Assert.True(JsonNode.DeepEquals(expected, ErrorWorkbookSeries(projected, index)["errorBars"]));
        }
        // Changing caps/stroke/format without values leaves the workbook bytes alone.
        ErrorWorkbookSeries(projected, index)["errorBars"]!["noEndCap"] = true;
        var styled = CompileTrendlineList(projected, source);
        Assert.True(styled.Ok, Diagnostics(styled));
        Assert.Equal(chartPath, Assert.Single(styled.PresentationProgram.ChangedParts).TrimStart('/'));
        Assert.Equal(ZipBytes(source, workbookPath), ZipBytes(styled.File.ToByteArray(), workbookPath));
    }

    [Fact]
    public void PpjErrorWorkbookRejectsBindingTopologyChanges()
    {
        var (source, _, _) = ErrorWorkbookSource("line", 0, false);
        foreach (var mutation in new Action<JsonObject>[]
        {
            errors => errors["plus"]!["formula"] = "'Sheet1'!$F$2:$F$5",
            errors => errors["plus"]!.AsObject().Remove("formula"),
            errors => errors.Remove("plus"),
        })
        {
            var program = ProjectTrendlineList(source);
            mutation(ErrorWorkbookSeries(program, 0)["errorBars"]!.AsObject());
            var rejected = CompileTrendlineList(program, source);
            Assert.False(rejected.Ok);
            Assert.Empty(rejected.File);
        }
        var authored = TrendlineListProgram("line", 0);
        TrendlineListSeries(authored, 0)["errorBars"] = JsonNode.Parse(CustomErrorData);
        TrendlineListSeries(authored, 0)["errorBars"]!["plus"]!["formula"] = "'Sheet1'!$D$2:$D$5";
        Assert.False(CompileTrendlineList(authored).Ok);
    }

    [Theory]
    [InlineData("stale")]
    [InlineData("formula-cell")]
    [InlineData("shared")]
    [InlineData("shared-sheet")]
    [InlineData("missing")]
    [InlineData("overlap")]
    public void PpjErrorWorkbookRejectsUnprovedClosure(string variant)
    {
        var (source, chartPath, workbookPath) = ErrorWorkbookSource("line", 0, false);
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        if (variant is "stale" or "formula-cell")
        {
            var book = ZipBytes(source, workbookPath);
            var sheetPath = SingleZipEntryPath(book, path => path.StartsWith("xl/worksheets/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
            book = ReplaceZipText(book, sheetPath, text =>
            {
                var xml = XDocument.Parse(text);
                var cell = xml.Descendants(s + "c").Single(cell => (string?)cell.Attribute("r") == "D2");
                if (variant == "stale") cell.Element(s + "v")!.Value = "9";
                else cell.AddFirst(new XElement(s + "f", "1-1"));
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            source = ReplaceErrorWorkbookBytes(source, workbookPath, book);
        }
        else if (variant == "shared-sheet")
        {
            var book = ReplaceZipText(ZipBytes(source, workbookPath), "xl/workbook.xml", text =>
            {
                var xml = XDocument.Parse(text);
                var sheet = new XElement(xml.Descendants(s + "sheet").Single());
                sheet.SetAttributeValue("name", "Alias");
                sheet.SetAttributeValue("sheetId", "2");
                xml.Root!.Element(s + "sheets")!.Add(sheet);
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            source = ReplaceErrorWorkbookBytes(source, workbookPath, book);
        }
        else if (variant == "shared")
        {
            var slash = chartPath.LastIndexOf('/');
            var relPath = chartPath[..slash] + "/_rels/" + chartPath[(slash + 1)..] + ".rels";
            source = ReplaceZipText(source, relPath, text =>
            {
                var xml = XDocument.Parse(text);
                var duplicate = new XElement(xml.Root!.Elements().Single());
                duplicate.SetAttributeValue("Id", "rIdSharedWorkbook");
                xml.Root.Add(duplicate);
                return xml.ToString(SaveOptions.DisableFormatting);
            });
        }
        else source = ReplaceZipText(source, chartPath, text =>
        {
            var xml = XDocument.Parse(text);
            if (variant == "missing") xml.Root!.Element(c + "externalData")!.Remove();
            else
            {
                var owner = xml.Descendants(c + "ser").Single(series => series.Element(c + "tx")!.Value == "Unchanged").Element(c + "val")!;
                var literal = owner.Element(c + "numLit")!;
                literal.Name = c + "numCache";
                literal.Remove();
                owner.Add(new XElement(c + "numRef", new XElement(c + "f", "'Sheet1'!$D$2:$D$5"), literal));
            }
            return xml.ToString(SaveOptions.DisableFormatting);
        });
        var program = ProjectTrendlineList(source);
        var noOp = CompileTrendlineList(program, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        if (variant == "overlap")
        {
            var otherChannel = ProjectTrendlineList(source);
            ErrorWorkbookSeries(otherChannel, 1)["values"]![0] = 100;
            var otherRejected = CompileTrendlineList(otherChannel, source);
            Assert.False(otherRejected.Ok);
            Assert.Empty(otherRejected.File);
            var imported = Import(source);
            Assert.True(imported.Ok, Diagnostics(imported));
            imported.Artifact.Presentation.Slides.SelectMany(slide => slide.Elements).Single(element => element.Chart is not null).Chart.Series[1].Values[0] = 100;
            var nativeRejected = Export(imported.Artifact);
            Assert.False(nativeRejected.Ok);
            Assert.Empty(nativeRejected.File);
            Assert.Contains("presentation_chart_error_data_binding", Diagnostics(nativeRejected));
        }
        ErrorWorkbookSeries(program, 0)["errorBars"]!["plus"]!["values"]![0] = 2;
        var rejected = CompileTrendlineList(program, source);
        Assert.False(rejected.Ok);
        Assert.Empty(rejected.File);
        Assert.Contains("presentation_chart_error_data_binding", Diagnostics(rejected));
    }

    private static (byte[] Source, string ChartPath, string WorkbookPath) ErrorWorkbookSource(string type, int index, bool grouped)
    {
        var program = TrendlineListProgram(type, index);
        TrendlineListSeries(program, index)["errorBars"] = JsonNode.Parse(CustomErrorData);
        if (grouped)
        {
            var chart = TrendlineListChart(program);
            var parent = chart.Parent!.AsArray();
            var position = parent.IndexOf(chart);
            parent.Remove(chart);
            parent.Insert(position, new JsonObject
            {
                ["id"] = "error-data-group", ["type"] = "group", ["frame"] = chart["frame"]!.DeepClone(),
                ["elements"] = new JsonArray(chart),
            });
        }
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        using var outer = new MemoryStream();
        outer.Write(RemoveEmbeddedPpj(authored.File.ToByteArray()));
        outer.Position = 0;
        string chartPath, workbookPath;
        using (var document = PresentationDocument.Open(outer, true))
        {
            var chartPart = document.PresentationPart!.SlideParts.SelectMany(slide => slide.ChartParts).Single();
            var workbookPart = chartPart.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rIdErrorWorkbook");
            chartPath = chartPart.Uri.OriginalString.TrimStart('/');
            workbookPath = workbookPart.Uri.OriginalString.TrimStart('/');
            using var inner = new MemoryStream();
            using (var book = SpreadsheetDocument.Create(inner, SpreadsheetDocumentType.Workbook))
            {
                var root = book.AddWorkbookPart();
                var sheet = root.AddNewPart<WorksheetPart>();
                root.Workbook = new S.Workbook(new S.Sheets(new S.Sheet { Name = "Sheet1", SheetId = 1, Id = root.GetIdOfPart(sheet) }));
                var errors = JsonNode.Parse(CustomErrorData)!;
                sheet.Worksheet = new S.Worksheet(new S.SheetData(Enumerable.Range(0, 4).Select(point => new S.Row(
                    new S.Cell { CellReference = $"D{point + 2}", CellValue = new S.CellValue(errors["plus"]!["values"]![point]!.GetValue<double>().ToString("R", CultureInfo.InvariantCulture)) },
                    new S.Cell { CellReference = $"E{point + 2}", CellValue = new S.CellValue(errors["minus"]!["values"]![point]!.GetValue<double>().ToString("R", CultureInfo.InvariantCulture)) }) { RowIndex = (uint)(point + 2) })));
            }
            inner.Position = 0;
            workbookPart.FeedData(inner);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XDocument xml;
            using (var input = chartPart.GetStream(FileMode.Open, FileAccess.Read)) xml = XDocument.Load(input);
            foreach (var (side, column) in new[] { ("plus", "D"), ("minus", "E") })
            {
                var owner = xml.Descendants(c + "errBars").Single().Element(c + side)!;
                var literal = owner.Element(c + "numLit")!;
                literal.Name = c + "numCache";
                literal.Remove();
                owner.Add(new XElement(c + "numRef", new XElement(c + "f", $"'Sheet1'!${column}$2:${column}$5"), literal));
            }
            xml.Root!.Add(new XElement(c + "externalData", new XAttribute(r + "id", "rIdErrorWorkbook"), new XElement(c + "autoUpdate", new XAttribute("val", 0))));
            using var output = chartPart.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = System.Xml.XmlWriter.Create(output, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false });
            xml.Save(writer);
        }
        return (RemoveEmbeddedPpj(outer.ToArray()), chartPath, workbookPath);
    }

    private static JsonObject ErrorWorkbookSeries(JsonObject program, int index)
    {
        IEnumerable<JsonObject> Descendants(JsonArray elements)
        {
            foreach (var node in elements)
            {
                var element = node!.AsObject();
                yield return element;
                if (element["type"]!.GetValue<string>() == "group")
                    foreach (var child in Descendants(element["elements"]!.AsArray())) yield return child;
            }
        }
        var chart = program["pages"]!.AsArray().SelectMany(page => Descendants(page!["elements"]!.AsArray()))
            .Single(element => element["type"]!.GetValue<string>() == "chart");
        return chart["data"]!["series"]![index]!.AsObject();
    }

    private static byte[] ReplaceErrorWorkbookBytes(byte[] source, string part, byte[] bytes)
    {
        using var memory = new MemoryStream();
        memory.Write(source);
        memory.Position = 0;
        using (var archive = new ZipArchive(memory, ZipArchiveMode.Update, true))
        {
            archive.GetEntry(part)!.Delete();
            AddZipBytes(archive, part, bytes);
        }
        return memory.ToArray();
    }
}
