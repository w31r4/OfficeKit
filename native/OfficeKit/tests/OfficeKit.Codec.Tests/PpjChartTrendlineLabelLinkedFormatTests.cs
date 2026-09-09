using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using OfficeKit.Artifact.Wire.V1;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineLabelLinkedFormatAuthorsEditsRemovesAndReprojects(string chartType, int seriesIndex)
    {
        var program = TrendlineListProgram(chartType, seriesIndex);
        var label = JsonNode.Parse(TrendlineLabelJson)!.AsObject();
        label["numberFormatSourceLinked"] = true;
        TrendlineListSeries(program, seriesIndex)["trendlines"]![0]!["label"] = label;
        SetTrendlineLabelTokens(program, seriesIndex, "Fit", "0.00");
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        byte[] source = [];
        string path = "";
        foreach (var state in new[] { "false", "null", "true" })
        {
            label["numberFormatSourceLinked"] = JsonNode.Parse(state);
            var authored = CompileTrendlineList(program);
            Assert.True(authored.Ok, Diagnostics(authored));
            source = RemoveEmbeddedPpj(authored.File.ToByteArray());
            path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(c + "trendlineLbl").Single().Element(c + "numFmt")!;
            Assert.Equal("0.00", (string?)native.Attribute("formatCode"));
            Assert.Equal(state == "null" ? null : state == "true" ? "1" : "0", (string?)native.Attribute("sourceLinked"));
            var fresh = TrendlineListSeries(ProjectTrendlineList(source), seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            Assert.Equal(state != "false", fresh.ContainsKey("numberFormatSourceLinked"));
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse(state == "false" ? "null" : state), fresh["numberFormatSourceLinked"]));
        }
        // Also import the other valid lexical form instead of only our writer's 1.
        source = ReplaceZipText(source, path, text => text.Replace("sourceLinked=\"1\"", "sourceLinked=\"true\"", StringComparison.Ordinal));
        var projected = ProjectTrendlineList(source);
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var state in new[] { "false", "null", "text-edit", "omit", "remove", "true" })
        {
            var item = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            if (state == "text-edit") item["text"] = "Updated fit";
            else if (state is "remove" or "omit") item.Remove("numberFormatSourceLinked");
            else item["numberFormatSourceLinked"] = JsonNode.Parse(state);
            if (state == "remove") item.Remove("numberFormat");
            if (state == "true")
            {
                item["numberFormat"] = "0.0%";
                SetTrendlineLabelTokens(projected, seriesIndex, "Updated fit", "0.0%");
            }
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var before = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var after = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var native = after.Descendants(c + "trendlineLbl").Single();
            var format = native.Element(c + "numFmt");
            Assert.Equal(state != "remove", format is not null);
            Assert.Equal(state is "null" or "text-edit" or "remove" ? null : state == "true" ? "1" : "0", (string?)format?.Attribute("sourceLinked"));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(new C.TrendlineLabel(native.ToString(SaveOptions.DisableFormatting))));
            before.Descendants(c + "trendlineLbl").Elements(c + "numFmt").Remove();
            after.Descendants(c + "trendlineLbl").Elements(c + "numFmt").Remove();
            if (state == "text-edit")
            {
                Assert.Equal("Updated fit", native.Element(c + "tx")!.Value);
                before.Descendants(c + "trendlineLbl").Elements(c + "tx").Remove();
                after.Descendants(c + "trendlineLbl").Elements(c + "tx").Remove();
            }
            Assert.True(XNode.DeepEquals(before, after), "Only requested label format/text state may change.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            var actual = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            Assert.Equal(state is "null" or "text-edit" or "true", actual.ContainsKey("numberFormatSourceLinked"));
            Assert.Equal(state == "true", actual["numberFormatSourceLinked"]?.GetValue<bool>() == true);
            Assert.Equal(state == "remove" ? null : state == "true" ? "0.0%" : "0.00", actual["numberFormat"]?.GetValue<string>());
        }
    }

    [Fact]
    public void PpjTrendlineLabelLinkedFormatValidatesWireAndRetainsLegacyDefault()
    {
        var legacy = new SpreadsheetChartTrendlineLabelArtifact { NumberFormatCode = "0" };
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        Assert.Equal("0", (string?)OpenXmlChartTrendlineLabelCodec.Element(legacy)!.Element(c + "numFmt")!.Attribute("sourceLinked"));
        foreach (var link in new[] { SpreadsheetChartNumberFormatLink.Source, SpreadsheetChartNumberFormatLink.Omitted, (SpreadsheetChartNumberFormatLink)99 })
            Assert.Throws<CodecException>(() => OpenXmlChartTrendlineLabelCodec.Validate(new SpreadsheetChartTrendlineLabelArtifact { NumberFormatLink = link }, "sheet", "chart", "series"));
        legacy.NumberFormatLink = (SpreadsheetChartNumberFormatLink)99;
        Assert.Throws<CodecException>(() => OpenXmlChartTrendlineLabelCodec.Validate(legacy, "sheet", "chart", "series"));
    }
}
