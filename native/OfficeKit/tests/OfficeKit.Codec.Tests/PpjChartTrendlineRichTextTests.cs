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
    private const string TrendlineRichTextJson = """
        {"paragraphs":[{"style":{"alignment":"center","fontSize":12,"bold":true},"runs":[{"text":"Fit ","style":{"color":"#2563EB"}},{"text":"A","style":{"bold":false}},{"break":true,"style":{"italic":true}},{"text":""}],"endStyle":{"italic":false}},{"runs":[{"text":"Next","style":{"fontFamilyEastAsia":"Noto Sans CJK SC"}}]},{"runs":[]}]}
        """;

    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineRichTextAuthorsEditsRemovesAndReprojects(string chartType, int seriesIndex)
    {
        var program = TrendlineListProgram(chartType, seriesIndex);
        var label = JsonNode.Parse(TrendlineLabelJson)!.AsObject();
        label["text"] = JsonNode.Parse(TrendlineRichTextJson);
        TrendlineListSeries(program, seriesIndex)["trendlines"]![0]!["label"] = label;
        SetTrendlineRichTextTokens(program, seriesIndex);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var nativeText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(c + "trendlineLbl").Single().Element(c + "tx")!;
        Assert.Equal(new[] { "Fit ", "A", "", "Next" }, nativeText.Descendants(a + "t").Select(t => t.Value));
        Assert.Single(nativeText.Descendants(a + "br"));
        Assert.Equal("0", (string?)nativeText.Descendants(a + "rPr").Single(r => r.Attribute("b") is not null).Attribute("b"));
        var projected = ProjectTrendlineList(source);
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse(TrendlineRichTextJson), TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!["text"]));
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var requested in new string?[]
        {
            TrendlineRichTextJson.Replace("Next", "Edited", StringComparison.Ordinal),
            "\"Literal\"", "{\"paragraphs\":[{\"runs\":[{\"text\":\"Simple\"}]}]}",
            "{\"paragraphs\":[{\"runs\":[]}]}", null, TrendlineRichTextJson,
        })
        {
            var item = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!.AsObject();
            if (requested is null) item.Remove("text");
            else item["text"] = JsonNode.Parse(requested);
            if (requested == TrendlineRichTextJson) SetTrendlineRichTextTokens(projected, seriesIndex);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var before = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var after = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var native = after.Descendants(c + "trendlineLbl").Single();
            Assert.Equal(requested is not null, native.Element(c + "tx") is not null);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(new C.TrendlineLabel(native.ToString(SaveOptions.DisableFormatting))));
            before.Descendants(c + "trendlineLbl").Elements(c + "tx").Remove();
            after.Descendants(c + "trendlineLbl").Elements(c + "tx").Remove();
            Assert.True(XNode.DeepEquals(before, after), "Only trendline label text may change.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            var expected = requested?.Contains("Simple", StringComparison.Ordinal) == true ? JsonValue.Create("Simple") : requested is null ? null : JsonNode.Parse(requested);
            Assert.True(JsonNode.DeepEquals(expected, TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]!["text"]));
        }
    }

    [Fact]
    public void PpjTrendlineRichTextRejectsInvalidWireAndBudgets()
    {
        var text = new SpreadsheetChartRichTextArtifact();
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
        var paragraph = new SpreadsheetChartTextParagraphArtifact();
        text.Paragraphs.Add(paragraph);
        var run = new SpreadsheetChartTextRunArtifact { LineBreak = false };
        paragraph.Runs.Add(run);
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
        run.Text = "valid";
        run.Style = new SpreadsheetChartTextStyleArtifact { Alignment = "ctr" };
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
        run.Style = null;
        Assert.Throws<CodecException>(() => OpenXmlChartTrendlineLabelCodec.Validate(new SpreadsheetChartTrendlineLabelArtifact { Text = "both", RichText = text }, "s", "c", "v"));
        run.Text = new string('x', 1_048_577);
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
        run.Text = "";
        for (var i = 0; i < 16_384; i++) paragraph.Runs.Add(new SpreadsheetChartTextRunArtifact { Text = "" });
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
        paragraph.Runs.Clear();
        for (var i = 0; i < 4096; i++) text.Paragraphs.Add(new SpreadsheetChartTextParagraphArtifact());
        Assert.Throws<CodecException>(() => OpenXmlChartRichTextCodec.Validate(text));
    }

    private static void SetTrendlineRichTextTokens(JsonObject program, int index)
    {
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["richRun"] = new JsonObject { ["kind"] = "string", ["value"] = "A" },
            ["richEmpty"] = new JsonObject { ["kind"] = "string", ["value"] = "" },
            ["richSize"] = new JsonObject { ["kind"] = "size", ["value"] = 12 },
        };
        var p = TrendlineListSeries(program, index)["trendlines"]![0]!["label"]!["text"]!["paragraphs"]![0]!;
        p["runs"]![1]!["text"] = new JsonObject { ["token"] = "richRun" };
        p["runs"]![3]!["text"] = new JsonObject { ["token"] = "richEmpty" };
        p["style"]!["fontSize"] = new JsonObject { ["token"] = "richSize" };
    }
}
