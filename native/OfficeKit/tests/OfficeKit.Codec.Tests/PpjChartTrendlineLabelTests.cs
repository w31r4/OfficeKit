using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    private const string TrendlineLabelJson = """
        {"text":"Fit","numberFormat":"0.00","textStyle":{"fontSize":11,"bold":true,"color":"#2563EB"},"fill":{"type":"solid","color":"#FEF3C7"},"line":{"color":"#D97706","width":1,"dash":"solid"}}
        """;

    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineLabelAuthorsEditsRemovesAndReprojects(string chartType, int seriesIndex)
    {
        var program = TrendlineListProgram(chartType, seriesIndex);
        var trendline = TrendlineListSeries(program, seriesIndex)["trendlines"]![0]!;
        trendline["displayEquation"] = true;
        trendline["displayRSquared"] = true;
        trendline["label"] = JsonNode.Parse(TrendlineLabelJson);
        SetTrendlineLabelTokens(program, seriesIndex, "Fit", "0.00");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectTrendlineList(source);
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse(TrendlineLabelJson), TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]));
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var requested in new string?[]
        {
            "{\"text\":\"Adjusted\",\"numberFormat\":\"0.0%\",\"textStyle\":{\"fontSize\":14,\"bold\":false},\"fill\":{\"type\":\"none\"},\"line\":{\"color\":\"#166534\",\"width\":2,\"dash\":\"dash\"}}",
            "{}", null, TrendlineLabelJson,
        })
        {
            var item = TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!.AsObject();
            if (requested is null) item.Remove("label");
            else item["label"] = JsonNode.Parse(requested);
            if (requested == TrendlineLabelJson) SetTrendlineLabelTokens(projected, seriesIndex, "Fit", "0.00");
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var before = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var after = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var native = after.Descendants(c + "ser").Single(series => series.Element(c + "tx")!.Value == "Editable").Elements(c + "trendline").First();
            Assert.Equal("1", (string?)native.Element(c + "dispEq")!.Attribute("val"));
            Assert.Equal("1", (string?)native.Element(c + "dispRSqr")!.Attribute("val"));
            Assert.Equal(requested is not null, native.Element(c + "trendlineLbl") is not null);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(new C.Trendline(native.ToString(SaveOptions.DisableFormatting))));
            before.Descendants(c + "trendlineLbl").Remove();
            after.Descendants(c + "trendlineLbl").Remove();
            Assert.True(XNode.DeepEquals(before, after), "Only label state may change inside ChartML.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            Assert.True(JsonNode.DeepEquals(requested is null ? null : JsonNode.Parse(requested), TrendlineListSeries(projected, seriesIndex)["trendlines"]![0]!["label"]));
        }
    }

    [Fact]
    public void PpjTrendlineLabelRejectsInvalidValuesAndPreservesUnsupportedOwners()
    {
        var program = TrendlineListProgram("line", 0);
        TrendlineListSeries(program, 0)["trendlines"]![0]!["label"] = JsonNode.Parse(TrendlineLabelJson);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        foreach (var invalid in new[] { "{\"text\":\"\"}", "{\"numberFormat\":\"\"}", "{\"position\":\"top\"}", "{\"text\":\"bad\\nlabel\"}",
            "{\"layout\":null}", "{\"layout\":{\"manual\":null}}", "{\"layout\":{\"manual\":{\"xMode\":\"invalid\"}}}",
            "{\"layout\":{\"manual\":{\"target\":\"invalid\"}}}", "{\"layout\":{\"manual\":{\"x\":\"0\"}}}",
            "{\"layout\":{\"manual\":{\"x\":1e999}}}", "{\"layout\":{\"manual\":{\"rotation\":1}}}",
            "{\"numberFormatSourceLinked\":true}", "{\"numberFormatSourceLinked\":false}", "{\"numberFormatSourceLinked\":null}",
            "{\"numberFormat\":\"0\",\"numberFormatSourceLinked\":\"true\"}", "{\"numberFormat\":\"0\",\"numberFormatSourceLinked\":1}",
            """{"text":{"paragraphs":[]}}""", """{"text":{"paragraphs":[{"runs":[{}]}]}}""",
            """{"text":{"paragraphs":[{"runs":[{"break":false}]}]}}""", """{"text":{"paragraphs":[{"runs":[{"text":"x","break":true}]}]}}""",
            """{"text":{"paragraphs":[{"runs":[{"text":"x","style":{"alignment":"center"}}]}]}}""",
            """{"text":{"paragraphs":[{"runs":[{"text":"bad\nrun"}]}]}}""" })
        {
            var input = program.DeepClone().AsObject();
            TrendlineListSeries(input, 0)["trendlines"]![0]!["label"] = JsonNode.Parse(invalid);
            Assert.False(CompileTrendlineList(input).Ok);
            input = ProjectTrendlineList(source);
            TrendlineListSeries(input, 0)["trendlines"]![0]!["label"] = JsonNode.Parse(invalid);
            var rejected = CompileTrendlineList(input, source);
            Assert.False(rejected.Ok);
            Assert.Empty(rejected.File);
        }
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        string Rich(string content) => $"<c:tx><c:rich xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>{content}</c:rich></c:tx>";
        foreach (var invalid in new[]
        {
            "<c:layout><c:manualLayout><c:extLst/></c:manualLayout></c:layout>",
            "<c:layout><c:extLst/></c:layout>",
            "<c:layout><c:manualLayout extra='1'/></c:layout>",
            "<c:layout><c:manualLayout><c:x val='0'/><c:x val='1'/></c:manualLayout></c:layout>",
            "<c:layout><c:manualLayout><c:x val='0'/><c:xMode val='edge'/></c:manualLayout></c:layout>",
            "<c:layout><c:manualLayout><c:x val='NaN'/></c:manualLayout></c:layout>",
            "<c:layout><c:manualLayout><c:x/></c:manualLayout></c:layout>",
            "<c:layout><c:manualLayout><c:xMode val='invalid'/></c:manualLayout></c:layout>",
            "<c:layout><c:manualLayout>unknown</c:manualLayout></c:layout>",
            "<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>",
            "<c:tx>unmodeled<c:rich xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Fit</a:t></a:r></a:p></c:rich></c:tx>",
            Rich("<a:bodyPr rot='60000'/><a:lstStyle/><a:p/>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/><a:defRPr/></a:pPr></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick/></a:rPr><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang='en_US'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr strike='single'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr cap='uppercase'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr spc='76801'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr baseline='400001'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr strike='sngStrike' baseline='100' cap='all' spc='100' kern='100'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang='en-US' altLang='zh-CN'/><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/><a:r><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b='1'>unmodeled</a:rPr><a:t>Fit</a:t></a:r></a:p>"),
            Rich("<a:bodyPr/><a:lstStyle/><a:p><a:fld id='field' type='slidenum'><a:t>1</a:t></a:fld></a:p>"),
            "<c:numFmt formatCode='0' sourceLinked='invalid'/>",
            "<c:numFmt formatCode='0' sourceLinked='0'/><c:numFmt formatCode='0.0' sourceLinked='0'/>",
            "<c:extLst/>",
        })
        {
            var input = ReplaceZipText(source, path, text =>
            {
                var xml = XDocument.Parse(text);
                xml.Descendants(c + "trendlineLbl").Single().ReplaceNodes(XElement.Parse($"<root xmlns:c='{c}'>{invalid}</root>").Elements());
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var projected = ProjectTrendlineList(input);
            var chart = TrendlineListChart(projected);
            Assert.DoesNotContain(chart["nativeRef"]!["capabilities"]!.AsArray(), item => item!["operation"]!.GetValue<string>() == "setChartSeriesAnalytics");
            var noOp = CompileTrendlineList(projected, input);
            Assert.True(noOp.Ok, Diagnostics(noOp));
            Assert.Equal(input, noOp.File.ToByteArray());
            TrendlineListSeries(projected, 0)["trendlines"] = new JsonArray(new JsonObject { ["type"] = "linear", ["label"] = new JsonObject() });
            var rejected = CompileTrendlineList(projected, input);
            Assert.False(rejected.Ok);
            Assert.Empty(rejected.File);
        }
    }

    private static void SetTrendlineLabelTokens(JsonObject program, int index, string text, string format)
    {
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["labelText"] = new JsonObject { ["kind"] = "string", ["value"] = text },
            ["labelFormat"] = new JsonObject { ["kind"] = "string", ["value"] = format },
        };
        var label = TrendlineListSeries(program, index)["trendlines"]![0]!["label"]!;
        label["text"] = new JsonObject { ["token"] = "labelText" };
        label["numberFormat"] = new JsonObject { ["token"] = "labelFormat" };
    }
}
