using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjChartTextLanguagePreservesEveryOwnerAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartLanguages(program, seriesIndex, "en-US");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        AssertChartLanguages(ProjectTrendlineList(source), seriesIndex, "en-US");
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());

        // Freshly project each original candidate, including the deletion case.
        foreach (var language in new string?[] { "zh-Hans-CN", null, "ar-SA", "en-US" })
        {
            var projected = ProjectTrendlineList(source);
            SetChartLanguages(projected, seriesIndex, language);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var tags = native.Descendants().Attributes("lang").ToArray();
            if (language is null) Assert.Empty(tags);
            else { Assert.NotEmpty(tags); Assert.All(tags, tag => Assert.Equal(language, tag.Value)); }
            AssertChartLanguages(ProjectTrendlineList(output), seriesIndex, language);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextLanguageRejectsInvalidTagsAndKeepsGlobalProfileRestricted()
    {
        foreach (var language in new[] { "", "en_US", " en-US", "en-", new string('a', 64) })
        {
            var style = new SpreadsheetChartTextStyleArtifact { Language = language };
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(style, "s", "c", "style"));
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("lang", language);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
            var program = TrendlineListProgram("line", 0);
            TrendlineListChart(program)["title"] = "Title";
            TrendlineListChart(program)["style"]!["titleTextStyle"] = new JsonObject { ["language"] = language };
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { Language = "EN-us" };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(null), XlsxChartTextStyleCodec.Semantics(valid));
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }),
            XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", Language = "default-language" }));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
        var sourceProgram = TrendlineListProgram("line", 0);
        TrendlineListChart(sourceProgram)["title"] = "Title";
        TrendlineListChart(sourceProgram)["style"]!["titleTextStyle"] = new JsonObject { ["language"] = "en-US" };
        var sourceCompiled = CompileTrendlineList(sourceProgram);
        Assert.True(sourceCompiled.Ok, Diagnostics(sourceCompiled));
        var source = RemoveEmbeddedPpj(sourceCompiled.File.ToByteArray());
        foreach (var token in new[] { "{\"kind\":\"size\",\"value\":12}", "{\"kind\":\"string\",\"value\":\"en_US\"}" })
        {
            var program = TrendlineListProgram("line", 0);
            var chart = TrendlineListChart(program);
            chart["title"] = "Title";
            chart["style"]!["titleTextStyle"] = JsonNode.Parse("""{"language":{"token":"invalid"}}""");
            program["design"]!["grammar"]!["tokens"] = new JsonObject { ["invalid"] = JsonNode.Parse(token) };
            Assert.False(CompileTrendlineList(program).Ok);
            var projected = ProjectTrendlineList(source);
            TrendlineListChart(projected)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"language":{"token":"invalid"}}""");
            projected["design"]!["grammar"]!["tokens"] = new JsonObject { ["invalid"] = JsonNode.Parse(token) };
            Assert.False(CompileTrendlineList(projected, source).Ok);
        }
    }

    [Fact]
    public void PpjChartTextLanguageHonorsFieldPrecedenceAndVectorRunOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "language-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["language"] = "en-US", ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"language-style","style":{"legendTextStyle":{"language":"ja-JP"}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.language","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("ja-JP", projected["style"]!["legendTextStyle"]!["language"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());

        chart.Remove("styleRef");
        chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"axisLanguage":{"kind":"string","value":"ja-JP"}}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"language":"zh-CN"},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"language":{"token":"axisLanguage"}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"language":"ar-SA"}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        Assert.Equal("zh-CN", (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == "Default")).Element(a + "rPr")?.Attribute("lang"));
        Assert.Equal("ar-SA", (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == "Override")).Element(a + "rPr")?.Attribute("lang"));
        Assert.Equal("ja-JP", (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == "Row")).Element(a + "rPr")?.Attribute("lang"));
    }

    private static JsonObject ChartTextStyleProgram(string type, int seriesIndex)
    {
        var program = TrendlineListProgram(type, seriesIndex);
        var chart = TrendlineListChart(program);
        chart["title"] = "Language chart";
        chart["style"]!["dataLabels"] = new JsonObject { ["showValue"] = true };
        foreach (var axis in type == "combo" ? new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" } : new[] { "xAxis", "yAxis" })
            chart[axis] = new JsonObject { ["title"] = axis };
        TrendlineListSeries(program, seriesIndex)["trendlines"]![0]!["label"] = JsonNode.Parse("""
            {"text":{"paragraphs":[{"runs":[{"text":"Fit "},{"break":true},{"text":"A"}]}]}}
            """);
        return program;
    }

    private static IEnumerable<(JsonObject Owner, string Field)> ChartTextStyleOwners(JsonObject program, int index)
    {
        var chart = TrendlineListChart(program);
        var style = chart["style"]!.AsObject();
        yield return (style, "titleTextStyle");
        yield return (style, "legendTextStyle");
        yield return (style["dataLabels"]!.AsObject(), "textStyle");
        foreach (var name in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (chart[name] is JsonObject axis) { yield return (axis, "textStyle"); yield return (axis, "titleTextStyle"); }
        var label = TrendlineListSeries(program, index)["trendlines"]![0]!["label"]!.AsObject();
        yield return (label, "textStyle");
        var paragraph = label["text"]!["paragraphs"]![0]!.AsObject();
        yield return (paragraph, "style");
        yield return (paragraph, "endStyle");
        foreach (var run in paragraph["runs"]!.AsArray()) yield return (run!.AsObject(), "style");
    }

    private static void SetChartLanguages(JsonObject program, int index, string? language)
    {
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        { ["chartLanguage"] = new JsonObject { ["kind"] = "string", ["value"] = language ?? "en-US" } };
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (language is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["language"] = new JsonObject { ["token"] = "chartLanguage" } };
    }

    private static void AssertChartLanguages(JsonObject program, int index, string? language)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(language, owner[field]?["language"]?.GetValue<string>());
    }
}
