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
    public void PpjChartTextStrikePreservesCancellationDeletionAndRecreation(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartStrikes(program, seriesIndex, "true");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        AssertChartStrikes(ProjectTrendlineList(source), seriesIndex, "sngStrike");
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());

        foreach (var requested in new string?[] { "\"dblStrike\"", "false", null, "\"sngStrike\"", "\"noStrike\"" })
        {
            var projected = ProjectTrendlineList(source);
            SetChartStrikes(projected, seriesIndex, requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var expected = requested is null ? null : requested == "false" ? "noStrike" : JsonNode.Parse(requested)!.GetValue<string>();
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var attributes = native.Descendants().Attributes("strike").ToArray();
            if (expected is null) Assert.Empty(attributes);
            else { Assert.NotEmpty(attributes); Assert.All(attributes, attribute => Assert.Equal(expected, attribute.Value)); }
            AssertChartStrikes(ProjectTrendlineList(output), seriesIndex, expected);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextStrikeRejectsInvalidValuesAndGlobalProfile()
    {
        foreach (var value in new[] { "", "single", "none", "noStrike " })
        {
            var style = new SpreadsheetChartTextStyleArtifact { Strike = value };
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(style, "s", "c", "style"));
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("strike", value);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
            var program = ChartTextStyleProgram("line", 0);
            SetChartStrikes(program, 0, JsonValue.Create(value)!.ToJsonString());
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { Strike = "noStrike" };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }), XlsxChartTextStyleCodec.Semantics(valid));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
    }

    [Fact]
    public void PpjChartTextStrikeHonorsPrecedenceAndVectorText()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "strike-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["strike"] = false, ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"strike-style","style":{"legendTextStyle":{"strike":true}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.strike","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("sngStrike", projected["style"]!["legendTextStyle"]!["strike"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());

        chart.Remove("styleRef");
        chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"strike":true},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"strike":false}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"strike":"dblStrike"}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, strike) in new[] { ("Default", "sngStrike"), ("Override", "dblStrike"), ("Row", "noStrike") })
            Assert.Equal(strike, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Attribute("strike"));
    }

    private static void SetChartStrikes(JsonObject program, int index, string? requested)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (requested is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["strike"] = JsonNode.Parse(requested) };
    }

    private static void AssertChartStrikes(JsonObject program, int index, string? strike)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(strike, owner[field]?["strike"]?.GetValue<string>());
    }
}
