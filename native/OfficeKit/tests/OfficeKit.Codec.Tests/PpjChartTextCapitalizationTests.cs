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
    public void PpjChartTextCapitalizationPreservesCancellationDeletionAndRecreation(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartCapitalizations(program, seriesIndex, "all");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        Assert.Contains("Language chart", originalText);
        Assert.Contains("Fit ", originalText);
        AssertChartCapitalizations(ProjectTrendlineList(source), seriesIndex, "all");
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());

        foreach (var requested in new string?[] { "small", "none", null, "all" })
        {
            var projected = ProjectTrendlineList(source);
            SetChartCapitalizations(projected, seriesIndex, requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var expected = requested;
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var attributes = native.Descendants().Attributes("cap").ToArray();
            if (expected is null) Assert.Empty(attributes);
            else { Assert.NotEmpty(attributes); Assert.All(attributes, attribute => Assert.Equal(expected, attribute.Value)); }
            AssertChartCapitalizations(ProjectTrendlineList(output), seriesIndex, expected);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextCapitalizationRejectsInvalidValuesAndGlobalProfile()
    {
        foreach (var value in new[] { "", "uppercase", "ALL", "none " })
        {
            var style = new SpreadsheetChartTextStyleArtifact { Capitalization = value };
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(style, "s", "c", "style"));
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("cap", value);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
            var program = ChartTextStyleProgram("line", 0);
            SetChartCapitalizations(program, 0, value);
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { Capitalization = "none" };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }), XlsxChartTextStyleCodec.Semantics(valid));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
    }

    [Fact]
    public void PpjChartTextCapitalizationHonorsPrecedenceAndVectorText()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "capitalization-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["capitalization"] = "none", ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"capitalization-style","style":{"legendTextStyle":{"capitalization":"all"}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.capitalization","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("all", projected["style"]!["legendTextStyle"]!["capitalization"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());

        chart.Remove("styleRef");
        chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"capitalization":"all"},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"capitalization":"small"}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"capitalization":"none"}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, capitalization) in new[] { ("Default", "all"), ("Override", "none"), ("Row", "small") })
            Assert.Equal(capitalization, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Attribute("cap"));
    }

    private static void SetChartCapitalizations(JsonObject program, int index, string? requested)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (requested is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["capitalization"] = requested };
    }

    private static void AssertChartCapitalizations(JsonObject program, int index, string? capitalization)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(capitalization, owner[field]?["capitalization"]?.GetValue<string>());
    }
}
