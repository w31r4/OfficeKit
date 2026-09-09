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
    public void PpjChartTextHighlightPreservesPaintTokensAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        TrendlineListChart(program)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":"#112233","fontFamily":"Arial"}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#ffee00"}}""");
        SetChartHighlights(program, seriesIndex, JsonNode.Parse("""{"token":"mark"}"""));
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        Assert.Contains("Language chart", originalText);
        Assert.Contains("Fit ", originalText);
        AssertChartHighlights(ProjectTrendlineList(source), seriesIndex, "#FFEE00");
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (input, expected) in new (string? Input, string? Expected)[]
        {
            ("\"#00ff00\"", "#00FF00"),
            ("""{"token":"mark","tint":1,"shade":0.5}""", "#808080"),
            (null, null), ("\"#000000\"", "#000000"),
        })
        {
            var projected = ProjectTrendlineList(source);
            projected["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#000000"}}""");
            SetChartHighlights(projected, seriesIndex, input is null ? null : JsonNode.Parse(input));
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var highlights = native.Descendants(a + "highlight").ToArray();
            if (expected is null) Assert.Empty(highlights);
            else { Assert.NotEmpty(highlights); Assert.All(highlights, value => Assert.Equal(expected[1..], (string?)Assert.Single(value.Elements()).Attribute("val"))); }
            var fresh = ProjectTrendlineList(output);
            AssertChartHighlights(fresh, seriesIndex, expected);
            var titleStyle = TrendlineListChart(fresh)["style"]!["titleTextStyle"]!;
            Assert.Equal("#112233", titleStyle["color"]!.GetValue<string>());
            Assert.Equal("Arial", titleStyle["fontFamily"]!.GetValue<string>());
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextHighlightValidatesNativeOrderAndRejectsInvalidPaint()
    {
        foreach (var rgb in new[] { "", "FFF", "#FFFFFF", "GGGGGG", "FFFFFF80" })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { HighlightRgb = rgb }, "s", "c", "style"));
        var valid = new SpreadsheetChartTextStyleArtifact { HighlightRgb = "FFAACC" };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var highlightOnly));
        Assert.Equal(valid, highlightOnly);
        valid.ColorRgb = "112233"; valid.FontFamily = "Arial"; valid.FontFamilyEastAsia = "SimSun";
        var native = XlsxChartTextStyleCodec.StyleProperties("rPr", valid);
        Assert.Equal(new[] { "solidFill", "highlight", "latin", "ea" }, native.Elements().Select(element => element.Name.LocalName));
        var properties = new DocumentFormat.OpenXml.Drawing.RunProperties(native.ToString(SaveOptions.DisableFormatting));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021).Validate(properties));
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(native, out var restored));
        Assert.Equal(valid, restored);
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", HighlightRgb = "FFFFFF" }));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(highlightOnly), XlsxChartTextStyleCodec.Semantics(null));
        Assert.Equal(XlsxChartTextStyleCodec.Semantics(highlightOnly), XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { HighlightRgb = "ffaacc" }));

        var baseline = ChartTextStyleProgram("line", 0);
        SetChartHighlights(baseline, 0, JsonValue.Create("#FFFF00"));
        var compiled = CompileTrendlineList(baseline);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var source = RemoveEmbeddedPpj(compiled.File.ToByteArray());
        foreach (var input in new[] { "\"#FFFFFF80\"", """{"token":"mark","alpha":0.5}""", """{"token":"missing"}""", """{"token":"wrongKind"}""" })
        {
            foreach (var sourceBound in new[] { false, true })
            {
                var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
                program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#FFFFFF"},"wrongKind":{"kind":"string","value":"#FFFFFF"}}""");
                SetChartHighlights(program, 0, JsonNode.Parse(input));
                Assert.False(CompileTrendlineList(program, sourceBound ? source : null).Ok);
            }
        }
    }

    [Fact]
    public void PpjChartTextHighlightHonorsPrecedenceAndVectorOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "highlight-style";
        chart["style"]!["legendTextStyle"] = JsonNode.Parse("""{"highlight":"#FFFFFF","bold":false}""");
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""{"id":"highlight-style","style":{"legendTextStyle":{"highlight":"#FF0000"}}}"""));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""[{"target":"chart.legendTextStyle.highlight","sources":["styleRef","inline"]}]""");
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("#FF0000", projected["style"]!["legendTextStyle"]!["highlight"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#00FF00"}}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"highlight":"#FFFF00"},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"highlight":{"token":"mark"}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"highlight":"#000000"}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, rgb) in new[] { ("Default", "FFFF00"), ("Override", "000000"), ("Row", "00FF00") })
            Assert.Equal(rgb, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Element(a + "highlight")?.Element(a + "srgbClr")?.Attribute("val"));
        chart["style"]!["heatmap"]!["axisTextStyle"]!["highlight"] = "#FFFFFF80";
        Assert.False(CompileTrendlineList(program).Ok);
    }

    private static void SetChartHighlights(JsonObject program, int index, JsonNode? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
        {
            var style = owner[field]?.AsObject();
            if (value is null)
            {
                style?.Remove("highlight");
                if (style?.Count == 0) owner.Remove(field);
            }
            else
            {
                if (style is null) owner[field] = style = new JsonObject();
                style["highlight"] = value.DeepClone();
            }
        }
    }

    private static void AssertChartHighlights(JsonObject program, int index, string? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(value, owner[field]?["highlight"]?.GetValue<string>());
    }
}
