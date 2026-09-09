using System.Globalization;
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
    public void PpjChartTextLetterSpacingPreservesSignedZeroAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartLetterSpacings(program, seriesIndex, 30);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        Assert.Contains("Language chart", originalText);
        Assert.Contains("Fit ", originalText);
        AssertChartLetterSpacings(ProjectTrendlineList(source), seriesIndex, 30);
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (requested, expected) in new (double? Input, double? Expected)[]
        { (-2.375, -2.38), (0, 0), (null, null), (1.23456, 1.23) })
        {
            var projected = ProjectTrendlineList(source);
            SetChartLetterSpacings(projected, seriesIndex, requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var attributes = native.Descendants().Attributes("spc").ToArray();
            if (expected is null) Assert.Empty(attributes);
            else
            {
                Assert.NotEmpty(attributes);
                Assert.All(attributes, attribute => Assert.Equal(checked((int)Math.Round(expected.Value * 100)), int.Parse(attribute.Value, CultureInfo.InvariantCulture)));
            }
            AssertChartLetterSpacings(ProjectTrendlineList(output), seriesIndex, expected);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextLetterSpacingValidatesNativeRangePrecisionAndPresence()
    {
        foreach (var value in new[] { -768.001, 768.001, double.NaN, double.PositiveInfinity })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.LetterSpacingHundredthPoints(value));
        foreach (var value in new[] { -76_801, 76_801 })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { LetterSpacingHundredthPoints = value }, "s", "c", "style"));
        foreach (var value in new[] { "", "1.5", "76801", "-76801" })
        {
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("spc", value);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
        }
        foreach (var (points, native) in new[] { (-768d, -76800), (768d, 76800), (0.005, 0), (0.015, 2), (-0.015, -2) })
            Assert.Equal(native, XlsxChartTextStyleCodec.LetterSpacingHundredthPoints(points));
        foreach (var points in new[] { -768.001, 768.001 })
        {
            var program = ChartTextStyleProgram("line", 0);
            SetChartLetterSpacings(program, 0, points);
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { LetterSpacingHundredthPoints = 0 };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }), XlsxChartTextStyleCodec.Semantics(valid));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
    }

    [Fact]
    public void PpjChartTextLetterSpacingHonorsPrecedenceAndVectorText()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "letterSpacing-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["letterSpacing"] = 0, ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"letterSpacing-style","style":{"legendTextStyle":{"letterSpacing":30}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.letterSpacing","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal(30, projected["style"]!["legendTextStyle"]!["letterSpacing"]!.GetValue<double>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"letterSpacing":30},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"letterSpacing":0}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"letterSpacing":-25}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, letterSpacing) in new[] { ("Default", "3000"), ("Override", "-2500"), ("Row", "0") })
            Assert.Equal(letterSpacing, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Attribute("spc"));
    }

    private static void SetChartLetterSpacings(JsonObject program, int index, double? letterSpacing)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (letterSpacing is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["letterSpacing"] = letterSpacing.Value };
    }

    private static void AssertChartLetterSpacings(JsonObject program, int index, double? letterSpacing)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(letterSpacing, owner[field]?["letterSpacing"]?.GetValue<double>());
    }
}
