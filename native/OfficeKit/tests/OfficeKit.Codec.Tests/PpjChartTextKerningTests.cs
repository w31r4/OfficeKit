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
    public void PpjChartTextKerningPreservesThresholdZeroAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartKernings(program, seriesIndex, 30);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        Assert.Contains("Language chart", originalText);
        Assert.Contains("Fit ", originalText);
        AssertChartKernings(ProjectTrendlineList(source), seriesIndex, 30);
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (requested, expected) in new (double? Input, double? Expected)[]
        { (12.375, 12.38), (0, 0), (null, null), (1.23456, 1.23) })
        {
            var projected = ProjectTrendlineList(source);
            SetChartKernings(projected, seriesIndex, requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var attributes = native.Descendants().Attributes("kern").ToArray();
            if (expected is null) Assert.Empty(attributes);
            else
            {
                Assert.NotEmpty(attributes);
                Assert.All(attributes, attribute => Assert.Equal(checked((int)Math.Round(expected.Value * 100)), int.Parse(attribute.Value, CultureInfo.InvariantCulture)));
            }
            AssertChartKernings(ProjectTrendlineList(output), seriesIndex, expected);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextKerningValidatesNativeRangePrecisionAndPresence()
    {
        foreach (var value in new[] { -0.001, 768.001, double.NaN, double.PositiveInfinity })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.KerningHundredthPoints(value));
        foreach (var value in new uint[] { 76_801, uint.MaxValue })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { KerningHundredthPoints = value }, "s", "c", "style"));
        foreach (var value in new[] { "", "1.5", "76801", "-1" })
        {
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("kern", value);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
        }
        foreach (var (points, native) in new[] { (0d, 0u), (768d, 76800u), (0.005, 0u), (0.015, 2u) })
            Assert.Equal(native, XlsxChartTextStyleCodec.KerningHundredthPoints(points));
        foreach (var points in new[] { -0.001, 768.001 })
        {
            var program = ChartTextStyleProgram("line", 0);
            SetChartKernings(program, 0, points);
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { KerningHundredthPoints = 0 };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }), XlsxChartTextStyleCodec.Semantics(valid));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
    }

    [Fact]
    public void PpjChartTextKerningHonorsPrecedenceAndVectorText()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "kerning-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["kerning"] = 0, ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"kerning-style","style":{"legendTextStyle":{"kerning":30}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.kerning","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal(30, projected["style"]!["legendTextStyle"]!["kerning"]!.GetValue<double>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"kerning":30},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"kerning":0}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"kerning":0}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, kerning) in new[] { ("Default", "3000"), ("Override", "0"), ("Row", "0") })
            Assert.Equal(kerning, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Attribute("kern"));
    }

    private static void SetChartKernings(JsonObject program, int index, double? kerning)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (kerning is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["kerning"] = kerning.Value };
    }

    private static void AssertChartKernings(JsonObject program, int index, double? kerning)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(kerning, owner[field]?["kerning"]?.GetValue<double>());
    }
}
