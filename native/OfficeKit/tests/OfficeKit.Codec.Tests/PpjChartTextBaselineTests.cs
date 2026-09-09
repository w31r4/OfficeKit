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
    public void PpjChartTextBaselinePreservesSignedZeroAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        SetChartBaselines(program, seriesIndex, 30);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        AssertChartBaselines(ProjectTrendlineList(source), seriesIndex, 30);
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (requested, expected) in new (double? Input, double? Expected)[]
        { (-25.125, -25.125), (0, 0), (null, null), (1.23456, 1.235) })
        {
            var projected = ProjectTrendlineList(source);
            SetChartBaselines(projected, seriesIndex, requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var attributes = native.Descendants().Attributes("baseline").ToArray();
            if (expected is null) Assert.Empty(attributes);
            else
            {
                Assert.NotEmpty(attributes);
                Assert.All(attributes, attribute => Assert.Equal(checked((int)Math.Round(expected.Value * 1000)), int.Parse(attribute.Value, CultureInfo.InvariantCulture)));
            }
            AssertChartBaselines(ProjectTrendlineList(output), seriesIndex, expected);
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextBaselineValidatesNativeRangePrecisionAndPresence()
    {
        foreach (var value in new[] { -400.001, 400.001, double.NaN, double.PositiveInfinity })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.BaselineThousandthPercent(value));
        foreach (var value in new[] { -400_001, 400_001 })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { BaselineThousandthPercent = value }, "s", "c", "style"));
        foreach (var value in new[] { "", "1.5", "400001", "-400001" })
        {
            var properties = XElement.Parse("<rPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
            properties.SetAttributeValue("baseline", value);
            Assert.False(XlsxChartTextStyleCodec.TryExactStyleProperties(properties, out _));
        }
        foreach (var (percentage, native) in new[] { (-400d, -400000), (400d, 400000), (0.0005, 0), (0.0015, 2), (-0.0015, -2) })
            Assert.Equal(native, XlsxChartTextStyleCodec.BaselineThousandthPercent(percentage));
        foreach (var percentage in new[] { -400.001, 400.001 })
        {
            var program = ChartTextStyleProgram("line", 0);
            SetChartBaselines(program, 0, percentage);
            Assert.False(CompileTrendlineList(program).Ok);
        }
        var valid = new SpreadsheetChartTextStyleArtifact { BaselineThousandthPercent = 0 };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var restored));
        Assert.Equal(valid, restored);
        valid.FontFamily = "Arial";
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial" }), XlsxChartTextStyleCodec.Semantics(valid));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(valid));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.Validate(new SpreadsheetChartArtifact { Id = "c", TextStyle = valid }, "s"));
    }

    [Fact]
    public void PpjChartTextBaselineHonorsPrecedenceAndVectorText()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "baseline-style";
        chart["style"]!["legendTextStyle"] = new JsonObject { ["baseline"] = 0, ["bold"] = false };
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""
            {"id":"baseline-style","style":{"legendTextStyle":{"baseline":30}}}
            """));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"chart.legendTextStyle.baseline","sources":["styleRef","inline"]}]
            """);
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal(30, projected["style"]!["legendTextStyle"]!["baseline"]!.GetValue<double>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"baseline":30},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"baseline":0}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"baseline":-25}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, baseline) in new[] { ("Default", "30000"), ("Override", "-25000"), ("Row", "0") })
            Assert.Equal(baseline, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Attribute("baseline"));
    }

    private static void SetChartBaselines(JsonObject program, int index, double? baseline)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            if (baseline is null) owner.Remove(field);
            else owner[field] = new JsonObject { ["baseline"] = baseline.Value };
    }

    private static void AssertChartBaselines(JsonObject program, int index, double? baseline)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(baseline, owner[field]?["baseline"]?.GetValue<double>());
    }
}
