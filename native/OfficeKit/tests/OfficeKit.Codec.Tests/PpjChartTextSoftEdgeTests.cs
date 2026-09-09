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
    public void PpjChartTextSoftEdgePreservesIndependentEffectsLifecycle(string type, int seriesIndex)
    {
        const string shadow = """{"color":{"token":"accent2"},"rotateWithShape":false}""";
        const string glow = """{"color":{"token":"accent1"},"radius":2,"opacity":0}""";
        var program = ChartTextStyleProgram(type, seriesIndex);
        TrendlineListChart(program)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":"#112233","highlight":"#FFFF00","fontFamily":"Arial"}""");
        SetChartShadows(program, seriesIndex, JsonNode.Parse(shadow));
        SetChartGlows(program, seriesIndex, JsonNode.Parse(glow));
        SetChartSoftEdges(program, seriesIndex, 2.5);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        var initialProjection = ProjectTrendlineList(source);
        AssertChartSoftEdges(initialProjection, seriesIndex, 2.5);
        AssertChartShadows(initialProjection, seriesIndex, shadow);
        AssertChartGlows(initialProjection, seriesIndex, glow);
        var noOp = CompileTrendlineList(initialProjection, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (radius, keepGlow, keepShadow) in new (double? Radius, bool KeepGlow, bool KeepShadow)[]
        {
            (0, true, true), (4, true, true), (null, true, true),
            (1, false, true), (1, true, false), (1, true, true),
            (1, false, false), (null, false, false), (0, false, false),
        })
        {
            var projected = ProjectTrendlineList(source);
            SetChartSoftEdges(projected, seriesIndex, radius);
            SetChartGlows(projected, seriesIndex, keepGlow ? JsonNode.Parse(glow) : null);
            SetChartShadows(projected, seriesIndex, keepShadow ? JsonNode.Parse(shadow) : null);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var edges = native.Descendants(a + "softEdge").ToArray();
            if (radius is null) Assert.Empty(edges);
            else
            {
                Assert.NotEmpty(edges);
                Assert.All(edges, edge => Assert.Equal((long)(radius.Value * 12700), (long?)edge.Attribute("rad")));
            }
            var effects = native.Descendants(a + "effectLst").ToArray();
            if (radius is null && !keepGlow && !keepShadow) Assert.Empty(effects);
            else
            {
                var order = new List<string>();
                if (keepGlow) order.Add("glow");
                if (keepShadow) order.Add("outerShdw");
                if (radius is not null) order.Add("softEdge");
                Assert.NotEmpty(effects);
                Assert.All(effects, effect => Assert.Equal(order, effect.Elements().Select(element => element.Name.LocalName)));
            }
            var fresh = ProjectTrendlineList(output);
            AssertChartSoftEdges(fresh, seriesIndex, radius);
            AssertChartGlows(fresh, seriesIndex, keepGlow ? glow : null);
            AssertChartShadows(fresh, seriesIndex, keepShadow ? shadow : null);
            Assert.Equal("#FFFF00", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["highlight"]!.GetValue<string>());
            var unchanged = CompileTrendlineList(fresh, output);
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Equal(output, unchanged.File.ToByteArray());
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextSoftEdgeValidatesOrderBoundsAndPrecision()
    {
        var valid = new SpreadsheetChartTextStyleArtifact { SoftEdge = new PresentationSoftEdge { RadiusEmu = 0 } };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var edgeOnly));
        Assert.Equal(valid, edgeOnly);
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(null));
        valid.ColorRgb = "112233"; valid.FontFamily = "Arial"; valid.HighlightRgb = "FFFF00";
        valid.Shadow = new PresentationShadow { ColorScheme = "accent1", RotateWithShape = false };
        valid.Glow = new PresentationGlow { ColorRgb = "445566", RadiusEmu = 12700, OpacityThousandthPercent = 0 };
        var native = XlsxChartTextStyleCodec.StyleProperties("rPr", valid);
        Assert.Equal(new[] { "solidFill", "effectLst", "highlight", "latin" }, native.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(new[] { "glow", "outerShdw", "softEdge" }, native.Elements().Single(element => element.Name.LocalName == "effectLst").Elements().Select(element => element.Name.LocalName));
        var properties = new DocumentFormat.OpenXml.Drawing.RunProperties(native.ToString(SaveOptions.DisableFormatting));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021).Validate(properties));
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(native, out var restored));
        Assert.Equal(valid, restored);
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", SoftEdge = valid.SoftEdge.Clone() }));
        foreach (var invalid in new[] { new PresentationSoftEdge(), new PresentationSoftEdge { RadiusEmu = -1 }, new PresentationSoftEdge { RadiusEmu = 12_700_001 } })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { SoftEdge = invalid }, "s", "c", "style"));

        // Standard themes are native effect colors, but cannot bypass an
        // explicitly declared wrong-kind token or the RGB foreground profile.
        foreach (var field in new[] { "shadow", "glow" })
        {
            var program = ChartTextStyleProgram("line", 0);
            program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent1":{"kind":"string","value":"#FFFFFF"}}""");
            var effect = new JsonObject { ["color"] = JsonNode.Parse("""{"token":"accent1"}""") };
            if (field == "glow") effect["radius"] = 1;
            TrendlineListChart(program)["style"]!["titleTextStyle"] = new JsonObject { [field] = effect };
            var rejected = CompileTrendlineList(program);
            Assert.False(rejected.Ok); Assert.Empty(rejected.File);
            Assert.Contains(rejected.Diagnostics, diagnostic => diagnostic.Code == "ppj.grammar.tokenKind");
        }
        var foreground = ChartTextStyleProgram("line", 0);
        TrendlineListChart(foreground)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":{"token":"accent1"},"softEdge":{"radius":1}}""");
        Assert.False(CompileTrendlineList(foreground).Ok);

        var baseline = ChartTextStyleProgram("line", 0);
        SetChartSoftEdges(baseline, 0, 2);
        var compiled = CompileTrendlineList(baseline);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var source = RemoveEmbeddedPpj(compiled.File.ToByteArray());
        foreach (var input in new[] { "{}", """{"radius":-1}""", """{"radius":1001}""", """{"radius":"1"}""", """{"radius":{"token":"size"}}""", """{"radius":1,"extra":0}""" })
            foreach (var sourceBound in new[] { false, true })
            {
                var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
                TrendlineListChart(program)["style"]!["titleTextStyle"] = new JsonObject { ["softEdge"] = JsonNode.Parse(input) };
                var rejected = CompileTrendlineList(program, sourceBound ? source : null);
                Assert.False(rejected.Ok, input); Assert.Empty(rejected.File);
            }
        foreach (var sourceBound in new[] { false, true })
        {
            var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
            SetChartSoftEdges(program, 0, 0.0001);
            var result = CompileTrendlineList(program, sourceBound ? source : null);
            Assert.True(result.Ok, Diagnostics(result));
            var projection = ProjectTrendlineList(RemoveEmbeddedPpj(result.File.ToByteArray()));
            Assert.Equal(1d / 12700, TrendlineListChart(projection)["style"]!["titleTextStyle"]!["softEdge"]!["radius"]!.GetValue<double>(), 6);
        }
    }

    [Fact]
    public void PpjChartTextSoftEdgeHonorsPrecedenceAndVectorOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "edge-style";
        chart["style"]!["legendTextStyle"] = JsonNode.Parse("""{"softEdge":{"radius":1},"bold":false}""");
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""{"id":"edge-style","style":{"legendTextStyle":{"softEdge":{"radius":3}}}}"""));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""[{"target":"chart.legendTextStyle.softEdge","sources":["styleRef","inline"]}]""");
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal(3d, projected["style"]!["legendTextStyle"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"softEdge":{"radius":2},"glow":{"color":"#FFFF00","radius":2},"shadow":{"color":"#112233"}},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"softEdge":{"radius":3}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"softEdge":{"radius":0}}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, radius) in new[] { ("Default", 25400L), ("Override", 0L), ("Row", 38100L) })
        {
            var effect = Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Element(a + "effectLst");
            Assert.Equal(radius, (long?)effect?.Element(a + "softEdge")?.Attribute("rad"));
            if (text != "Row")
            {
                Assert.Equal("112233", (string?)effect?.Element(a + "outerShdw")?.Element(a + "srgbClr")?.Attribute("val"));
                Assert.Equal("FFFF00", (string?)effect?.Element(a + "glow")?.Element(a + "srgbClr")?.Attribute("val"));
            }
        }
    }

    private static void SetChartSoftEdges(JsonObject program, int index, double? radius)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
        {
            var style = owner[field]?.AsObject();
            if (radius is null)
            {
                style?.Remove("softEdge");
                if (style?.Count == 0) owner.Remove(field);
            }
            else
            {
                if (style is null) owner[field] = style = new JsonObject();
                style["softEdge"] = new JsonObject { ["radius"] = radius.Value };
            }
        }
    }

    private static void AssertChartSoftEdges(JsonObject program, int index, double? radius)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.Equal(radius, owner[field]?["softEdge"]?["radius"]?.GetValue<double>());
    }
}
