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
    public void PpjChartTextGlowPreservesIndependentEffectsLifecycle(string type, int seriesIndex)
    {
        const string shadow = """{"color":"#112233","rotateWithShape":false}""";
        var program = ChartTextStyleProgram(type, seriesIndex);
        TrendlineListChart(program)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":"#112233","highlight":"#FFFF00","fontFamily":"Arial"}""");
        SetChartShadows(program, seriesIndex, JsonNode.Parse(shadow));
        SetChartGlows(program, seriesIndex, JsonNode.Parse("""{"color":"#445566","radius":2.5}"""));
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        AssertChartGlows(ProjectTrendlineList(source), seriesIndex, """{"color":"#445566","radius":2.5}""");
        AssertChartShadows(ProjectTrendlineList(source), seriesIndex, shadow);
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (input, expected, keepShadow) in new (string? Input, string? Expected, bool KeepShadow)[]
        {
            ("""{"color":{"token":"accent1"},"radius":0,"opacity":0}""", """{"color":{"token":"accent1"},"radius":0,"opacity":0}""", true),
            ("""{"color":{"token":"mark","tint":1,"shade":0.5},"radius":4,"opacity":{"token":"fade"}}""", """{"color":"#808080","radius":4,"opacity":0.25}""", true),
            ("""{"color":{"token":"accent1","alpha":1},"radius":2}""", """{"color":{"token":"accent1"},"radius":2,"opacity":1}""", true),
            ("""{"color":{"token":"accent1"},"radius":2}""", """{"color":{"token":"accent1"},"radius":2}""", true),
            (null, null, true),
            ("""{"color":"#000000","radius":0}""", """{"color":"#000000","radius":0}""", false),
            ("""{"color":"#000000","radius":0}""", """{"color":"#000000","radius":0}""", true),
            (null, null, false),
            ("""{"color":"#FF0000","radius":3}""", """{"color":"#FF0000","radius":3}""", false),
        })
        {
            var projected = ProjectTrendlineList(source);
            projected["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#000000"},"fade":{"kind":"opacity","value":0.25}}""");
            SetChartGlows(projected, seriesIndex, input is null ? null : JsonNode.Parse(input));
            SetChartShadows(projected, seriesIndex, keepShadow ? JsonNode.Parse(shadow) : null);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var glows = native.Descendants(a + "glow").ToArray();
            if (expected is null) Assert.Empty(glows);
            else
            {
                Assert.NotEmpty(glows);
                var value = JsonNode.Parse(expected)!;
                Assert.All(glows, glow =>
                {
                    Assert.Equal((long)(value["radius"]!.GetValue<double>() * 12700), (long?)glow.Attribute("rad"));
                    Assert.Equal(value["opacity"] is null ? null : (int?)(value["opacity"]!.GetValue<double>() * 100000), (int?)Assert.Single(glow.Elements()).Element(a + "alpha")?.Attribute("val"));
                });
            }
            var effects = native.Descendants(a + "effectLst").ToArray();
            if (expected is null && !keepShadow) Assert.Empty(effects);
            else
            {
                var order = new List<string>();
                if (expected is not null) order.Add("glow");
                if (keepShadow) order.Add("outerShdw");
                Assert.NotEmpty(effects);
                Assert.All(effects, effect => Assert.Equal(order, effect.Elements().Select(element => element.Name.LocalName)));
            }
            var fresh = ProjectTrendlineList(output);
            AssertChartGlows(fresh, seriesIndex, expected);
            AssertChartShadows(fresh, seriesIndex, keepShadow ? shadow : null);
            Assert.Equal("#FFFF00", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["highlight"]!.GetValue<string>());
            var unchanged = CompileTrendlineList(fresh, output);
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Equal(output, unchanged.File.ToByteArray());
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextGlowValidatesOrderBoundsAndTokens()
    {
        var valid = new SpreadsheetChartTextStyleArtifact { Glow = new PresentationGlow { ColorRgb = "112233", RadiusEmu = 0 } };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var glowOnly));
        Assert.Equal(valid, glowOnly);
        var zeroOpacity = valid.Clone(); zeroOpacity.Glow.OpacityThousandthPercent = 0;
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(zeroOpacity));
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(null));
        valid.ColorRgb = "112233"; valid.FontFamily = "Arial"; valid.HighlightRgb = "FFFF00";
        valid.Shadow = new PresentationShadow { ColorScheme = "accent1", RotateWithShape = false };
        var native = XlsxChartTextStyleCodec.StyleProperties("rPr", valid);
        Assert.Equal(new[] { "solidFill", "effectLst", "highlight", "latin" }, native.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(new[] { "glow", "outerShdw" }, native.Elements().Single(element => element.Name.LocalName == "effectLst").Elements().Select(element => element.Name.LocalName));
        var properties = new DocumentFormat.OpenXml.Drawing.RunProperties(native.ToString(SaveOptions.DisableFormatting));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021).Validate(properties));
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(native, out var restored));
        Assert.Equal(valid, restored);
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", Glow = valid.Glow.Clone() }));
        foreach (var invalid in new[]
        {
            new PresentationGlow { ColorRgb = "123456" },
            new PresentationGlow { ColorRgb = "123456", RadiusEmu = -1 },
            new PresentationGlow { ColorRgb = "123456", RadiusEmu = 12_700_001 },
            new PresentationGlow { ColorRgb = "123456", RadiusEmu = 0, OpacityThousandthPercent = 100001 },
        })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { Glow = invalid }, "s", "c", "style"));

        var baseline = ChartTextStyleProgram("line", 0);
        SetChartGlows(baseline, 0, JsonNode.Parse("""{"color":"#112233","radius":2}"""));
        var compiled = CompileTrendlineList(baseline);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var source = RemoveEmbeddedPpj(compiled.File.ToByteArray());
        foreach (var input in new[] { "{}", """{"color":"#112233"}""", """{"color":"#112233","radius":-1}""", """{"color":"#112233","radius":1001}""", """{"color":"#112233","radius":1,"opacity":1.1}""", """{"color":{"token":"accent1","tint":0.5},"radius":1}""", """{"color":{"token":"missing"},"radius":1}""", """{"color":{"token":"wrongKind"},"radius":1}""", """{"color":"#112233","radius":1,"opacity":{"token":"wrongKind"}}""" })
            foreach (var sourceBound in new[] { false, true })
            {
                var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
                program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"wrongKind":{"kind":"string","value":"#FFFFFF"}}""");
                SetChartGlows(program, 0, JsonNode.Parse(input));
                var rejected = CompileTrendlineList(program, sourceBound ? source : null);
                Assert.False(rejected.Ok, input); Assert.Empty(rejected.File);
            }
        foreach (var sourceBound in new[] { false, true })
        {
            var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
            program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent1":{"kind":"color","value":"#000000"},"fade":{"kind":"opacity","value":0.25}}""");
            SetChartGlows(program, 0, JsonNode.Parse("""{"color":{"token":"accent1","tint":1,"shade":0.5},"opacity":{"token":"fade"},"radius":0.0001}"""));
            var result = CompileTrendlineList(program, sourceBound ? source : null);
            Assert.True(result.Ok, Diagnostics(result));
            var projection = ProjectTrendlineList(RemoveEmbeddedPpj(result.File.ToByteArray()));
            var glow = TrendlineListChart(projection)["style"]!["titleTextStyle"]!["glow"]!;
            Assert.Equal("#808080", glow["color"]!.GetValue<string>());
            Assert.Equal(0.25, glow["opacity"]!.GetValue<double>());
            Assert.Equal(1d / 12700, glow["radius"]!.GetValue<double>(), 6);
        }
    }

    [Fact]
    public void PpjChartTextGlowHonorsPrecedenceAndVectorOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "glow-style";
        chart["style"]!["legendTextStyle"] = JsonNode.Parse("""{"glow":{"color":"#FFFFFF","radius":2},"bold":false}""");
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""{"id":"glow-style","style":{"legendTextStyle":{"glow":{"color":"#FF0000","radius":3}}}}"""));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""[{"target":"chart.legendTextStyle.glow","sources":["styleRef","inline"]}]""");
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("#FF0000", projected["style"]!["legendTextStyle"]!["glow"]!["color"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#00FF00"}}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"glow":{"color":"#FFFF00","radius":2},"shadow":{"color":"#112233"}},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"glow":{"color":{"token":"mark"},"radius":0}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"glow":{"color":"#000000","radius":3}}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, rgb, radius) in new[] { ("Default", "FFFF00", 25400L), ("Override", "000000", 38100L), ("Row", "00FF00", 0L) })
        {
            var effect = Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Element(a + "effectLst");
            Assert.Equal(rgb, (string?)effect?.Element(a + "glow")?.Element(a + "srgbClr")?.Attribute("val"));
            Assert.Equal(radius, (long?)effect?.Element(a + "glow")?.Attribute("rad"));
            if (text != "Row") Assert.Equal("112233", (string?)effect?.Element(a + "outerShdw")?.Element(a + "srgbClr")?.Attribute("val"));
        }
    }

    private static void SetChartGlows(JsonObject program, int index, JsonNode? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
        {
            var style = owner[field]?.AsObject();
            if (value is null)
            {
                style?.Remove("glow");
                if (style?.Count == 0) owner.Remove(field);
            }
            else
            {
                if (style is null) owner[field] = style = new JsonObject();
                style["glow"] = value.DeepClone();
            }
        }
    }

    private static void AssertChartGlows(JsonObject program, int index, string? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.True(JsonNode.DeepEquals(value is null ? null : JsonNode.Parse(value), owner[field]?["glow"]), owner[field]?.ToJsonString());
    }
}
