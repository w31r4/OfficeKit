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
    public void PpjChartTextShadowPreservesPresenceAndSourceLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        TrendlineListChart(program)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":"#112233","highlight":"#FFFF00","fontFamily":"Arial"}""");
        SetChartShadows(program, seriesIndex, JsonNode.Parse("""{"color":"#445566"}"""));
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        AssertChartShadows(ProjectTrendlineList(source), seriesIndex, """{"color":"#445566"}""");
        Assert.All(XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "outerShdw"), outer => Assert.Empty(outer.Attributes().Where(attr => !attr.IsNamespaceDeclaration)));
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (input, expected) in new (string? Input, string? Expected)[]
        {
            ("""{"color":{"token":"accent1"},"blur":0,"distance":0,"angle":0,"opacity":0,"rotateWithShape":false,"alignment":"ctr"}""",
             """{"color":{"token":"accent1"},"blur":0,"distance":0,"angle":0,"opacity":0,"rotateWithShape":false,"alignment":"ctr"}"""),
            ("""{"color":{"token":"mark","tint":1,"shade":0.5},"blur":2.5,"distance":4,"angle":-90,"opacity":{"token":"fade"},"rotateWithShape":true,"alignment":"br"}""",
             """{"color":"#808080","blur":2.5,"distance":4,"angle":270,"opacity":0.25,"rotateWithShape":true,"alignment":"br"}"""),
            ("""{"color":{"token":"accent1","alpha":1}}""", """{"color":{"token":"accent1"},"opacity":1}"""),
            ("""{"color":{"token":"accent1"}}""", """{"color":{"token":"accent1"}}"""),
            (null, null), ("""{"color":"#000000"}""", """{"color":"#000000"}"""),
        })
        {
            var projected = ProjectTrendlineList(source);
            projected["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#000000"},"fade":{"kind":"opacity","value":0.25}}""");
            SetChartShadows(projected, seriesIndex, input is null ? null : JsonNode.Parse(input));
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var shadows = native.Descendants(a + "outerShdw").ToArray();
            if (expected is null) Assert.Empty(shadows);
            else
            {
                Assert.NotEmpty(shadows);
                var value = JsonNode.Parse(expected)!;
                Assert.All(shadows, outer =>
                {
                    Assert.Equal(value["rotateWithShape"] is null ? null : value["rotateWithShape"]!.GetValue<bool>() ? "1" : "0", (string?)outer.Attribute("rotWithShape"));
                    Assert.Equal(value["blur"] is null ? null : (long?)(value["blur"]!.GetValue<double>() * 12700), (long?)outer.Attribute("blurRad"));
                    Assert.Equal(value["distance"] is null ? null : (long?)(value["distance"]!.GetValue<double>() * 12700), (long?)outer.Attribute("dist"));
                    Assert.Equal(value["angle"] is null ? null : (int?)(value["angle"]!.GetValue<double>() * 60000), (int?)outer.Attribute("dir"));
                    Assert.Equal(value["opacity"] is null ? null : (int?)(value["opacity"]!.GetValue<double>() * 100000), (int?)Assert.Single(outer.Elements()).Element(a + "alpha")?.Attribute("val"));
                });
            }
            var fresh = ProjectTrendlineList(output);
            AssertChartShadows(fresh, seriesIndex, expected);
            Assert.Equal("#112233", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["color"]!.GetValue<string>());
            Assert.Equal("#FFFF00", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["highlight"]!.GetValue<string>());
            var unchanged = CompileTrendlineList(fresh, output);
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Equal(output, unchanged.File.ToByteArray());
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextShadowValidatesNativeOrderTokensAndGeometry()
    {
        var valid = new SpreadsheetChartTextStyleArtifact { Shadow = new PresentationShadow { ColorRgb = "112233" } };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var shadowOnly));
        Assert.Equal(valid, shadowOnly);
        var explicitZero = valid.Clone(); explicitZero.Shadow.BlurRadiusEmu = 0; explicitZero.Shadow.RotateWithShape = false;
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(explicitZero));
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(null));
        valid.ColorRgb = "112233"; valid.FontFamily = "Arial"; valid.HighlightRgb = "FFFF00";
        var native = XlsxChartTextStyleCodec.StyleProperties("rPr", valid);
        Assert.Equal(new[] { "solidFill", "effectLst", "highlight", "latin" }, native.Elements().Select(element => element.Name.LocalName));
        var properties = new DocumentFormat.OpenXml.Drawing.RunProperties(native.ToString(SaveOptions.DisableFormatting));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021).Validate(properties));
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(native, out var restored));
        Assert.Equal(valid, restored);
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", Shadow = valid.Shadow.Clone() }));
        Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { Shadow = new PresentationShadow { ColorRgb = "123456", OpacityThousandthPercent = 100001 } }, "s", "c", "style"));

        foreach (var oversized in new[]
        {
            new PresentationShadow { ColorRgb = "123456", BlurRadiusEmu = 12_700_001 },
            new PresentationShadow { ColorRgb = "123456", DistanceEmu = 1_270_000_001 },
        })
            Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { Shadow = oversized }, "s", "c", "style"));

        var baseline = ChartTextStyleProgram("line", 0);
        SetChartShadows(baseline, 0, JsonNode.Parse("""{"color":"#112233"}"""));
        var compiled = CompileTrendlineList(baseline);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var source = RemoveEmbeddedPpj(compiled.File.ToByteArray());
        foreach (var input in new[] { "{}", """{"color":"#112233","blur":-1}""", """{"color":"#112233","distance":100001}""", """{"color":"#112233","angle":361}""", """{"color":"#112233","opacity":1.1}""", """{"color":{"token":"accent1","tint":0.5}}""", """{"color":{"token":"missing"}}""", """{"color":{"token":"wrongKind"}}""", """{"color":"#112233","opacity":{"token":"wrongKind"}}""" })
            foreach (var sourceBound in new[] { false, true })
            {
                var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
                program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"wrongKind":{"kind":"string","value":"#FFFFFF"}}""");
                SetChartShadows(program, 0, JsonNode.Parse(input));
                var rejected = CompileTrendlineList(program, sourceBound ? source : null);
                Assert.False(rejected.Ok, input); Assert.Empty(rejected.File);
            }
        foreach (var sourceBound in new[] { false, true })
        {
            var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
            program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent1":{"kind":"color","value":"#000000"},"fade":{"kind":"opacity","value":0.25}}""");
            SetChartShadows(program, 0, JsonNode.Parse("""{"color":{"token":"accent1","tint":1,"shade":0.5},"opacity":{"token":"fade"},"angle":360,"blur":0.0001}"""));
            var result = CompileTrendlineList(program, sourceBound ? source : null);
            Assert.True(result.Ok, Diagnostics(result));
            var projection = ProjectTrendlineList(RemoveEmbeddedPpj(result.File.ToByteArray()));
            var shadow = TrendlineListChart(projection)["style"]!["titleTextStyle"]!["shadow"]!;
            Assert.Equal("#808080", shadow["color"]!.GetValue<string>());
            Assert.Equal(0.25, shadow["opacity"]!.GetValue<double>());
            Assert.Equal(0d, shadow["angle"]!.GetValue<double>());
            Assert.Equal(1d / 12700, shadow["blur"]!.GetValue<double>(), 6);
        }
    }

    [Fact]
    public void PpjChartTextShadowHonorsPrecedenceAndVectorOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "shadow-style";
        chart["style"]!["legendTextStyle"] = JsonNode.Parse("""{"shadow":{"color":"#FFFFFF"},"bold":false}""");
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""{"id":"shadow-style","style":{"legendTextStyle":{"shadow":{"color":"#FF0000"}}}}"""));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""[{"target":"chart.legendTextStyle.shadow","sources":["styleRef","inline"]}]""");
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal("#FF0000", projected["style"]!["legendTextStyle"]!["shadow"]!["color"]!.GetValue<string>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#00FF00"}}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"shadow":{"color":"#FFFF00"}},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"shadow":{"color":{"token":"mark"},"rotateWithShape":false}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"shadow":{"color":"#000000","blur":0,"distance":0,"angle":0}}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, rgb) in new[] { ("Default", "FFFF00"), ("Override", "000000"), ("Row", "00FF00") })
            Assert.Equal(rgb, (string?)Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text)).Element(a + "rPr")?.Element(a + "effectLst")?.Element(a + "outerShdw")?.Element(a + "srgbClr")?.Attribute("val"));
    }

    private static void SetChartShadows(JsonObject program, int index, JsonNode? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
        {
            var style = owner[field]?.AsObject();
            if (value is null)
            {
                style?.Remove("shadow");
                if (style?.Count == 0) owner.Remove(field);
            }
            else
            {
                if (style is null) owner[field] = style = new JsonObject();
                style["shadow"] = value.DeepClone();
            }
        }
    }

    private static void AssertChartShadows(JsonObject program, int index, string? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.True(JsonNode.DeepEquals(value is null ? null : JsonNode.Parse(value), owner[field]?["shadow"]), owner[field]?.ToJsonString());
    }
}
