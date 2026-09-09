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
    public void PpjChartTextReflectionPreservesIndependentEffectsLifecycle(string type, int seriesIndex)
    {
        var program = ChartTextStyleProgram(type, seriesIndex);
        TrendlineListChart(program)["style"]!["titleTextStyle"] = JsonNode.Parse("""{"color":"#112233","highlight":"#FFFF00","fontFamily":"Arial"}""");
        SetChartReflections(program, seriesIndex, new JsonObject());
        SetChartGlows(program, seriesIndex, JsonNode.Parse("""{"color":"#445566","radius":2}"""));
        SetChartInnerShadows(program, seriesIndex, JsonNode.Parse("""{"color":{"token":"accent1"}}"""));
        SetChartShadows(program, seriesIndex, JsonNode.Parse("""{"color":"#112233","rotateWithShape":false}"""));
        SetChartSoftEdges(program, seriesIndex, 1);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var originalText = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "t").Select(text => text.Value).ToArray();
        AssertChartReflections(ProjectTrendlineList(source), seriesIndex, "{}");
        Assert.All(XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path))).Descendants(a + "reflection"), native =>
            Assert.Equal(new[] { "stPos", "endPos" }, native.Attributes().Where(attr => !attr.IsNamespaceDeclaration).Select(attr => attr.Name.LocalName)));
        var noOp = CompileTrendlineList(ProjectTrendlineList(source), source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Equal(source, noOp.File.ToByteArray());
        foreach (var (input, expected) in new (string? Input, string? Expected)[]
        {
            ("""{"blur":0,"distance":0,"angle":0,"startOpacity":0,"endOpacity":0}""", """{"blur":0,"distance":0,"angle":0,"startOpacity":0,"endOpacity":0}"""),
            ("""{"blur":2.5,"distance":4,"angle":-90,"startOpacity":{"token":"fade"},"endOpacity":0.75}""", """{"blur":2.5,"distance":4,"angle":270,"startOpacity":0.25,"endOpacity":0.75}"""),
            ("""{"endOpacity":{"token":"fade"}}""", """{"endOpacity":0.25}"""),
            ("{}", "{}"), (null, null), ("{}", "{}"),
        })
        {
            var projected = ProjectTrendlineList(source);
            projected["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"fade":{"kind":"opacity","value":0.25}}""");
            SetChartReflections(projected, seriesIndex, input is null ? null : JsonNode.Parse(input));
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path)) Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            var native = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            Assert.Equal(originalText, native.Descendants(a + "t").Select(text => text.Value).ToArray());
            var reflections = native.Descendants(a + "reflection").ToArray();
            if (expected is null) Assert.Empty(reflections);
            else
            {
                Assert.NotEmpty(reflections);
                var value = JsonNode.Parse(expected)!;
                Assert.All(reflections, reflection =>
                {
                    Assert.Equal("0", (string?)reflection.Attribute("stPos"));
                    Assert.Equal("100000", (string?)reflection.Attribute("endPos"));
                    foreach (var (key, attr, scale) in new[] { ("blur", "blurRad", 12700), ("distance", "dist", 12700), ("angle", "dir", 60000), ("startOpacity", "stA", 100000), ("endOpacity", "endA", 100000) })
                        Assert.Equal(value[key] is null ? null : (long?)(value[key]!.GetValue<double>() * scale), (long?)reflection.Attribute(attr));
                });
            }
            Assert.All(native.Descendants(a + "effectLst"), effects =>
                Assert.Equal(expected is null ? new[] { "glow", "innerShdw", "outerShdw", "softEdge" } : new[] { "glow", "innerShdw", "outerShdw", "reflection", "softEdge" },
                    effects.Elements().Select(effect => effect.Name.LocalName)));
            var fresh = ProjectTrendlineList(output);
            AssertChartGlows(fresh, seriesIndex, """{"color":"#445566","radius":2}""");
            AssertChartInnerShadows(fresh, seriesIndex, """{"color":{"token":"accent1"}}""");
            AssertChartShadows(fresh, seriesIndex, """{"color":"#112233","rotateWithShape":false}""");
            AssertChartSoftEdges(fresh, seriesIndex, 1);
            AssertChartReflections(fresh, seriesIndex, expected);
            Assert.Equal("#112233", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["color"]!.GetValue<string>());
            Assert.Equal("#FFFF00", TrendlineListChart(fresh)["style"]!["titleTextStyle"]!["highlight"]!.GetValue<string>());
            var unchanged = CompileTrendlineList(fresh, output);
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Equal(output, unchanged.File.ToByteArray());
            source = output;
        }
    }

    [Fact]
    public void PpjChartTextReflectionValidatesOrderBoundsTokensAndPresence()
    {
        var valid = new SpreadsheetChartTextStyleArtifact { Reflection = new PresentationReflection() };
        XlsxChartTextStyleCodec.ValidateStyle(valid, "s", "c", "style");
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(XlsxChartTextStyleCodec.StyleProperties("rPr", valid), out var reflectionOnly));
        Assert.Equal(valid, reflectionOnly);
        var zero = valid.Clone(); zero.Reflection.BlurRadiusEmu = 0; zero.Reflection.EndOpacityThousandthPercent = 0;
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(zero));
        Assert.NotEqual(XlsxChartTextStyleCodec.Semantics(valid), XlsxChartTextStyleCodec.Semantics(null));
        Assert.False(XlsxChartTextStyleCodec.IsGlobalFontFamilyProfile(new SpreadsheetChartTextStyleArtifact { FontFamily = "Arial", Reflection = new PresentationReflection() }));
        valid.Glow = new PresentationGlow { ColorRgb = "FFFFFF", RadiusEmu = 0 };
        valid.InnerShadow = new PresentationInnerShadow { ColorScheme = "accent1" };
        valid.Shadow = new PresentationShadow { ColorRgb = "000000" };
        valid.SoftEdge = new PresentationSoftEdge { RadiusEmu = 0 };
        valid.ColorRgb = "112233"; valid.FontFamily = "Arial"; valid.HighlightRgb = "FFFF00";
        var native = XlsxChartTextStyleCodec.StyleProperties("rPr", valid);
        Assert.Equal(new[] { "solidFill", "effectLst", "highlight", "latin" }, native.Elements().Select(element => element.Name.LocalName));
        var properties = new DocumentFormat.OpenXml.Drawing.RunProperties(native.ToString(SaveOptions.DisableFormatting));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021).Validate(properties));
        Assert.True(XlsxChartTextStyleCodec.TryExactStyleProperties(native, out var restored));
        Assert.Equal(valid, restored);
        // The shared native writer also retains optional values for vector text runs.
        var direct = new DocumentFormat.OpenXml.Drawing.RunProperties();
        PptxReflectionCodec.Apply(direct, zero.Reflection);
        Assert.True(PptxReflectionCodec.TryRead(direct, out var read));
        Assert.Equal(zero.Reflection, read);
        foreach (var invalid in new[]
        {
            new PresentationReflection { BlurRadiusEmu = -1 }, new PresentationReflection { BlurRadiusEmu = 12_700_001 },
            new PresentationReflection { DistanceEmu = 1_270_000_001 }, new PresentationReflection { DirectionAngle60000 = 21_600_000 },
            new PresentationReflection { StartOpacityThousandthPercent = 100001 }, new PresentationReflection { EndOpacityThousandthPercent = 100001 },
        }) Assert.Throws<CodecException>(() => XlsxChartTextStyleCodec.ValidateStyle(new SpreadsheetChartTextStyleArtifact { Reflection = invalid }, "s", "c", "style"));

        var baseline = ChartTextStyleProgram("line", 0);
        SetChartReflections(baseline, 0, new JsonObject());
        var compiled = CompileTrendlineList(baseline);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var source = RemoveEmbeddedPpj(compiled.File.ToByteArray());
        foreach (var input in new[] { "null", """{"blur":-1}""", """{"distance":100001}""", """{"angle":361}""", """{"startOpacity":1.1}""", """{"endOpacity":-0.1}""", """{"blur":{"token":"fade"}}""", """{"startOpacity":{"token":"wrongKind"}}""", """{"endOpacity":{"token":"missing"}}""", """{"extra":0}""" })
            foreach (var sourceBound in new[] { false, true })
            {
                var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
                program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"wrongKind":{"kind":"string","value":"0.5"}}""");
                // Assign null explicitly to prove the schema distinguishes it from field omission.
                foreach (var (owner, field) in ChartTextStyleOwners(program, 0))
                {
                    var style = owner[field]?.AsObject() ?? new JsonObject();
                    if (owner[field] is null) owner[field] = style;
                    style["reflection"] = JsonNode.Parse(input);
                }
                var rejected = CompileTrendlineList(program, sourceBound ? source : null);
                Assert.False(rejected.Ok, input); Assert.Empty(rejected.File);
            }
        foreach (var sourceBound in new[] { false, true })
        {
            var program = sourceBound ? ProjectTrendlineList(source) : ChartTextStyleProgram("line", 0);
            program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"fade":{"kind":"opacity","value":0.25}}""");
            SetChartReflections(program, 0, JsonNode.Parse("""{"startOpacity":{"token":"fade"},"endOpacity":{"token":"fade"},"angle":360,"blur":0.0001}"""));
            var result = CompileTrendlineList(program, sourceBound ? source : null);
            Assert.True(result.Ok, Diagnostics(result));
            var reflection = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(result.File.ToByteArray())))["style"]!["titleTextStyle"]!["reflection"]!;
            Assert.Equal(0.25, reflection["startOpacity"]!.GetValue<double>());
            Assert.Equal(0.25, reflection["endOpacity"]!.GetValue<double>());
            Assert.Equal(0d, reflection["angle"]!.GetValue<double>());
            Assert.Equal(1d / 12700, reflection["blur"]!.GetValue<double>(), 6);
            Assert.Null(reflection["distance"]);
        }
    }

    [Fact]
    public void PpjChartTextReflectionHonorsPrecedenceAndVectorOverrides()
    {
        var program = TrendlineListProgram("line", 0);
        var chart = TrendlineListChart(program);
        chart["styleRef"] = "reflection-style";
        chart["style"]!["legendTextStyle"] = JsonNode.Parse("""{"reflection":{"blur":1},"bold":false}""");
        program["design"]!["styles"]!["chart"]!.AsArray().Add(JsonNode.Parse("""{"id":"reflection-style","style":{"legendTextStyle":{"reflection":{"blur":3}}}}"""));
        program["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""[{"target":"chart.legendTextStyle.reflection","sources":["styleRef","inline"]}]""");
        var compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        var projected = TrendlineListChart(ProjectTrendlineList(RemoveEmbeddedPpj(compiled.File.ToByteArray())));
        Assert.Equal(3d, projected["style"]!["legendTextStyle"]!["reflection"]!["blur"]!.GetValue<double>());
        Assert.False(projected["style"]!["legendTextStyle"]!["bold"]!.GetValue<bool>());
        chart.Remove("styleRef"); chart.Remove("xAxis"); chart.Remove("yAxis");
        chart["chartType"] = "heatmap";
        chart["data"] = JsonNode.Parse("""{"categories":["A","B"],"series":[{"id":"row","name":"Row","values":[1,2]}]}""");
        program["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"fade":{"kind":"opacity","value":0.25}}""");
        chart["style"] = JsonNode.Parse("""{"titleTextStyle":{"reflection":{},"glow":{"color":"#445566","radius":1}},"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false,"axisTextStyle":{"reflection":{"endOpacity":{"token":"fade"}}}}}""");
        chart["title"] = JsonNode.Parse("""{"paragraphs":[{"runs":[{"text":"Default"},{"text":"Override","style":{"reflection":{"blur":0,"distance":0,"angle":0,"startOpacity":0,"endOpacity":0}}}]}]}""");
        compiled = CompileTrendlineList(program);
        Assert.True(compiled.Ok, Diagnostics(compiled));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var runs = ZipPartPaths(compiled.File.ToByteArray()).Where(path => path.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(path => XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(compiled.File.ToByteArray(), path))).Descendants(a + "r")).ToArray();
        foreach (var (text, alpha) in new (string Text, long? Alpha)[] { ("Default", null), ("Override", 0), ("Row", 25000) })
        {
            var run = Assert.Single(runs.Where(run => run.Element(a + "t")?.Value == text));
            var reflection = run.Element(a + "rPr")?.Element(a + "effectLst")?.Element(a + "reflection");
            Assert.NotNull(reflection);
            Assert.Equal(alpha, (long?)reflection.Attribute("endA"));
            Assert.Equal(text == "Override" ? 0L : (long?)null, (long?)reflection.Attribute("blurRad"));
            if (text != "Row") Assert.NotNull(run.Element(a + "rPr")?.Element(a + "effectLst")?.Element(a + "glow"));
        }
    }

    private static void SetChartReflections(JsonObject program, int index, JsonNode? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
        {
            var style = owner[field]?.AsObject();
            if (value is null)
            {
                style?.Remove("reflection");
                if (style?.Count == 0) owner.Remove(field);
            }
            else
            {
                if (style is null) owner[field] = style = new JsonObject();
                style["reflection"] = value.DeepClone();
            }
        }
    }

    private static void AssertChartReflections(JsonObject program, int index, string? value)
    {
        foreach (var (owner, field) in ChartTextStyleOwners(program, index))
            Assert.True(JsonNode.DeepEquals(value is null ? null : JsonNode.Parse(value), owner[field]?["reflection"]), owner[field]?.ToJsonString());
    }
}
