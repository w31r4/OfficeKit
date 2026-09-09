using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    private const string CustomErrorData = """{"valueType":"custom","direction":"y","type":"both","plus":{"values":[0,0.5,1,2],"formatCode":"0.0"},"minus":{"values":[0.25,1,0,3]}}""";

    [Theory]
    [InlineData("column", 0)]
    [InlineData("bar", 0)]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjCustomErrorDataAuthorsAndEditsLiteralOwners(string chartType, int seriesIndex)
    {
        var program = TrendlineListProgram(chartType, seriesIndex);
        TrendlineListSeries(program, seriesIndex)["errorBars"] = JsonNode.Parse(CustomErrorData);
        SetErrorDataFormatToken(program, seriesIndex, "0.0");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectTrendlineList(source);
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse(CustomErrorData), TrendlineListSeries(projected, seriesIndex)["errorBars"]));
        var noOp = CompileTrendlineList(projected, source);
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);
        Assert.Equal(ZipBytes(source, path), ZipBytes(noOp.File.ToByteArray(), path));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var requested in new string?[]
        {
            """{"valueType":"custom","direction":"y","type":"both","plus":{"values":[1,2,3,4],"formatCode":"0.00%"},"minus":{"values":[4,3,2,1],"formatCode":"0.000"}}""",
            """{"valueType":"custom","direction":"x","type":"plus","plus":{"values":[0,1,0,2]},"noEndCap":true,"stroke":{"color":"#2563EB","width":1.5,"dash":"solid"}}""",
            """{"valueType":"custom","direction":"y","type":"minus","minus":{"values":[2,0,1,0]}}""",
            """{"valueType":"fixed-value","direction":"y","type":"both","value":2}""",
            CustomErrorData,
            null,
            """{"valueType":"custom","direction":"y","type":"both","plus":{"values":[0,1,2,3]},"minus":{"values":[3,2,1,0]}}""",
        })
        {
            var series = TrendlineListSeries(projected, seriesIndex);
            if (requested is null) series.Remove("errorBars");
            else
            {
                series["errorBars"] = JsonNode.Parse(requested);
                if (series["errorBars"]!["plus"]?["formatCode"] is { } format)
                    SetErrorDataFormatToken(projected, seriesIndex, format.GetValue<string>());
            }
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var beforeXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var afterXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var beforeSeries = beforeXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            var afterSeries = afterXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            var expected = requested is null ? null : JsonNode.Parse(requested);
            if (expected is null) Assert.Empty(afterSeries.Elements(c + "errBars"));
            else
            {
                var owner = Assert.Single(afterSeries.Elements(c + "errBars"));
                Assert.Contains(owner.ElementsBeforeSelf(), element => element.Name == c + "trendline");
                Assert.Contains(owner.ElementsAfterSelf(), element => element.Name == c + "cat");
                foreach (var side in new[] { "plus", "minus" })
                {
                    if (expected[side] is not { } data) Assert.Null(owner.Element(c + side));
                    else
                    {
                        var cache = Assert.Single(owner.Element(c + side)!.Elements());
                        Assert.Equal(c + "numLit", cache.Name);
                        Assert.Equal(data["values"]!.AsArray().Select(value => value!.GetValue<double>()), cache.Elements(c + "pt").Select(point => double.Parse(point.Element(c + "v")!.Value, System.Globalization.CultureInfo.InvariantCulture)));
                        Assert.Equal(data["formatCode"]?.GetValue<string>(), cache.Element(c + "formatCode")?.Value);
                    }
                }
            }
            OpenXmlElement nativeElement = chartType is "column" or "bar"
                ? new C.BarChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting))
                : new C.LineChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(nativeElement));
            beforeSeries.Elements(c + "errBars").Remove();
            afterSeries.Elements(c + "errBars").Remove();
            Assert.True(XNode.DeepEquals(beforeXml, afterXml));
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            Assert.True(JsonNode.DeepEquals(expected, TrendlineListSeries(projected, seriesIndex)["errorBars"]));
        }
    }

    [Fact]
    public void PpjCustomErrorDataRejectsInvalidDeclarations()
    {
        var valid = TrendlineListProgram("line", 0);
        TrendlineListSeries(valid, 0)["errorBars"] = JsonNode.Parse(CustomErrorData);
        var authored = CompileTrendlineList(valid);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        Action<JsonObject>[] mutations =
        [
            error => error.Remove("minus"),
            error => error["type"] = "plus",
            error => error["value"] = 1,
            error => { error["valueType"] = "fixed-value"; error["value"] = 1; },
            error => error["plus"]!["values"] = new JsonArray(1, 2),
            error => error["plus"]!["values"]![0] = -1,
            error => error["plus"]!["values"]![0] = null,
            error => error["plus"]!["values"]![0] = 9_007_199_254_740_992D,
            error => error["plus"]!["formatCode"] = "",
            error => error["plus"]!["formatCode"] = "0\n0",
            error => error["plus"]!["formatCode"] = new string('0', 256),
            error => error["plus"]!["formula"] = "'Sheet1'!$D$2:$D$5",
        ];
        foreach (var mutation in mutations)
        {
            var program = valid.DeepClone().AsObject();
            mutation(TrendlineListSeries(program, 0)["errorBars"]!.AsObject());
            Assert.False(CompileTrendlineList(program).Ok);
            var projected = ProjectTrendlineList(source);
            mutation(TrendlineListSeries(projected, 0)["errorBars"]!.AsObject());
            Assert.False(CompileTrendlineList(projected, source).Ok);
        }
        foreach (var kind in new[] { "string", "size" })
        {
            var program = valid.DeepClone().AsObject();
            SetErrorDataFormatToken(program, 0, "");
            program["design"]!["grammar"]!["tokens"]!["errorFormat"]!["kind"] = kind;
            if (kind == "size") program["design"]!["grammar"]!["tokens"]!["errorFormat"]!["value"] = 2;
            Assert.False(CompileTrendlineList(program).Ok);
            var projected = ProjectTrendlineList(source);
            SetErrorDataFormatToken(projected, 0, "");
            projected["design"]!["grammar"]!["tokens"] = program["design"]!["grammar"]!["tokens"]!.DeepClone();
            Assert.False(CompileTrendlineList(projected, source).Ok);
        }
    }

    [Fact]
    public void PpjCustomErrorDataPreservesMalformedCaches()
    {
        var program = TrendlineListProgram("line", 0);
        TrendlineListSeries(program, 0)["errorBars"] = JsonNode.Parse(CustomErrorData);
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var variant in new[] { "missing-point", "duplicate-index", "extension", "negative" })
        {
            var input = ReplaceZipText(source, path, text =>
            {
                var xml = XDocument.Parse(text);
                var cache = xml.Descendants(c + "plus").Single().Element(c + "numLit")!;
                if (variant == "missing-point") cache.Elements(c + "pt").Last().Remove();
                if (variant == "duplicate-index") cache.Elements(c + "pt").Last().SetAttributeValue("idx", "0");
                if (variant == "extension") cache.Add(new XElement(c + "extLst"));
                if (variant == "negative") cache.Descendants(c + "v").First().Value = "-1";
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var projected = ProjectTrendlineList(input);
            var capabilities = projected["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
                .SelectMany(element => element!["nativeRef"]?["capabilities"]?.AsArray() ?? new JsonArray());
            Assert.DoesNotContain(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setChartSeriesAnalytics");
            var preserved = CompileTrendlineList(projected, input);
            Assert.True(preserved.Ok, Diagnostics(preserved));
            Assert.Empty(preserved.PresentationProgram.ChangedParts);
            Assert.Equal(ZipBytes(input, path), ZipBytes(preserved.File.ToByteArray(), path));
        }
    }

    private static void SetErrorDataFormatToken(JsonObject program, int seriesIndex, string format)
    {
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["errorFormat"] = new JsonObject { ["kind"] = "string", ["value"] = format },
        };
        TrendlineListSeries(program, seriesIndex)["errorBars"]!["plus"]!["formatCode"] = new JsonObject { ["token"] = "errorFormat" };
    }
}
