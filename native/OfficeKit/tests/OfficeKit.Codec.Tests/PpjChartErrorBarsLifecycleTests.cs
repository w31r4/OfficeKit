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
    // Reuse the trendline fixture's two-series chart so every error-bar edit
    // also checks that neighboring analytics and the other series survive.
    [Theory]
    [InlineData("column", 0)]
    [InlineData("bar", 0)]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjErrorBarsLifecycleRoundTrips(string chartType, int seriesIndex)
    {
        var authored = CompileTrendlineList(TrendlineListProgram(chartType, seriesIndex));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectTrendlineList(source);
        Assert.Equal(1, TrendlineListSeries(projected, seriesIndex)["errorBars"]!["value"]!.GetValue<double>());
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var requested in new string?[]
        {
            null,
            """{"valueType":"fixed-value","value":0,"direction":"y","type":"both"}""",
            """{"valueType":"percentage","value":5,"direction":"x","type":"plus"}""",
            """{"valueType":"standard-deviation","value":1.5,"direction":"y","type":"minus","noEndCap":true,"stroke":{"color":"#2563EB","width":1.5,"dash":"solid"}}""",
            """{"valueType":"standard-error","direction":"y","type":"both"}""",
            null,
            """{"valueType":"fixed-value","value":2,"direction":"y","type":"both"}""",
        })
        {
            var series = TrendlineListSeries(projected, seriesIndex);
            if (requested is null) series.Remove("errorBars");
            else series["errorBars"] = JsonNode.Parse(requested);
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var beforeXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var afterXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var beforeSeries = beforeXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            var afterSeries = afterXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            if (requested is null) Assert.Empty(afterSeries.Elements(c + "errBars"));
            else
            {
                var native = Assert.Single(afterSeries.Elements(c + "errBars"));
                Assert.Contains(native.ElementsBeforeSelf(), element => element.Name == c + "trendline");
                Assert.Contains(native.ElementsAfterSelf(), element => element.Name == c + "cat");
                var expectedType = series["errorBars"]!["valueType"]!.GetValue<string>() switch
                {
                    "fixed-value" => "fixedVal", "percentage" => "percentage",
                    "standard-deviation" => "stdDev", _ => "stdErr",
                };
                Assert.Equal(expectedType, (string?)native.Element(c + "errValType")!.Attribute("val"));
                if (expectedType == "stdErr") Assert.Null(native.Element(c + "val"));
            }
            OpenXmlElement nativeElement = chartType is "column" or "bar"
                ? new C.BarChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting))
                : new C.LineChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(nativeElement));
            beforeSeries.Elements(c + "errBars").Remove();
            afterSeries.Elements(c + "errBars").Remove();
            Assert.True(XNode.DeepEquals(beforeXml, afterXml), "Only the selected errorBars owner may change inside ChartML.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            var recovered = TrendlineListSeries(projected, seriesIndex)["errorBars"];
            if (requested is null) Assert.Null(recovered);
            else Assert.True(JsonNode.DeepEquals(JsonNode.Parse(requested), recovered), $"Error bars lost requested semantics: {recovered}");
        }
    }

    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjErrorBarsLifecyclePreservesUnsupportedAndCustomOwners(string chartType, int seriesIndex)
    {
        var authored = CompileTrendlineList(TrendlineListProgram(chartType, seriesIndex));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var variant in new[] { "duplicate", "duplicate-direction", "extension", "invalid-value", "custom" })
        {
            var input = ReplaceZipText(source, path, text =>
            {
                var xml = XDocument.Parse(text);
                var owner = xml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable").Element(c + "errBars")!;
                switch (variant)
                {
                    case "duplicate": owner.AddAfterSelf(new XElement(owner)); break;
                    case "duplicate-direction": owner.AddFirst(new XElement(owner.Element(c + "errDir")!)); break;
                    case "extension": owner.Add(new XElement(c + "extLst")); break;
                    case "invalid-value": owner.Element(c + "val")!.SetAttributeValue("val", "-1"); break;
                    case "custom":
                        owner.Element(c + "errValType")!.SetAttributeValue("val", "cust");
                        owner.Element(c + "val")!.Remove();
                        foreach (var side in new[] { "plus", "minus" })
                            owner.Add(new XElement(c + side, new XElement(c + "numLit",
                                new XElement(c + "formatCode", "0.0"), new XElement(c + "ptCount", new XAttribute("val", 4)),
                                Enumerable.Range(0, 4).Select(index => new XElement(c + "pt", new XAttribute("idx", index), new XElement(c + "v", index + 1))))));
                        break;
                }
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var program = ProjectTrendlineList(input);
            var preserved = CompileTrendlineList(program, input);
            Assert.True(preserved.Ok, Diagnostics(preserved));
            Assert.Empty(preserved.PresentationProgram.ChangedParts);
            Assert.Equal(ZipBytes(input, path), ZipBytes(preserved.File.ToByteArray(), path));
            if (variant == "custom")
            {
                var series = TrendlineListSeries(program, seriesIndex);
                Assert.Null(series["errorBars"]);
                series["errorBars"] = JsonNode.Parse("""{"valueType":"fixed-value","value":1}""");
                var rejected = CompileTrendlineList(program, input);
                Assert.False(rejected.Ok);
                Assert.Contains("unprojected source error bars", Diagnostics(rejected));
            }
            else
            {
                var capabilities = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
                    .SelectMany(element => element!["nativeRef"]?["capabilities"]?.AsArray() ?? new JsonArray());
                Assert.DoesNotContain(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setChartSeriesAnalytics");
            }
        }
    }

    [Fact]
    public void PpjErrorBarsLifecycleRejectsInvalidInsertions()
    {
        var program = TrendlineListProgram("line", 0);
        TrendlineListSeries(program, 0).Remove("errorBars");
        var authored = CompileTrendlineList(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        foreach (var invalid in new[]
        {
            """{"valueType":"fixed-value","value":-1}""",
            """{"valueType":"fixed-value"}""",
            """{"valueType":"standard-error","value":1}""",
            """{"valueType":"custom"}""",
            """{"valueType":"fixed-value","value":1,"direction":"z"}""",
        })
        {
            var projected = ProjectTrendlineList(source);
            TrendlineListSeries(projected, 0)["errorBars"] = JsonNode.Parse(invalid);
            Assert.False(CompileTrendlineList(projected, source).Ok);
        }
    }
}
