using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("column", 0)]
    [InlineData("bar", 0)]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineListSupportsSourceBoundLifecycle(string chartType, int seriesIndex)
    {
        var authored = CompileTrendlineList(TrendlineListProgram(chartType, seriesIndex));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectTrendlineList(source);
        Assert.Equal(new[] { "Fit", "Curve" }, TrendlineListSeries(projected, seriesIndex)["trendlines"]!.AsArray().Select(item => item!["name"]!.GetValue<string>()));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        // Insert in the middle, reorder, shrink, omit, add from empty, clear
        // explicitly, and add again. Each compile changes only this list.
        foreach (var requested in new string?[]
        {
            "[{\"type\":\"linear\",\"name\":\"Fit\"},{\"type\":\"exponential\",\"name\":\"Growth\"},{\"type\":\"polynomial\",\"order\":2,\"name\":\"Curve\"}]",
            "[{\"type\":\"polynomial\",\"order\":3,\"name\":\"Curve\"},{\"type\":\"linear\",\"name\":\"Fit\"},{\"type\":\"power\",\"name\":\"Power\"}]",
            "[{\"type\":\"moving-average\",\"period\":2,\"name\":\"Average\"}]",
            null,
            "[{\"type\":\"logarithmic\",\"name\":\"Log\"}]",
            "[]",
            "[{\"type\":\"linear\",\"name\":\"Recreated\",\"forward\":0.5,\"displayEquation\":true,\"stroke\":{\"color\":\"#2563EB\",\"width\":1.5,\"dash\":\"solid\"}}]",
        })
        {
            var series = TrendlineListSeries(projected, seriesIndex);
            if (requested is null) series.Remove("trendlines");
            else series["trendlines"] = JsonNode.Parse(requested);
            var expected = series["trendlines"]?.DeepClone().AsArray() ?? new JsonArray();
            var edited = CompileTrendlineList(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(path, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var beforeXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, path)));
            var afterXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, path)));
            var beforeSeries = beforeXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            var afterSeries = afterXml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable");
            var native = afterSeries.Elements(c + "trendline").ToArray();
            Assert.Equal(expected.Select(item => item!["name"]!.GetValue<string>()), native.Select(item => item.Element(c + "name")!.Value));
            foreach (var trendline in native)
            {
                Assert.Contains(trendline.ElementsAfterSelf(), element => element.Name == c + "errBars");
                Assert.Contains(trendline.ElementsAfterSelf(), element => element.Name == c + "cat");
            }
            OpenXmlElement nativeElement = chartType is "column" or "bar"
                ? new C.BarChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting))
                : new C.LineChartSeries(afterSeries.ToString(SaveOptions.DisableFormatting));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(nativeElement));
            beforeSeries.Elements(c + "trendline").Remove();
            afterSeries.Elements(c + "trendline").Remove();
            Assert.True(XNode.DeepEquals(beforeXml, afterXml), "Only the selected series trendline list may change inside ChartML.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(part => part != path))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectTrendlineList(source);
            var recovered = TrendlineListSeries(projected, seriesIndex)["trendlines"]?.AsArray() ?? new JsonArray();
            Assert.Equal(expected.Count, recovered.Count);
            for (var index = 0; index < expected.Count; index++)
                foreach (var property in expected[index]!.AsObject())
                    Assert.True(JsonNode.DeepEquals(property.Value, recovered[index]![property.Key]), $"Trendline {index} lost {property.Key}.");
        }
    }

    [Theory]
    [InlineData("line", 0)]
    [InlineData("combo", 1)]
    public void PpjTrendlineListPreservesUnsupportedNativeOwners(string chartType, int seriesIndex)
    {
        var authored = CompileTrendlineList(TrendlineListProgram(chartType, seriesIndex));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var path = SingleZipEntryPath(source, name => name.Contains("/charts/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var invalid in new[]
        {
            "<c:trendlineType val='linear'/><c:trendlineLbl/>",
            "<c:trendlineType val='linear'/><c:extLst/>",
            "<c:trendlineType val='linear'/><c:trendlineType val='power'/>",
            "<c:trendlineType val='unknown'/>",
        })
        {
            var input = ReplaceZipText(source, path, text =>
            {
                var xml = XDocument.Parse(text);
                xml.Descendants(c + "ser").Single(item => item.Element(c + "tx")!.Value == "Editable")
                    .Elements(c + "trendline").First().ReplaceNodes(XElement.Parse($"<root xmlns:c='{c}'>{invalid}</root>").Elements().ToArray());
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var program = ProjectTrendlineList(input);
            var capabilities = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
                .SelectMany(element => element!["nativeRef"]?["capabilities"]?.AsArray() ?? new JsonArray());
            Assert.DoesNotContain(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setChartSeriesAnalytics");
            var preserved = CompileTrendlineList(program, input);
            Assert.True(preserved.Ok, Diagnostics(preserved));
            Assert.Empty(preserved.PresentationProgram.ChangedParts);
            Assert.Equal(ZipBytes(input, path), ZipBytes(preserved.File.ToByteArray(), path));
        }
    }

    [Fact]
    public void PpjTrendlineListRejectsInvalidInsertions()
    {
        var authored = CompileTrendlineList(TrendlineListProgram("line", 0));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        foreach (var invalid in new[] { "{\"type\":\"polynomial\"}", "{\"type\":\"moving-average\",\"period\":99}", "{\"type\":\"linear\",\"forward\":0.25}" })
        {
            var program = ProjectTrendlineList(source);
            TrendlineListSeries(program, 0)["trendlines"]!.AsArray().Add(JsonNode.Parse(invalid));
            Assert.False(CompileTrendlineList(program, source).Ok);
        }
        var tooMany = ProjectTrendlineList(source);
        TrendlineListSeries(tooMany, 0)["trendlines"] = new JsonArray(Enumerable.Range(0, 17)
            .Select(_ => (JsonNode)new JsonObject { ["type"] = "linear" }).ToArray());
        Assert.False(CompileTrendlineList(tooMany, source).Ok);
    }

    private static JsonObject TrendlineListProgram(string chartType, int seriesIndex)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json"))) root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(root!.FullName,
            "test", "fixtures", "presentation", "evidence-ledger-canonical.ppj")))!.AsObject();
        foreach (var page in program["pages"]!.AsArray())
        {
            page!.AsObject().Remove("animations");
            var elements = page["elements"]!.AsArray();
            foreach (var element in elements.Where(item => item!["type"]!.GetValue<string>() == "image").ToArray()) elements.Remove(element);
        }
        program["assets"] = new JsonArray();
        var chart = TrendlineListChart(program);
        chart["chartType"] = chartType;
        chart.Remove("styleRef");
        chart["style"] = new JsonObject { ["legend"] = "right" };
        chart["data"] = JsonNode.Parse("""
            {"categories":["A","B","C","D"],"series":[
              {"id":"one","name":"Unchanged","values":[2,4,6,8],"trendlines":[{"type":"linear","name":"Keep"}]},
              {"id":"two","name":"Unchanged","values":[3,5,7,9],"trendlines":[{"type":"linear","name":"Keep"}]}
            ]}
            """);
        if (chartType == "combo")
        {
            chart["data"]!["series"]![0]!["chartType"] = "column";
            chart["data"]!["series"]![1]!["chartType"] = "line";
            chart["data"]!["series"]![1]!["axis"] = "secondary";
            chart["secondaryXAxis"] = new JsonObject();
            chart["secondaryYAxis"] = new JsonObject();
        }
        var series = TrendlineListSeries(program, seriesIndex);
        series["name"] = "Editable";
        series["trendlines"] = JsonNode.Parse("""[{"type":"linear","name":"Fit"},{"type":"polynomial","order":2,"name":"Curve"}]""");
        series["errorBars"] = JsonNode.Parse("""{"valueType":"fixed-value","value":1,"direction":"y","type":"both"}""");
        return program;
    }

    private static CodecResponse CompileTrendlineList(JsonObject program, byte[]? source = null) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.CompilePpjToPptx,
        Family = ArtifactFamily.Presentation,
        File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
    });

    private static JsonObject ProjectTrendlineList(byte[] source)
    {
        var result = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "trendline-list/source.pptx" },
        });
        Assert.True(result.Ok, Diagnostics(result));
        Assert.False(result.PresentationProgram.RestoredEmbeddedProgram);
        return JsonNode.Parse(result.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    }

    private static JsonObject TrendlineListChart(JsonObject program) => program["pages"]!.AsArray()
        .SelectMany(page => page!["elements"]!.AsArray()).Select(item => item!.AsObject())
        .Single(item => item["type"]!.GetValue<string>() == "chart");

    private static JsonObject TrendlineListSeries(JsonObject program, int index) => TrendlineListChart(program)["data"]!["series"]![index]!.AsObject();
}
