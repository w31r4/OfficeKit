using Google.Protobuf;
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
    [InlineData("column", "yAxis")]
    [InlineData("scatter", "xAxis")]
    [InlineData("bubble", "xAxis")]
    [InlineData("radar", "yAxis")]
    [InlineData("radar", "spokeAxis")]
    [InlineData("combo", "yAxis")]
    [InlineData("combo", "secondaryYAxis")]
    public void PpjAxisLogBaseAuthorsEditsAndRemoves(string chartType, string axisName)
    {
        var program = LogBaseProgram(chartType, axisName);
        if (chartType == "radar") axisName = "spokeAxis";
        var authored = CompileLogBase(program);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var chartPath = SingleZipEntryPath(source, path => path.Contains("/charts/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectLogBase(source);
        Assert.Equal(10, LogBaseChart(projected)[axisName]!["logBase"]!.GetValue<double>());
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var originalXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, chartPath)));
        var originalLog = Assert.Single(originalXml.Descendants(c + "logBase"));
        Assert.Equal("10", (string?)originalLog.Attribute("val"));
        Assert.Equal(c + "orientation", originalLog.ElementsAfterSelf().First().Name);
        var axisId = (string?)originalLog.Parent!.Parent!.Element(c + "axId")!.Attribute("val");

        foreach (double? requested in new double?[] { 2.5, null, 2, 1000, null })
        {
            var axis = LogBaseChart(projected)[axisName]!.AsObject();
            if (requested is { } value)
            {
                projected["design"]!["grammar"]!["tokens"] = new JsonObject
                {
                    ["scale"] = new JsonObject { ["kind"] = "size", ["value"] = value },
                };
                axis["logBase"] = new JsonObject { ["token"] = "scale" };
            }
            else axis.Remove("logBase");
            var edited = CompileLogBase(projected, source);
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(chartPath, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var beforeXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(source, chartPath)));
            var afterXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, chartPath)));
            var logs = afterXml.Descendants(c + "logBase").ToArray();
            if (requested is { } expected)
            {
                var log = Assert.Single(logs);
                Assert.Equal(expected, double.Parse((string)log.Attribute("val")!, CultureInfo.InvariantCulture));
                Assert.Equal(c + "orientation", log.ElementsAfterSelf().First().Name);
                Assert.Equal(axisId, (string?)log.Parent!.Parent!.Element(c + "axId")!.Attribute("val"));
            }
            else Assert.Empty(logs);
            beforeXml.Descendants(c + "logBase").Remove();
            afterXml.Descendants(c + "logBase").Remove();
            Assert.True(XNode.DeepEquals(beforeXml, afterXml), "Only logBase may change inside the ChartPart.");
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(path => path != chartPath))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectLogBase(source);
            var recovered = LogBaseChart(projected)[axisName]?.AsObject();
            Assert.Equal(requested.HasValue, recovered?.ContainsKey("logBase") == true);
            if (requested is { } result) Assert.Equal(result, recovered!["logBase"]!.GetValue<double>());
        }
    }

    [Fact]
    public void PpjAxisLogBaseRejectsInvalidAuthoredAndSourceEdits()
    {
        var authored = CompileLogBase(LogBaseProgram("column", "yAxis"));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        foreach (var sourceBound in new[] { false, true })
        foreach (var invalid in new[] { "baseLow", "baseHigh", "tokenLow", "wrongKind", "category", "min", "max" })
        {
            var program = sourceBound ? ProjectLogBase(source) : LogBaseProgram("column", "yAxis");
            var chart = LogBaseChart(program);
            var axis = chart["yAxis"]!.AsObject();
            switch (invalid)
            {
                case "baseLow": axis["logBase"] = 1; break;
                case "baseHigh": axis["logBase"] = 1001; break;
                case "tokenLow":
                case "wrongKind":
                    program["design"]!["grammar"]!["tokens"] = new JsonObject
                    {
                        ["bad"] = new JsonObject { ["kind"] = invalid == "wrongKind" ? "spacing" : "size", ["value"] = invalid == "tokenLow" ? 1 : 10 },
                    };
                    axis["logBase"] = new JsonObject { ["token"] = "bad" };
                    break;
                case "category": chart["xAxis"] = new JsonObject { ["logBase"] = 10 }; break;
                default: axis[invalid] = 0; break;
            }
            var rejected = CompileLogBase(program, sourceBound ? source : null);
            Assert.False(rejected.Ok, $"{invalid}, sourceBound={sourceBound}");
            Assert.NotEmpty(rejected.Diagnostics);
        }
    }

    [Theory]
    [InlineData("column", "yAxis")]
    [InlineData("combo", "secondaryYAxis")]
    public void PpjAxisLogBaseProtectsAmbiguousNativeOwners(string chartType, string axisName)
    {
        var authored = CompileLogBase(LogBaseProgram(chartType, axisName));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var chartPath = SingleZipEntryPath(source, path => path.Contains("/charts/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var invalid in new[]
        {
            "<c:logBase val='NaN'/>", "<c:logBase val='1'/>", "<c:logBase val='1001'/>",
            "<c:logBase val='10'/><c:logBase val='2'/>", "<c:logBase/>",
            "<c:logBase val='10' custom='keep'/>", "<c:logBase val='10'><c:extLst/></c:logBase>",
        })
        {
            var input = ReplaceZipText(source, chartPath, text =>
            {
                var xml = XDocument.Parse(text);
                xml.Descendants(c + "logBase").Single().ReplaceWith(XElement.Parse($"<root xmlns:c='{c}'>{invalid}</root>").Elements().ToArray());
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var projected = ProjectLogBase(input);
            var capabilities = projected["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
                .SelectMany(element => element!["nativeRef"]?["capabilities"]?.AsArray() ?? new JsonArray());
            Assert.DoesNotContain(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setChartAxis");
            var unchanged = CompileLogBase(projected, input);
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Empty(unchanged.PresentationProgram.ChangedParts);
            Assert.Equal(ZipBytes(input, chartPath), ZipBytes(unchanged.File.ToByteArray(), chartPath));
        }
    }

    private static JsonObject LogBaseProgram(string chartType, string axisName)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json"))) root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(root!.FullName,
            "test", "fixtures", "presentation", "evidence-ledger-canonical.ppj")))!.AsObject();
        foreach (var page in program["pages"]!.AsArray())
        {
            page!.AsObject().Remove("animations");
            var elements = page!["elements"]!.AsArray();
            foreach (var element in elements.Where(item => item!["type"]!.GetValue<string>() == "image").ToArray()) elements.Remove(element);
        }
        program["assets"] = new JsonArray();
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["scale"] = new JsonObject { ["kind"] = "size", ["value"] = 10 },
        };
        var chart = LogBaseChart(program);
        chart["chartType"] = chartType;
        chart.Remove("styleRef");
        chart["style"] = new JsonObject { ["legend"] = "right" };
        chart["data"] = JsonNode.Parse("""
            {"categories":["A","B"],"series":[
              {"id":"one","name":"One","values":[2,4]},
              {"id":"two","name":"Two","values":[3,5]}
            ]}
            """);
        if (chartType == "combo")
        {
            chart["secondaryXAxis"] = new JsonObject();
            chart["secondaryYAxis"] = new JsonObject();
            chart["data"]!["series"]![0]!["chartType"] = "column";
            chart["data"]!["series"]![1]!["chartType"] = "line";
            chart["data"]!["series"]![1]!["axis"] = "secondary";
        }
        if (chartType is "scatter" or "bubble")
        {
            chart["data"]!["categories"] = new JsonArray();
            foreach (var series in chart["data"]!["series"]!.AsArray())
            {
                series!["xValues"] = new JsonArray(1, 2);
                if (chartType == "bubble") series["bubbleSizes"] = new JsonArray(4, 9);
            }
        }
        chart[axisName] = new JsonObject
        {
            ["logBase"] = new JsonObject { ["token"] = "scale" },
            ["min"] = 1, ["max"] = 100,
        };
        return program;
    }

    private static CodecResponse CompileLogBase(JsonObject program, byte[]? source = null) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.CompilePpjToPptx,
        Family = ArtifactFamily.Presentation,
        File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
    });

    private static JsonObject ProjectLogBase(byte[] source)
    {
        var result = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "log-base/source.pptx" },
        });
        Assert.True(result.Ok, Diagnostics(result));
        Assert.False(result.PresentationProgram.RestoredEmbeddedProgram);
        return JsonNode.Parse(result.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    }

    private static JsonObject LogBaseChart(JsonObject program) => program["pages"]!.AsArray()
        .SelectMany(page => page!["elements"]!.AsArray()).Select(item => item!.AsObject())
        .Single(item => item["type"]!.GetValue<string>() == "chart");
}
