using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PptxContentComparisonIgnoresStateButRetainsOptionalPresence()
    {
        var original = new PresentationElement { Chart = new PresentationChart { RoundedCorners = false } };
        var requested = original.Clone();
        requested.Hidden = true;
        requested.Locked = true;
        Assert.True(PptxElementStateCodec.EqualExceptState(original, requested));
        requested.Chart.ClearRoundedCorners();
        Assert.False(PptxElementStateCodec.EqualExceptState(original, requested));
    }

    [Theory]
    [InlineData("column")]
    [InlineData("combo")]
    public void PpjChartRoundedCornersPreservesPresenceAcrossEdits(string chartType)
    {
        var source = RoundedCornersSource(chartType);
        var chartPath = SingleZipEntryPath(source, path => path.Contains("/charts/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
        var projected = ProjectRoundedCorners(source);
        Assert.False(RoundedCornersChart(projected)["roundedCorners"]!.GetValue<bool>());

        // Each edit is isolated: false -> absent -> false -> true -> absent.
        // A second field must not accidentally force the ChartPart to be written.
        foreach (bool? requested in new bool?[] { null, false, true, null })
        {
            var chart = RoundedCornersChart(projected);
            if (requested is { } value) chart["roundedCorners"] = value;
            else chart.Remove("roundedCorners");
            var edited = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.CompilePpjToPptx,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(source),
                PresentationProgram = new PresentationProgramRequest
                {
                    ProgramJson = ByteString.CopyFromUtf8(projected.ToJsonString()),
                },
            });
            Assert.True(edited.Ok, Diagnostics(edited));
            Assert.Equal(chartPath, Assert.Single(edited.PresentationProgram.ChangedParts).TrimStart('/'));
            var output = RemoveEmbeddedPpj(edited.File.ToByteArray());
            var xml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(output, chartPath)));
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var corners = xml.Root!.Elements(c + "roundedCorners").ToArray();
            if (requested is { } expected)
            {
                Assert.Equal(expected ? "1" : "0", (string?)Assert.Single(corners).Attribute("val"));
                Assert.Contains(corners[0].ElementsAfterSelf(), element => element.Name == c + "style");
            }
            else Assert.Empty(corners);

            // The receipt is backed by actual package content, not only metadata.
            Assert.Equal(ZipPartPaths(source), ZipPartPaths(output));
            foreach (var part in ZipPartPaths(source).Where(path => path != chartPath))
                Assert.Equal(ZipBytes(source, part), ZipBytes(output, part));
            source = output;
            projected = ProjectRoundedCorners(source);
            var recovered = RoundedCornersChart(projected);
            Assert.Equal(requested.HasValue, recovered.ContainsKey("roundedCorners"));
            if (requested is { } result) Assert.Equal(result, recovered["roundedCorners"]!.GetValue<bool>());
        }
    }

    [Theory]
    [InlineData("column")]
    [InlineData("combo")]
    public void PpjChartRoundedCornersRejectsAmbiguousNativeOwners(string chartType)
    {
        var source = RoundedCornersSource(chartType);
        var chartPath = SingleZipEntryPath(source, path => path.Contains("/charts/", StringComparison.Ordinal) && path.EndsWith(".xml", StringComparison.Ordinal));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (var invalid in new[]
        {
            "<c:roundedCorners val='invalid'/>",
            "<c:roundedCorners val='0'/><c:roundedCorners val='1'/>",
            "<c:roundedCorners val='0' custom='keep'/>",
            "<c:roundedCorners val='0'><c:extLst/></c:roundedCorners>",
        })
        {
            var input = ReplaceZipText(source, chartPath, text =>
            {
                var xml = XDocument.Parse(text);
                var wrapper = XElement.Parse($"<root xmlns:c='{c}'>{invalid}</root>");
                xml.Root!.Element(c + "roundedCorners")!.ReplaceWith(wrapper.Elements().ToArray());
                return xml.ToString(SaveOptions.DisableFormatting);
            });
            var program = ProjectRoundedCorners(input);
            var capabilities = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
                .SelectMany(element => element!["nativeRef"]?["capabilities"]?.AsArray() ?? new JsonArray());
            Assert.DoesNotContain(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setChartPlot");
            var unchanged = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.CompilePpjToPptx,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(input),
                PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
            });
            Assert.True(unchanged.Ok, Diagnostics(unchanged));
            Assert.Empty(unchanged.PresentationProgram.ChangedParts);
            Assert.Equal(ZipBytes(input, chartPath), ZipBytes(unchanged.File.ToByteArray(), chartPath));
        }
    }

    private static byte[] RoundedCornersSource(string chartType)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json"))) root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(root!.FullName,
            "test", "fixtures", "presentation", "evidence-ledger-canonical.ppj")))!.AsObject();
        foreach (var page in program["pages"]!.AsArray())
        {
            var elements = page!["elements"]!.AsArray();
            foreach (var element in elements.Where(item => item!["type"]!.GetValue<string>() == "image").ToArray())
                elements.Remove(element);
        }
        program["assets"] = new JsonArray();
        var chart = RoundedCornersChart(program);
        chart["chartType"] = chartType;
        chart["styleIndex"] = 12;
        chart["roundedCorners"] = false;
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
            chart["data"]!["series"]![0]!["chartType"] = "column";
            chart["data"]!["series"]![1]!["chartType"] = "line";
        }
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        return RemoveEmbeddedPpj(authored.File.ToByteArray());
    }

    private static JsonObject ProjectRoundedCorners(byte[] source)
    {
        var result = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "rounded-corners/source.pptx" },
        });
        Assert.True(result.Ok, Diagnostics(result));
        Assert.False(result.PresentationProgram.RestoredEmbeddedProgram);
        return JsonNode.Parse(result.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    }

    private static JsonObject RoundedCornersChart(JsonObject program) => program["pages"]!.AsArray()
        .SelectMany(page => page!["elements"]!.AsArray()).Select(item => item!.AsObject())
        .Single(item => item["type"]!.GetValue<string>() == "chart");
}
