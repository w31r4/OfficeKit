using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjChartPointDataLabelTextAuthorAndEditSourceChart()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        foreach (var page in program["pages"]!.AsArray())
        {
            var elements = page!["elements"]!.AsArray();
            foreach (var image in elements
                .Where(item => item!["type"]!.GetValue<string>() == "image")
                .ToArray())
                elements.Remove(image);
        }
        program["assets"] = new JsonArray();
        var namedChartStyle = program["design"]!["styles"]!["chart"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "evidence-chart")["style"]!.AsObject();
        namedChartStyle.Remove("showDataLabels");
        namedChartStyle.Remove("dataLabelPosition");
        var chart = program["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        chart.Remove("showDataLabels");
        chart.Remove("dataLabelPosition");
        var series = chart["data"]!["series"]![0]!.AsObject();
        series["dataLabels"] = new JsonObject
        {
            ["points"] = new JsonArray
            {
                new JsonObject
                {
                    ["index"] = 0,
                    ["text"] = "Incident 24H1",
                    ["fill"] = new JsonObject
                    {
                        ["type"] = "solid",
                        ["color"] = "#FFE8A3",
                    },
                },
                new JsonObject
                {
                    ["index"] = 1,
                    ["text"] = "Incident 24H2",
                },
            },
        };

        var programBytes = Encoding.UTF8.GetBytes(program.ToJsonString());
        var validation = PpjProgramValidator.Validate(programBytes);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics));
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFrom(programBytes),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));

        var authoredBytes = authored.File.ToByteArray();
        var sourceBytes = RemoveEmbeddedPpj(authoredBytes);
        var chartPaths = ZipPartPaths(sourceBytes).Where(path =>
            path.Contains("/charts/", StringComparison.OrdinalIgnoreCase) &&
            path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).ToArray();
        Assert.True(chartPaths.Length > 0, string.Join(",", ZipPartPaths(sourceBytes)));
        var chartPath = Assert.Single(chartPaths);
        var authoredXml = Encoding.UTF8.GetString(ZipBytes(authoredBytes, chartPath));
        Assert.Contains("<c:tx>", authoredXml, StringComparison.Ordinal);
        Assert.Contains("<c:rich>", authoredXml, StringComparison.Ordinal);
        Assert.Contains("Incident 24H1", authoredXml, StringComparison.Ordinal);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-point-data-label-text/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var projectedPoints = projectedChart["data"]!["series"]![0]!["dataLabels"]!["points"]!.AsArray();
        var projectedPoint0 = projectedPoints.Single(item => item!["index"]!.GetValue<int>() == 0)!.AsObject();
        var projectedPoint1 = projectedPoints.Single(item => item!["index"]!.GetValue<int>() == 1)!.AsObject();
        Assert.Equal("Incident 24H1", projectedPoint0["text"]!.GetValue<string>());
        Assert.Equal("Incident 24H2", projectedPoint1["text"]!.GetValue<string>());
        Assert.Equal("#FFE8A3", projectedPoint0["fill"]!["color"]!.GetValue<string>());

        projectedPoint0["text"] = "P1: 24H1";
        projectedPoint1["text"] = "P1: 24H2";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(projectedProgram.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal([chartPath], edited.PresentationProgram.ChangedParts);
        var editedXml = Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), chartPath));
        Assert.Contains("P1: 24H1", editedXml, StringComparison.Ordinal);
        Assert.Contains("P1: 24H2", editedXml, StringComparison.Ordinal);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-point-data-label-text/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var reprojectedPoints = reprojectedChart["data"]!["series"]![0]!["dataLabels"]!["points"]!.AsArray();
        Assert.Equal("P1: 24H1", reprojectedPoints.Single(item => item!["index"]!.GetValue<int>() == 0)!["text"]!.GetValue<string>());
        Assert.Equal("P1: 24H2", reprojectedPoints.Single(item => item!["index"]!.GetValue<int>() == 1)!["text"]!.GetValue<string>());
    }
}
