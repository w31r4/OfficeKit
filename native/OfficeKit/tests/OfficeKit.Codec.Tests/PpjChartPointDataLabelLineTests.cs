using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjChartPointDataLabelLineAuthorAndEditSourceChart()
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
                    ["fill"] = new JsonObject
                    {
                        ["type"] = "solid",
                        ["color"] = "#FFE8A3",
                    },
                    ["line"] = new JsonObject
                    {
                        ["color"] = "#7B61FF",
                        ["width"] = 1.25,
                        ["dash"] = "dash",
                    },
                },
                new JsonObject
                {
                    ["index"] = 1,
                    ["line"] = new JsonObject
                    {
                        ["color"] = "#2F80ED",
                        ["width"] = 2.0,
                        ["dash"] = "dot",
                    },
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
        Assert.Contains("dLbl", authoredXml, StringComparison.Ordinal);
        Assert.Contains("<a:ln", authoredXml, StringComparison.Ordinal);
        Assert.Contains("prstDash", authoredXml, StringComparison.Ordinal);
        Assert.Contains("val=\"7B61FF\"", authoredXml, StringComparison.Ordinal);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-point-data-label-line/source.pptx" },
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
        Assert.Equal("#7B61FF", projectedPoint0["line"]!["color"]!.GetValue<string>());
        Assert.Equal("dash", projectedPoint0["line"]!["dash"]!.GetValue<string>());
        Assert.Equal("#2F80ED", projectedPoint1["line"]!["color"]!.GetValue<string>());
        Assert.Equal("dot", projectedPoint1["line"]!["dash"]!.GetValue<string>());

        projectedPoint0["line"]!["color"] = "#00FF00";
        projectedPoint1["line"]!["width"] = 3.0;
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
        Assert.Contains("val=\"00FF00\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("w=\"38100\"", editedXml, StringComparison.Ordinal);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-point-data-label-line/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var reprojectedPoints = reprojectedChart["data"]!["series"]![0]!["dataLabels"]!["points"]!.AsArray();
        var reprojectedPoint0 = reprojectedPoints.Single(item => item!["index"]!.GetValue<int>() == 0)!.AsObject();
        var reprojectedPoint1 = reprojectedPoints.Single(item => item!["index"]!.GetValue<int>() == 1)!.AsObject();
        Assert.Equal("#00FF00", reprojectedPoint0["line"]!["color"]!.GetValue<string>());
        Assert.Equal(3.0, reprojectedPoint1["line"]!["width"]!.GetValue<double>());
    }
}
