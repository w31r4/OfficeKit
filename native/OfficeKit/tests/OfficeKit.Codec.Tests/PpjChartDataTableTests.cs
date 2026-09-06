using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjChartDataTableAuthorAndEditSourceChart()
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
        var chart = program["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        chart["dataTable"] = new JsonObject
        {
            ["showHorizontalBorder"] = true,
            ["showVerticalBorder"] = false,
            ["showOutlineBorder"] = true,
            ["showLegendKey"] = false,
            ["fill"] = new JsonObject
            {
                ["type"] = "solid",
                ["color"] = "#EAF2F8",
                ["opacity"] = 0.75,
            },
            ["stroke"] = new JsonObject
            {
                ["color"] = "#7B61FF",
                ["width"] = 1.25,
                ["dash"] = "dash",
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

        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-data-table/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var projectedDataTable = projectedChart["dataTable"]!.AsObject();
        Assert.True(projectedDataTable["showHorizontalBorder"]!.GetValue<bool>());
        Assert.False(projectedDataTable["showVerticalBorder"]!.GetValue<bool>());
        Assert.True(projectedDataTable["showOutlineBorder"]!.GetValue<bool>());
        Assert.False(projectedDataTable["showLegendKey"]!.GetValue<bool>());
        Assert.Equal("solid", projectedDataTable["fill"]!["type"]!.GetValue<string>());
        Assert.Equal("#EAF2F8", projectedDataTable["fill"]!["color"]!.GetValue<string>());
        Assert.Equal(0.75, projectedDataTable["fill"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal("#7B61FF", projectedDataTable["stroke"]!["color"]!.GetValue<string>());
        Assert.Equal(1.25, projectedDataTable["stroke"]!["width"]!.GetValue<double>(), precision: 6);
        Assert.Equal("dash", projectedDataTable["stroke"]!["dash"]!.GetValue<string>());

        projectedDataTable["showHorizontalBorder"] = false;
        projectedDataTable["showVerticalBorder"] = true;
        projectedDataTable["showOutlineBorder"] = false;
        projectedDataTable["showLegendKey"] = true;
        projectedDataTable["fill"]!["color"] = "#D9EAD3";
        projectedDataTable["fill"]!["opacity"] = 0.5;
        projectedDataTable["stroke"]!["color"] = "#0B8F8F";
        projectedDataTable["stroke"]!["width"] = 2.5;
        projectedDataTable["stroke"]!["dash"] = "dot";
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
        Assert.Single(edited.PresentationProgram.ChangedParts);
        Assert.Contains("/charts/", edited.PresentationProgram.ChangedParts[0], StringComparison.Ordinal);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-data-table/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var reprojectedDataTable = reprojectedChart["dataTable"]!.AsObject();
        Assert.False(reprojectedDataTable["showHorizontalBorder"]!.GetValue<bool>());
        Assert.True(reprojectedDataTable["showVerticalBorder"]!.GetValue<bool>());
        Assert.False(reprojectedDataTable["showOutlineBorder"]!.GetValue<bool>());
        Assert.True(reprojectedDataTable["showLegendKey"]!.GetValue<bool>());
        Assert.Equal("#D9EAD3", reprojectedDataTable["fill"]!["color"]!.GetValue<string>());
        Assert.Equal(0.5, reprojectedDataTable["fill"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal("#0B8F8F", reprojectedDataTable["stroke"]!["color"]!.GetValue<string>());
        Assert.Equal(2.5, reprojectedDataTable["stroke"]!["width"]!.GetValue<double>(), precision: 6);
        Assert.Equal("dot", reprojectedDataTable["stroke"]!["dash"]!.GetValue<string>());
    }
}
