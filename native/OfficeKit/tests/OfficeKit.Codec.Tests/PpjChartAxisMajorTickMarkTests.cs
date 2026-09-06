using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjChartAxisMajorTickMarksAuthorAndEditSourceChart()
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
        var grammar = program["design"]!["grammar"]!.AsObject();
        grammar["tokens"] = new JsonObject
        {
            ["xMajorTickMark"] = new JsonObject { ["kind"] = "string", ["value"] = "out" },
            ["yMajorTickMark"] = new JsonObject { ["kind"] = "string", ["value"] = "in" },
        };
        foreach (var page in program["pages"]!.AsArray())
        {
            var elements = page!["elements"]!.AsArray();
            foreach (var image in elements
                .Where(item => item!["type"]!.GetValue<string>() == "image")
                .ToArray())
                elements.Remove(image);
        }
        program["assets"] = new JsonArray();
        var chartStyle = program["design"]!["styles"]!["chart"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "evidence-chart")["style"]!.AsObject();
        chartStyle.Remove("showCategoryAxis");
        chartStyle.Remove("showValueAxis");
        chartStyle.Remove("showGridlines");
        var chart = program["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        chart["xAxis"] = new JsonObject
        {
            ["majorTickMark"] = new JsonObject { ["token"] = "xMajorTickMark" },
        };
        chart["yAxis"] = new JsonObject
        {
            ["majorTickMark"] = new JsonObject { ["token"] = "yMajorTickMark" },
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
            PresentationProgram = new PresentationProgramRequest { SourceUri = "major-tick-mark/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("out", projectedChart["xAxis"]!["majorTickMark"]!.GetValue<string>());
        Assert.Equal("in", projectedChart["yAxis"]!["majorTickMark"]!.GetValue<string>());

        projectedProgram["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["xMajorTickMarkEdited"] = new JsonObject { ["kind"] = "string", ["value"] = "none" },
            ["yMajorTickMarkEdited"] = new JsonObject { ["kind"] = "string", ["value"] = "cross" },
        };
        projectedChart["xAxis"]!["majorTickMark"] = new JsonObject { ["token"] = "xMajorTickMarkEdited" };
        projectedChart["yAxis"]!["majorTickMark"] = new JsonObject { ["token"] = "yMajorTickMarkEdited" };
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
            PresentationProgram = new PresentationProgramRequest { SourceUri = "major-tick-mark/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("none", reprojectedChart["xAxis"]!["majorTickMark"]!.GetValue<string>());
        Assert.Equal("cross", reprojectedChart["yAxis"]!["majorTickMark"]!.GetValue<string>());
    }
}
