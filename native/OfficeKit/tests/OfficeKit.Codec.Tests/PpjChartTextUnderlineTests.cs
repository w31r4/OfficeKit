using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjChartTextUnderlineAuthorAndEditSourceChart()
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
        chart["title"] = "Evidence chart";
        chart.Remove("showDataLabels");
        chart.Remove("dataLabelPosition");
        chart["style"] = new JsonObject
        {
            ["legend"] = "bottom",
            ["titleTextStyle"] = new JsonObject { ["underline"] = "single" },
            ["legendTextStyle"] = new JsonObject { ["underline"] = "double" },
            ["dataLabels"] = new JsonObject
            {
                ["showValue"] = true,
                ["showCategory"] = false,
                ["textStyle"] = new JsonObject { ["underline"] = "wavy" },
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
        var chartPaths = ZipPartPaths(sourceBytes).Where(path =>
            path.Contains("/charts/", StringComparison.OrdinalIgnoreCase) &&
            path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).ToArray();
        Assert.True(chartPaths.Length > 0, string.Join(",", ZipPartPaths(sourceBytes)));
        var chartPath = Assert.Single(chartPaths);
        var authoredXml = Encoding.UTF8.GetString(ZipBytes(authored.File.ToByteArray(), chartPath));
        Assert.Contains("u=\"sng\"", authoredXml, StringComparison.Ordinal);
        Assert.Contains("u=\"dbl\"", authoredXml, StringComparison.Ordinal);
        Assert.Contains("u=\"wavy\"", authoredXml, StringComparison.Ordinal);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-text-underline/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("single", projectedChart["style"]!["titleTextStyle"]!["underline"]!.GetValue<string>());
        Assert.Equal("double", projectedChart["style"]!["legendTextStyle"]!["underline"]!.GetValue<string>());
        Assert.Equal("wavy", projectedChart["style"]!["dataLabels"]!["textStyle"]!["underline"]!.GetValue<string>());

        projectedChart["style"]!["titleTextStyle"]!["underline"] = "double";
        projectedChart["style"]!["legendTextStyle"]!["underline"] = "single";
        projectedChart["style"]!["dataLabels"]!["textStyle"]!["underline"] = "none";
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
        Assert.Contains("u=\"dbl\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("u=\"sng\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("u=\"none\"", editedXml, StringComparison.Ordinal);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "chart-text-underline/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("double", reprojectedChart["style"]!["titleTextStyle"]!["underline"]!.GetValue<string>());
        Assert.Equal("single", reprojectedChart["style"]!["legendTextStyle"]!["underline"]!.GetValue<string>());
        Assert.Equal("none", reprojectedChart["style"]!["dataLabels"]!["textStyle"]!["underline"]!.GetValue<string>());
    }
}
