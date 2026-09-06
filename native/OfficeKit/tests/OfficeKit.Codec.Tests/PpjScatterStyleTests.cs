using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
    public void PpjScatterStyleAuthorProjectAndEditOnlyChartPart()
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
        program["assets"] = new JsonArray();
        program["components"] = new JsonArray();
        program["sections"] = new JsonArray();
        program["customShows"] = new JsonArray();
        program["comments"] = new JsonArray();
        var page = program["pages"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "page-evidence");
        page.Remove("animations");
        page.Remove("notes");
        page["elements"] = new JsonArray(page["elements"]!.AsArray()
            .Where(item => item!["type"]!.GetValue<string>() == "chart")
            .Select(item => item!.DeepClone())
            .ToArray());
        var chart = page["elements"]!.AsArray()[0]!.AsObject();
        chart.Remove("styleRef");
        chart["chartType"] = "scatter";
        chart["style"] = new JsonObject
        {
            ["legend"] = "none",
            ["scatterStyle"] = "smoothWithMarkers",
        };
        chart["data"] = new JsonObject
        {
            ["categories"] = new JsonArray(),
            ["series"] = new JsonArray(new JsonObject
            {
                ["id"] = "scatter-series",
                ["name"] = "Reach",
                ["xValues"] = new JsonArray(10, 20, 34),
                ["values"] = new JsonArray(35, 68, 84),
            }),
        };
        chart["xAxis"] = new JsonObject { ["title"] = "Reach" };
        chart["yAxis"] = new JsonObject { ["title"] = "Return" };
        program["pages"] = new JsonArray(page.DeepClone());

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

        string chartPath;
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            chartPath = Assert.Single(package.PresentationPart!.SlideParts.Single().ChartParts).Uri.OriginalString.TrimStart('/');
        }
        XNamespace chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var authoredChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(authored.File.ToByteArray(), chartPath)));
        Assert.Equal("smoothMarker", authoredChart.Descendants(chartNs + "scatterStyle").Single().Attribute("val")!.Value);

        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "scatter-style/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = state["pages"]!.AsArray()
            .SelectMany(item => item!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("smoothWithMarkers", projectedChart["style"]!["scatterStyle"]!.GetValue<string>());
        Assert.Contains(projectedChart["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setChartPlot");

        projectedChart["style"]!["scatterStyle"] = "lineWithMarkers";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(new[] { chartPath }, edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var editedChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), chartPath)));
        Assert.Equal("lineMarker", editedChart.Descendants(chartNs + "scatterStyle").Single().Attribute("val")!.Value);
        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "scatter-style/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var output = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var outputChart = output["pages"]!.AsArray()
            .SelectMany(item => item!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.Equal("lineWithMarkers", outputChart["style"]!["scatterStyle"]!.GetValue<string>());
    }
}
