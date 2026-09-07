using Google.Protobuf;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("pie")]
    [InlineData("doughnut")]
    public void PpjChartSeriesExplosionAuthorAndEditSourceChart(string chartType)
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
        chart.Remove("styleRef");
        chart["chartType"] = chartType;
        foreach (var property in new[]
        {
            "xAxis",
            "yAxis",
            "secondaryXAxis",
            "secondaryYAxis",
            "spokeAxis",
            "dataTable",
            "displayBlanksAs",
        })
            chart.Remove(property);
        chart["data"] = new JsonObject
        {
            ["categories"] = new JsonArray("Direct", "Partner", "Trial"),
            ["series"] = new JsonArray(new JsonObject
            {
                ["id"] = "series-explosion",
                ["name"] = "Revenue",
                ["values"] = new JsonArray(24, 40, 36),
                ["explosion"] = 24,
            }),
        };
        chart["style"] = new JsonObject { ["legend"] = "right" };

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
            chartPath = Assert.Single(package.PresentationPart!.SlideParts.SelectMany(slide => slide.ChartParts))
                .Uri.OriginalString.TrimStart('/');
        }
        XNamespace chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var plotName = chartType == "pie" ? "pieChart" : "doughnutChart";
        var authoredChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(authored.File.ToByteArray(), chartPath)));
        Assert.Equal(
            "24",
            authoredChart.Descendants(chartNs + plotName).Single()
                .Elements(chartNs + "ser").Single()
                .Element(chartNs + "explosion")!.Attribute("val")!.Value);

        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = $"chart-series-explosion-{chartType}/source.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var projectedSeries = projectedChart["data"]!["series"]![0]!.AsObject();
        Assert.Equal(24, projectedSeries["explosion"]!.GetValue<int>());

        projectedSeries["explosion"] = 0;
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
        Assert.Equal(new[] { chartPath }, edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        var editedChart = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), chartPath)));
        Assert.Equal(
            "0",
            editedChart.Descendants(chartNs + plotName).Single()
                .Elements(chartNs + "ser").Single()
                .Element(chartNs + "explosion")!.Attribute("val")!.Value);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = $"chart-series-explosion-{chartType}/edited.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        var reprojectedSeries = reprojectedChart["data"]!["series"]![0]!.AsObject();
        Assert.Equal(0, reprojectedSeries["explosion"]!.GetValue<int>());
    }
}
