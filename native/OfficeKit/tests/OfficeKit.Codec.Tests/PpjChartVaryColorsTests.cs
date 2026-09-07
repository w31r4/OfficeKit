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
    [InlineData("combo")]
    [InlineData("column")]
    [InlineData("area")]
    [InlineData("scatter")]
    public void PpjChartVaryColorsAuthorAndEditSourceChart(string chartType)
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
        if (chartType is "column" or "area")
        {
            chart.Remove("styleRef");
            chart["chartType"] = chartType;
            foreach (var series in chart["data"]!["series"]!.AsArray())
            {
                var seriesObject = series!.AsObject();
                seriesObject.Remove("chartType");
                seriesObject.Remove("axis");
                seriesObject.Remove("marker");
            }
        }
        else if (chartType == "scatter")
        {
            chart.Remove("styleRef");
            chart["chartType"] = "scatter";
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
        }
        chart["style"] = new JsonObject
        {
            ["legend"] = "right",
            ["varyColors"] = true,
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

        string? nativeChartPath = null;
        if (chartType is "scatter" or "area")
        {
            using var stream = new MemoryStream(authored.File.ToByteArray(), writable: false);
            using var package = PresentationDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            nativeChartPath = Assert.Single(package.PresentationPart!.SlideParts.SelectMany(slide => slide.ChartParts)).Uri.OriginalString.TrimStart('/');
            var chartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(authored.File.ToByteArray(), nativeChartPath)));
            XNamespace chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var plotName = chartType == "scatter" ? "scatterChart" : "areaChart";
            Assert.Equal("1", chartXml.Descendants(chartNs + plotName).Single().Element(chartNs + "varyColors")!.Attribute("val")!.Value);
        }

        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = $"chart-vary-colors-{chartType}/source.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedChart = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.True(projectedChart["style"]!["varyColors"]!.GetValue<bool>());

        projectedChart["style"]!["varyColors"] = false;
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
        if (nativeChartPath is not null)
        {
            Assert.Equal(new[] { nativeChartPath }, edited.PresentationProgram.ChangedParts);
            using var stream = new MemoryStream(edited.File.ToByteArray(), writable: false);
            using var package = PresentationDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var chartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), nativeChartPath)));
            XNamespace chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var plotName = chartType == "scatter" ? "scatterChart" : "areaChart";
            Assert.Equal("0", chartXml.Descendants(chartNs + plotName).Single().Element(chartNs + "varyColors")!.Attribute("val")!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = $"chart-vary-colors-{chartType}/edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedChart = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "chart");
        Assert.False(reprojectedChart["style"]!["varyColors"]!.GetValue<bool>());
    }
}
