using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundTableLastColumnEmphasisEditsLeafAndReprojects()
    {
        var request = ExportRequest();
        request.Artifact.Presentation.Slides[0].Elements.Clear();
        var table = new PresentationTable
        {
            LeftEmu = 500_000,
            TopEmu = 900_000,
            WidthEmu = 4_000_000,
            HeightEmu = 1_800_000,
            LastColumn = true,
        };
        table.ColumnWidthsEmu.Add([2_000_000, 2_000_000]);
        table.Rows.Add(new PresentationTableRow
        {
            HeightEmu = 600_000,
            Cells = { new PresentationTableCell { Text = "Metric" }, new PresentationTableCell { Text = "Value" } },
        });
        table.Rows.Add(new PresentationTableRow
        {
            HeightEmu = 1_200_000,
            Cells = { new PresentationTableCell { Text = "Revenue" }, new PresentationTableCell { Text = "$42M" } },
        });
        request.Artifact.Presentation.Slides[0].Elements.Add(new PresentationElement
        {
            Id = "presentation/slide/1/table/native-last-column",
            Name = "Native last-column table",
            Table = table,
        });

        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());

        using (var stream = new MemoryStream(sourceBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeTable = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>());
            Assert.True(nativeTable.TableProperties!.LastColumn!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "table-last-column/source.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTable = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "table");
        Assert.True(projectedTable["style"]!["lastColumnEmphasis"]!.GetValue<bool>());
        var leaves = projectedTable["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "tableLastColumnEmphasis")
            .ToArray();
        Assert.Single(leaves);
        Assert.True(leaves[0]["value"]!.GetValue<bool>());

        leaves[0]["value"] = false;
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
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);

        var editedBytes = edited.File.ToByteArray();
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeTable = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>());
            Assert.False(nativeTable.TableProperties!.LastColumn!.Value);
            Assert.Contains("Revenue", nativeTable.InnerText);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        foreach (var path in ZipPartPaths(sourceBytes).Where(path => !path.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(sourceBytes, path), ZipBytes(editedBytes, path));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "table-last-column/edited.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedTable = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["type"]!.GetValue<string>() == "table");
        Assert.False(reprojectedTable["style"]!["lastColumnEmphasis"]!.GetValue<bool>());
        Assert.False(reprojectedTable["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "tableLastColumnEmphasis")["value"]!.GetValue<bool>());
    }
}
