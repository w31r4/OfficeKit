using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundCustomGeometryPathLineToXEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var shape = authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape;
        shape.Geometry = "custom";
        var path = new PresentationCustomGeometryPath
        {
            Width = 1_000,
            Height = 1_000,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
        };
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            MoveTo = new PresentationCustomGeometryPoint { X = 100, Y = 100 },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            LineTo = new PresentationCustomGeometryPoint { X = 900, Y = 100 },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            LineTo = new PresentationCustomGeometryPoint { X = 500, Y = 900 },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        shape.CustomPaths.Add(path);

        var authored = Invoke(authoredRequest);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-geometry-path-line-to-x.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));

        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var lineLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "customGeometryPathLineToX")
            .ToArray();
        Assert.Equal(2, lineLeaves.Length);
        Assert.Contains(lineLeaves, leaf => leaf["value"]!.GetValue<long>() == 500);
        var target = Assert.Single(lineLeaves, leaf => leaf["value"]!.GetValue<long>() == 900);
        target["value"] = 850;

        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(projectedProgram.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);

        var editedBytes = edited.File.ToByteArray();
        var editedXml = Encoding.UTF8.GetString(ZipBytes(editedBytes, "ppt/slides/slide1.xml"));
        Assert.Contains("x=\"850\" y=\"100\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("x=\"500\" y=\"900\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("x=\"100\" y=\"100\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        foreach (var pathName in ZipPartPaths(source).Where(pathName => !pathName.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, pathName), ZipBytes(editedBytes, pathName));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(editedBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/custom-geometry-path-line-to-x.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var reprojectedLeaves = reprojectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "customGeometryPathLineToX")
            .ToArray();
        Assert.Contains(reprojectedLeaves, leaf => leaf["value"]!.GetValue<long>() == 850);
        Assert.Contains(reprojectedLeaves, leaf => leaf["value"]!.GetValue<long>() == 500);
    }
}
