using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundShape3dExtrusionHeightLeafEditsAndReprojects()
    {
        var authored = Invoke(ExportRequest());
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = AddShape3dExtrusionHeight(RemoveEmbeddedPpj(authored.File.ToByteArray()), 2_400);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/shape-3d-extrusion-height.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));

        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var leaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dExtrusionHeightEmu")
            .ToArray();
        var target = Assert.Single(leaves);
        Assert.Equal(2_400, target["value"]!.GetValue<long>());
        target["value"] = 3_600;

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
        Assert.Contains("extrusionH=\"3600\"", editedXml, StringComparison.Ordinal);
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
                SourceUri = "deck.assets/edited/shape-3d-extrusion-height.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var reprojectedLeaves = reprojectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dExtrusionHeightEmu")
            .ToArray();
        Assert.Equal(3_600, Assert.Single(reprojectedLeaves)["value"]!.GetValue<long>());
    }

    private static byte[] AddShape3dExtrusionHeight(byte[] source, long height)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true))
        {
            var slidePart = package.PresentationPart!.SlideParts.Single();
            var shape = slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            shape.ShapeProperties!.AppendChild(new A.Shape3DType
            {
                ExtrusionHeight = height,
            });
            slidePart.Slide.Save();
        }
        return stream.ToArray();
    }
}
