using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Security.Cryptography;
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
    public void PpjSourceBoundPictureShape3dSceneCameraRotationLongitudeLeafEditsAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        if (root is null)
        {
            root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
                root = root.Parent;
        }
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var assetBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
        var assetSha256 = Convert.ToHexString(SHA256.HashData(assetBytes)).ToLowerInvariant();
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
                Assets =
                {
                    new Asset
                    {
                        Id = "evidence-mark",
                        FileName = "evidence-mark.svg",
                        ContentType = "image/svg+xml",
                        Data = ByteString.CopyFrom(assetBytes),
                        Sha256 = assetSha256,
                    },
                },
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));

        var source = AddPictureShape3dSceneCameraRotationLongitude(RemoveEmbeddedPpj(authored.File.ToByteArray()), "0");
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/picture-shape-3d-scene-camera-rotation-longitude.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));

        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedImage = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var target = Assert.Single(projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneCameraRotationLongitude60000"));
        Assert.Equal(0, target["value"]!.GetValue<long>());
        target["value"] = 1_800_000;

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
        Assert.Contains("<a:rot", editedXml, StringComparison.Ordinal);
        Assert.Contains("lat=\"0\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("lon=\"1800000\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("rev=\"6000000\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("prst=\"perspectiveFront\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("rig=\"threePt\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("dir=\"t\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("z=\"1200\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("extrusionH=\"2400\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("contourW=\"600\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("prstMaterial=\"metal\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        foreach (var pathName in ZipPartPaths(source).Where(pathName => !pathName.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, pathName), ZipBytes(editedBytes, pathName));

        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var picture = package.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>())
                .Single(item => item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var scene = picture.ShapeProperties!.GetFirstChild<A.Scene3DType>()!;
            var camera = scene.GetFirstChild<A.Camera>()!;
            Assert.Equal("perspectiveFront", camera.GetAttribute("prst", string.Empty).Value);
            var rotation = camera.GetFirstChild<A.Rotation>()!;
            Assert.Equal("0", rotation.GetAttribute("lat", string.Empty).Value);
            Assert.Equal("1800000", rotation.GetAttribute("lon", string.Empty).Value);
            Assert.Equal("6000000", rotation.GetAttribute("rev", string.Empty).Value);
            var lightRig = scene.GetFirstChild<A.LightRig>()!;
            Assert.Equal("threePt", lightRig.GetAttribute("rig", string.Empty).Value);
            Assert.Equal("t", lightRig.GetAttribute("dir", string.Empty).Value);
            var shape3d = picture.ShapeProperties.GetFirstChild<A.Shape3DType>()!;
            Assert.Equal("1200", shape3d.GetAttribute("z", string.Empty).Value);
            Assert.Equal("2400", shape3d.GetAttribute("extrusionH", string.Empty).Value);
            Assert.Equal("600", shape3d.GetAttribute("contourW", string.Empty).Value);
            Assert.Equal("metal", shape3d.GetAttribute("prstMaterial", string.Empty).Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(editedBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/picture-shape-3d-scene-camera-rotation-longitude.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var reprojectedTarget = Assert.Single(reprojectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneCameraRotationLongitude60000"));
        Assert.Equal(1_800_000, reprojectedTarget["value"]!.GetValue<long>());
    }

    private static byte[] AddPictureShape3dSceneCameraRotationLongitude(byte[] source, string longitude)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true))
        {
            var matches = package.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                    .Select(picture => (slidePart, picture)))
                .Where(item => item.picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity")
                .ToArray();
            Assert.Single(matches);
            var slidePart = matches[0].slidePart;
            var picture = matches[0].picture;
            var camera = new A.Camera();
            camera.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "perspectiveFront"));
            var rotation = new A.Rotation();
            rotation.SetAttribute(new OpenXmlAttribute("lat", string.Empty, "0"));
            rotation.SetAttribute(new OpenXmlAttribute("lon", string.Empty, longitude));
            rotation.SetAttribute(new OpenXmlAttribute("rev", string.Empty, "6000000"));
            camera.AppendChild(rotation);
            var lightRig = new A.LightRig();
            lightRig.SetAttribute(new OpenXmlAttribute("rig", string.Empty, "threePt"));
            lightRig.SetAttribute(new OpenXmlAttribute("dir", string.Empty, "t"));
            picture.ShapeProperties!.AppendChild(new A.Scene3DType(camera, lightRig));
            var shape3d = new A.Shape3DType();
            shape3d.SetAttribute(new OpenXmlAttribute("z", string.Empty, "1200"));
            shape3d.SetAttribute(new OpenXmlAttribute("extrusionH", string.Empty, "2400"));
            shape3d.SetAttribute(new OpenXmlAttribute("contourW", string.Empty, "600"));
            shape3d.SetAttribute(new OpenXmlAttribute("prstMaterial", string.Empty, "metal"));
            picture.ShapeProperties.AppendChild(shape3d);
            slidePart.Slide.Save();
        }
        return stream.ToArray();
    }
}
