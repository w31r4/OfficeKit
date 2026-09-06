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
    public void PpjSourceBoundPictureShape3dSceneBackdropAnchorZLeafEditsAndReprojects()
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
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(fixtureDirectory, "evidence-ledger-canonical.ppj")))!.AsObject();
        var assetBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
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
                        Sha256 = Convert.ToHexString(SHA256.HashData(assetBytes)).ToLowerInvariant(),
                    },
                },
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));

        var source = AddPictureShape3dSceneBackdropAnchorZ(RemoveEmbeddedPpj(authored.File.ToByteArray()), "200");
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/picture-backdrop-z.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var image = projectedProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject()).Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var target = Assert.Single(image["nativeRef"]!["leaves"]!.AsArray().Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneBackdropAnchorZEmu"));
        Assert.Equal(200, target["value"]!.GetValue<long>());
        target["value"] = 500;

        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(projectedProgram.ToJsonString()) },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        var editedBytes = edited.File.ToByteArray();
        var xml = Encoding.UTF8.GetString(ZipBytes(editedBytes, "ppt/slides/slide1.xml"));
        Assert.Contains("<a:anchor x=\"0\" y=\"100\" z=\"500\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:norm dx=\"1\" dy=\"2\" dz=\"3\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:up dx=\"4\" dy=\"5\" dz=\"6\"", xml, StringComparison.Ordinal);
        Assert.Contains("prst=\"orthographicFront\"", xml, StringComparison.Ordinal);
        Assert.Contains("rig=\"threePt\"", xml, StringComparison.Ordinal);
        Assert.Contains("z=\"1200\"", xml, StringComparison.Ordinal);
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
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/edited/picture-backdrop-z.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject()).Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var reprojectedTarget = Assert.Single(reprojectedImage["nativeRef"]!["leaves"]!.AsArray().Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneBackdropAnchorZEmu"));
        Assert.Equal(500, reprojectedTarget["value"]!.GetValue<long>());
    }

    private static byte[] AddPictureShape3dSceneBackdropAnchorZ(byte[] source, string anchorZ)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true))
        {
            var match = package.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                    .Select(picture => (slidePart, picture)))
                .Single(item => item.picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var camera = new A.Camera();
            camera.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "orthographicFront"));
            var lightRig = new A.LightRig();
            lightRig.SetAttribute(new OpenXmlAttribute("rig", string.Empty, "threePt"));
            lightRig.SetAttribute(new OpenXmlAttribute("dir", string.Empty, "t"));
            var anchor = new A.Anchor();
            anchor.SetAttribute(new OpenXmlAttribute("x", string.Empty, "0"));
            anchor.SetAttribute(new OpenXmlAttribute("y", string.Empty, "100"));
            anchor.SetAttribute(new OpenXmlAttribute("z", string.Empty, anchorZ));
            var normal = new A.Normal();
            normal.SetAttribute(new OpenXmlAttribute("dx", string.Empty, "1"));
            normal.SetAttribute(new OpenXmlAttribute("dy", string.Empty, "2"));
            normal.SetAttribute(new OpenXmlAttribute("dz", string.Empty, "3"));
            var up = new A.UpVector();
            up.SetAttribute(new OpenXmlAttribute("dx", string.Empty, "4"));
            up.SetAttribute(new OpenXmlAttribute("dy", string.Empty, "5"));
            up.SetAttribute(new OpenXmlAttribute("dz", string.Empty, "6"));
            match.picture.ShapeProperties!.AppendChild(new A.Scene3DType(camera, lightRig, new A.Backdrop(anchor, normal, up)));
            var shape3d = new A.Shape3DType();
            shape3d.SetAttribute(new OpenXmlAttribute("z", string.Empty, "1200"));
            shape3d.SetAttribute(new OpenXmlAttribute("extrusionH", string.Empty, "2400"));
            shape3d.SetAttribute(new OpenXmlAttribute("contourW", string.Empty, "600"));
            shape3d.SetAttribute(new OpenXmlAttribute("prstMaterial", string.Empty, "metal"));
            match.picture.ShapeProperties.AppendChild(shape3d);
            match.slidePart.Slide.Save();
        }
        return stream.ToArray();
    }
}
