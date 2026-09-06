using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundPictureShape3dSceneBackdropUpDxLeafEditsAndReprojects()
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

        var source = AddPictureShape3dSceneBackdropAnchorY(RemoveEmbeddedPpj(authored.File.ToByteArray()), "100");
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/picture-backdrop-up-dx.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var image = projectedProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject()).Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var target = Assert.Single(image["nativeRef"]!["leaves"]!.AsArray().Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneBackdropUpDxEmu"));
        Assert.Equal(4, target["value"]!.GetValue<long>());
        target["value"] = 5;

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
        Assert.Contains("<a:up dx=\"5\" dy=\"5\" dz=\"6\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:anchor x=\"0\" y=\"100\" z=\"200\"", xml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(editedBytes),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/edited/picture-backdrop-up-dx.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject()).Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var reprojectedTarget = Assert.Single(reprojectedImage["nativeRef"]!["leaves"]!.AsArray().Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dSceneBackdropUpDxEmu"));
        Assert.Equal(5, reprojectedTarget["value"]!.GetValue<long>());
    }
}
