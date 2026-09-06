using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundImageShadowDirectionEditsLeafAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["shadow"]!.AsObject()["rotateWithShape"] = true;

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

        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var picture = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Picture>())
                .Single(item => item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var shadow = Assert.IsType<A.OuterShadow>(Assert.Single(picture.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.True(shadow.RotateWithShape!.Value);
            Assert.Equal(2_700_000, shadow.Direction!.Value);
            Assert.Equal(76_200U, shadow.BlurRadius!.Value);
            Assert.Equal(38_100U, shadow.Distance!.Value);
        }

        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/image-shadow-direction.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(45d, projectedImage["shadow"]!["angle"]!.GetValue<double>());
        Assert.Equal(6d, projectedImage["shadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(3d, projectedImage["shadow"]!["distance"]!.GetValue<double>());
        Assert.True(projectedImage["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        var directionLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "imageShadowDirectionDegrees")
            .ToArray();
        Assert.Single(directionLeaves);
        Assert.Equal(45d, directionLeaves[0]["value"]!.GetValue<double>());

        directionLeaves[0]["value"] = 90;
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
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var picture = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Picture>())
                .Single(item => item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var shadow = Assert.IsType<A.OuterShadow>(Assert.Single(picture.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.True(shadow.RotateWithShape!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal(76_200U, shadow.BlurRadius!.Value);
            Assert.Equal(38_100U, shadow.Distance!.Value);
            Assert.Equal(24_000, shadow.Descendants<A.Alpha>().Single().Val!.Value);
        }

        foreach (var path in ZipPartPaths(source).Where(path => !path.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, path), ZipBytes(editedBytes, path));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/image-shadow-direction.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(90d, reprojectedImage["shadow"]!["angle"]!.GetValue<double>());
        Assert.Equal(6d, reprojectedImage["shadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(3d, reprojectedImage["shadow"]!["distance"]!.GetValue<double>());
        Assert.True(reprojectedImage["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.Equal(90d, reprojectedImage["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "imageShadowDirectionDegrees")!["value"]!.GetValue<double>());
    }
}
