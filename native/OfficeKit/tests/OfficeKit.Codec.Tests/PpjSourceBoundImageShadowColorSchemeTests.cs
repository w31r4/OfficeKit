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
    public void PpjSourceBoundImageShadowColorSchemeEditsLeafAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        program["design"]!["theme"]!["colors"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "accent1",
            ["value"] = "#4F81BD",
        });
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["shadow"]!.AsObject()["color"] = new JsonObject { ["token"] = "accent1" };
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
            var color = Assert.IsType<A.SchemeColor>(shadow.FirstChild);
            Assert.Equal(A.SchemeColorValues.Accent1, color.Val!.Value);
            Assert.True(shadow.RotateWithShape!.Value);
            Assert.Equal(76_200U, shadow.BlurRadius!.Value);
            Assert.Equal(38_100U, shadow.Distance!.Value);
            Assert.Equal(24_000, color.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/source/image-shadow-color-scheme.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal("accent1", projectedImage["shadow"]!["color"]!["token"]!.GetValue<string>());
        Assert.True(projectedImage["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        var colorLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "imageShadowColorScheme")
            .ToArray();
        Assert.Single(colorLeaves);
        Assert.Equal("accent1", colorLeaves[0]["value"]!.GetValue<string>());

        colorLeaves[0]["value"] = "accent2";
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
            var color = Assert.IsType<A.SchemeColor>(shadow.FirstChild);
            Assert.Equal(A.SchemeColorValues.Accent2, color.Val!.Value);
            Assert.True(shadow.RotateWithShape!.Value);
            Assert.Equal(76_200U, shadow.BlurRadius!.Value);
            Assert.Equal(38_100U, shadow.Distance!.Value);
            Assert.Equal(24_000, color.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/edited/image-shadow-color-scheme.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal("accent2", reprojectedImage["shadow"]!["color"]!["token"]!.GetValue<string>());
        Assert.True(reprojectedImage["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.Equal("accent2", reprojectedImage["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "imageShadowColorScheme")["value"]!.GetValue<string>());
    }
}
