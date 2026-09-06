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
    public void PpjSourceBoundShapeImageGlowEditsOwnersAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();

        var shapeStyle = program["design"]!["styles"]!["shape"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "decision-band")!["style"]!.AsObject();
        shapeStyle["shadow"] = new JsonObject
        {
            ["color"] = "#17324D",
            ["opacity"] = 0.24,
            ["blur"] = 6,
            ["distance"] = 3,
            ["angle"] = 45,
        };
        shapeStyle["glow"] = new JsonObject
        {
            ["color"] = "#D9A514",
            ["radius"] = 12,
            ["opacity"] = 0.42,
        };

        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-glow",
                ["style"] = new JsonObject
                {
                    ["glow"] = new JsonObject
                    {
                        ["color"] = "#445566",
                        ["radius"] = 6,
                        ["opacity"] = 0.27,
                    },
                },
            },
        };
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["styleRef"] = "evidence-glow";

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
            var shape = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Shape>())
                .Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeEffects = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.Glow>(shapeEffects.ChildElements[0]);
            Assert.IsType<A.OuterShadow>(shapeEffects.ChildElements[1]);

            var picture = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Picture>())
                .Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.Glow>(pictureEffects.ChildElements[0]);
            Assert.IsType<A.OuterShadow>(pictureEffects.ChildElements[1]);
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
                SourceUri = "deck.assets/source/shape-image-glow.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElements = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var projectedShape = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var projectedImage = projectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(12, projectedShape["style"]!["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal(0.42, projectedShape["style"]!["glow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.True(projectedImage["glow"] is not null, projectedImage.ToJsonString());
        Assert.Equal(6, projectedImage["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal(0.27, projectedImage["glow"]!["opacity"]!.GetValue<double>(), precision: 6);

        var shapeGlowLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("shapeGlow", StringComparison.Ordinal))
            .ToArray();
        var imageGlowLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("imageGlow", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(3, shapeGlowLeaves.Length);
        Assert.Equal(3, imageGlowLeaves.Length);
        Assert.Equal(152_400, shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal("#d9a514", shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowColorRgb")["value"]!.GetValue<string>());
        Assert.Equal(42_000, shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(76_200, imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal("#445566", imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowColorRgb")["value"]!.GetValue<string>());
        Assert.Equal(27_000, imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowOpacityThousandthPercent")["value"]!.GetValue<long>());

        shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowRadiusEmu")["value"] = 177_800;
        shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowColorRgb")["value"] = "#AABBCC";
        shapeGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeGlowOpacityThousandthPercent")["value"] = 66_000;
        imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowRadiusEmu")["value"] = 101_600;
        imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowColorRgb")["value"] = "#112233";
        imageGlowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageGlowOpacityThousandthPercent")["value"] = 55_000;

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
            var shape = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Shape>())
                .Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeGlow = Assert.IsType<A.Glow>(shape.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[0]);
            Assert.Equal(177_800U, shapeGlow.Radius!.Value);
            Assert.Equal("AABBCC", shapeGlow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, shapeGlow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(shape.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements[1]);

            var picture = package.PresentationPart!.SlideParts
                .SelectMany(part => part.Slide!.Descendants<P.Picture>())
                .Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureGlow = Assert.IsType<A.Glow>(picture.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[0]);
            Assert.Equal(101_600U, pictureGlow.Radius!.Value);
            Assert.Equal("112233", pictureGlow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(55_000, pictureGlow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(picture.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements[1]);
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
                SourceUri = "deck.assets/edited/shape-image-glow.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElements = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var reprojectedShape = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var reprojectedImage = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal("#AABBCC", reprojectedShape["style"]!["glow"]!["color"]!.GetValue<string>());
        Assert.Equal(14, reprojectedShape["style"]!["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal(0.66, reprojectedShape["style"]!["glow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal("#112233", reprojectedImage["glow"]!["color"]!.GetValue<string>());
        Assert.Equal(8, reprojectedImage["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal(0.55, reprojectedImage["glow"]!["opacity"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundShapeImageGlowLeavesStayOpaqueForComplexEffectLists()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var shapeStyle = program["design"]!["styles"]!["shape"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "decision-band")!["style"]!.AsObject();
        shapeStyle["shadow"] = new JsonObject
        {
            ["color"] = "#17324D",
            ["blur"] = 6,
            ["distance"] = 3,
            ["angle"] = 45,
        };
        shapeStyle["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 12 };
        shapeStyle["innerShadow"] = new JsonObject
        {
            ["color"] = "#112233",
            ["blur"] = 2,
            ["distance"] = 1,
            ["angle"] = 30,
        };
        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-glow",
                ["style"] = new JsonObject
                {
                    ["glow"] = new JsonObject { ["color"] = "#445566", ["radius"] = 6 },
                    ["innerShadow"] = new JsonObject
                    {
                        ["color"] = "#112233",
                        ["blur"] = 2,
                        ["distance"] = 1,
                        ["angle"] = 30,
                    },
                },
            },
        };
        program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark")["styleRef"] = "evidence-glow";

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
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/shape-image-glow-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var leaves = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Where(item => item!["nativeRef"]?["leaves"] is not null)
            .SelectMany(item => item!["nativeRef"]!["leaves"]!.AsArray())
            .Select(leaf => leaf!["kind"]!.GetValue<string>())
            .ToArray();
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("shapeGlow", StringComparison.Ordinal));
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("imageGlow", StringComparison.Ordinal));
        var complexShape = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "decision rule")!.AsObject();
        Assert.Equal("shape", complexShape["type"]!.GetValue<string>());
        Assert.DoesNotContain(complexShape["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "shape.glow"));
        var complexImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "evidence identity")!.AsObject();
        Assert.Equal("opaque", complexImage["type"]!.GetValue<string>());
    }
}
