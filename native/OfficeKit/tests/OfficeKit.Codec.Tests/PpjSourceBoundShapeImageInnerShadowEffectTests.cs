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
    public void PpjSourceBoundShapeImageInnerShadowEditsOwnersAndReprojects()
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
        shapeStyle["innerShadow"] = new JsonObject
        {
            ["color"] = "#AABBCC",
            ["blur"] = 5,
            ["distance"] = 2,
            ["angle"] = 45,
            ["opacity"] = 0.36,
        };

        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-inner-shadow",
                ["style"] = new JsonObject
                {
                    ["shadow"] = new JsonObject
                    {
                        ["color"] = "#17324D",
                        ["blur"] = 4,
                        ["distance"] = 2,
                        ["angle"] = 90,
                    },
                    ["innerShadow"] = new JsonObject
                    {
                        ["color"] = "#445566",
                        ["blur"] = 3,
                        ["distance"] = 1,
                        ["angle"] = 90,
                        ["opacity"] = 0.27,
                    },
                },
            },
        };
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["styleRef"] = "evidence-inner-shadow";

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
            var slideParts = package.PresentationPart!.SlideParts.ToArray();
            var shape = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeEffects = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.InnerShadow>(shapeEffects.ChildElements[0]);
            Assert.IsType<A.OuterShadow>(shapeEffects.ChildElements[1]);

            var picture = slideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.InnerShadow>(pictureEffects.ChildElements[0]);
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
                SourceUri = "deck.assets/source/shape-image-inner-shadow.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElements = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var projectedShape = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var projectedImage = projectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(5, projectedShape["style"]!["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.36, projectedShape["style"]!["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(3, projectedImage["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.27, projectedImage["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);

        var shapeLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("shapeInnerShadow", StringComparison.Ordinal))
            .ToArray();
        var imageLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("imageInnerShadow", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(5, shapeLeaves.Length);
        Assert.Equal(5, imageLeaves.Length);
        Assert.Equal(63_500, shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowBlurRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal(45, shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowDirectionDegrees")["value"]!.GetValue<double>());
        Assert.Equal(27_000, imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowOpacityThousandthPercent")["value"]!.GetValue<long>());

        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowBlurRadiusEmu")["value"] = 101_600;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowDistanceEmu")["value"] = 38_100;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowDirectionDegrees")["value"] = 90;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowColorRgb")["value"] = "#DDEEFF";
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeInnerShadowOpacityThousandthPercent")["value"] = 66_000;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowBlurRadiusEmu")["value"] = 76_200;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowDistanceEmu")["value"] = 25_400;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowDirectionDegrees")["value"] = 180;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowColorRgb")["value"] = "#112233";
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageInnerShadowOpacityThousandthPercent")["value"] = 55_000;

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
            var shape = package.PresentationPart!.SlideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeShadow = Assert.IsType<A.InnerShadow>(shape.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[0]);
            Assert.Equal(101_600, shapeShadow.BlurRadius!.Value);
            Assert.Equal(38_100, shapeShadow.Distance!.Value);
            Assert.Equal(5_400_000, shapeShadow.Direction!.Value);
            Assert.Equal("DDEEFF", shapeShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, shapeShadow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(shape.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements[1]);

            var picture = package.PresentationPart.SlideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureShadow = Assert.IsType<A.InnerShadow>(picture.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[0]);
            Assert.Equal(76_200, pictureShadow.BlurRadius!.Value);
            Assert.Equal(25_400, pictureShadow.Distance!.Value);
            Assert.Equal(10_800_000, pictureShadow.Direction!.Value);
            Assert.Equal("112233", pictureShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(55_000, pictureShadow.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/edited/shape-image-inner-shadow.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElements = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var reprojectedShape = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var reprojectedImage = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal("#DDEEFF", reprojectedShape["style"]!["innerShadow"]!["color"]!.GetValue<string>());
        Assert.Equal(8, reprojectedShape["style"]!["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(3, reprojectedShape["style"]!["innerShadow"]!["distance"]!.GetValue<double>());
        Assert.Equal(90, reprojectedShape["style"]!["innerShadow"]!["angle"]!.GetValue<double>());
        Assert.Equal(0.66, reprojectedShape["style"]!["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal("#112233", reprojectedImage["innerShadow"]!["color"]!.GetValue<string>());
        Assert.Equal(6, reprojectedImage["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(2, reprojectedImage["innerShadow"]!["distance"]!.GetValue<double>());
        Assert.Equal(180, reprojectedImage["innerShadow"]!["angle"]!.GetValue<double>());
        Assert.Equal(0.55, reprojectedImage["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundShapeImageInnerShadowLeavesStayOpaqueForComplexEffectLists()
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
        shapeStyle["innerShadow"] = new JsonObject { ["color"] = "#112233", ["blur"] = 2, ["distance"] = 1, ["angle"] = 30 };
        shapeStyle["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 12 };
        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-inner-shadow",
                ["style"] = new JsonObject
                {
                    ["innerShadow"] = new JsonObject { ["color"] = "#445566", ["blur"] = 2, ["distance"] = 1, ["angle"] = 30 },
                    ["glow"] = new JsonObject { ["color"] = "#667788", ["radius"] = 6 },
                },
            },
        };
        program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark")["styleRef"] = "evidence-inner-shadow";
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
                SourceUri = "deck.assets/source/shape-image-inner-shadow-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var leaves = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Where(item => item!["nativeRef"]?["leaves"] is not null)
            .SelectMany(item => item!["nativeRef"]!["leaves"]!.AsArray())
            .Select(leaf => leaf!["kind"]!.GetValue<string>())
            .ToArray();
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("shapeInnerShadow", StringComparison.Ordinal));
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("imageInnerShadow", StringComparison.Ordinal));
        var complexShape = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "decision rule")!.AsObject();
        Assert.Equal("shape", complexShape["type"]!.GetValue<string>());
        Assert.DoesNotContain(complexShape["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "shape.innerShadow"));
        var complexImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "evidence identity")!.AsObject();
        Assert.Equal("opaque", complexImage["type"]!.GetValue<string>());
    }
}
