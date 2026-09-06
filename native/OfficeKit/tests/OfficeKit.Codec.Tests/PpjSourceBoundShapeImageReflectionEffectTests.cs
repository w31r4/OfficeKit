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
    public void PpjSourceBoundShapeImageReflectionEditsOwnersAndReprojects()
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
        shapeStyle["reflection"] = new JsonObject
        {
            ["blur"] = 5,
            ["startOpacity"] = 0.42,
            ["endOpacity"] = 0.08,
            ["distance"] = 12,
            ["angle"] = 45,
        };

        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-reflection",
                ["style"] = new JsonObject
                {
                    ["shadow"] = new JsonObject
                    {
                        ["color"] = "#17324D",
                        ["blur"] = 4,
                        ["distance"] = 2,
                        ["angle"] = 90,
                    },
                    ["reflection"] = new JsonObject
                    {
                        ["blur"] = 3,
                        ["startOpacity"] = 0.27,
                        ["endOpacity"] = 0.05,
                        ["distance"] = 6,
                        ["angle"] = 90,
                    },
                },
            },
        };
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["styleRef"] = "evidence-reflection";

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
            Assert.IsType<A.OuterShadow>(shapeEffects.ChildElements[0]);
            var shapeReflection = Assert.IsType<A.Reflection>(shapeEffects.ChildElements[1]);
            Assert.Equal(63_500, shapeReflection.BlurRadius!.Value);
            Assert.Equal(42_000, shapeReflection.StartOpacity!.Value);
            Assert.Equal(8_000, shapeReflection.EndAlpha!.Value);
            Assert.Equal(152_400, shapeReflection.Distance!.Value);
            Assert.Equal(2_700_000, shapeReflection.Direction!.Value);
            Assert.Equal(0, shapeReflection.StartPosition!.Value);
            Assert.Equal(100_000, shapeReflection.EndPosition!.Value);

            var picture = slideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.OuterShadow>(pictureEffects.ChildElements[0]);
            var pictureReflection = Assert.IsType<A.Reflection>(pictureEffects.ChildElements[1]);
            Assert.Equal(38_100, pictureReflection.BlurRadius!.Value);
            Assert.Equal(27_000, pictureReflection.StartOpacity!.Value);
            Assert.Equal(5_000, pictureReflection.EndAlpha!.Value);
            Assert.Equal(76_200, pictureReflection.Distance!.Value);
            Assert.Equal(5_400_000, pictureReflection.Direction!.Value);
            Assert.Equal(0, pictureReflection.StartPosition!.Value);
            Assert.Equal(100_000, pictureReflection.EndPosition!.Value);
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
                SourceUri = "deck.assets/source/shape-image-reflection.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElements = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var projectedShape = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var projectedImage = projectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(5, projectedShape["style"]!["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.42, projectedShape["style"]!["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.08, projectedShape["style"]!["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(12, projectedShape["style"]!["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(45, projectedShape["style"]!["reflection"]!["angle"]!.GetValue<double>());
        Assert.Equal(3, projectedImage["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.27, projectedImage["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.05, projectedImage["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(6, projectedImage["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(90, projectedImage["reflection"]!["angle"]!.GetValue<double>());

        var shapeLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("shapeReflection", StringComparison.Ordinal))
            .ToArray();
        var imageLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("imageReflection", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(5, shapeLeaves.Length);
        Assert.Equal(5, imageLeaves.Length);
        Assert.Equal(63_500, shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionBlurRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal(42_000, shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionStartOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(45, shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionDirectionDegrees")["value"]!.GetValue<double>());
        Assert.Equal(27_000, imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionStartOpacityThousandthPercent")["value"]!.GetValue<long>());

        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionBlurRadiusEmu")["value"] = 101_600;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionStartOpacityThousandthPercent")["value"] = 64_000;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionEndOpacityThousandthPercent")["value"] = 12_000;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionDistanceEmu")["value"] = 228_600;
        shapeLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "shapeReflectionDirectionDegrees")["value"] = 90;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionBlurRadiusEmu")["value"] = 76_200;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionStartOpacityThousandthPercent")["value"] = 55_000;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionEndOpacityThousandthPercent")["value"] = 11_000;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionDistanceEmu")["value"] = 127_000;
        imageLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "imageReflectionDirectionDegrees")["value"] = 180;

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
            var shapeReflection = Assert.IsType<A.Reflection>(shape.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[1]);
            Assert.Equal(101_600, shapeReflection.BlurRadius!.Value);
            Assert.Equal(64_000, shapeReflection.StartOpacity!.Value);
            Assert.Equal(12_000, shapeReflection.EndAlpha!.Value);
            Assert.Equal(228_600, shapeReflection.Distance!.Value);
            Assert.Equal(5_400_000, shapeReflection.Direction!.Value);
            Assert.IsType<A.OuterShadow>(shape.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements[0]);

            var picture = package.PresentationPart.SlideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureReflection = Assert.IsType<A.Reflection>(picture.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[1]);
            Assert.Equal(76_200, pictureReflection.BlurRadius!.Value);
            Assert.Equal(55_000, pictureReflection.StartOpacity!.Value);
            Assert.Equal(11_000, pictureReflection.EndAlpha!.Value);
            Assert.Equal(127_000, pictureReflection.Distance!.Value);
            Assert.Equal(10_800_000, pictureReflection.Direction!.Value);
            Assert.IsType<A.OuterShadow>(picture.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements[0]);
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
                SourceUri = "deck.assets/edited/shape-image-reflection.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElements = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var reprojectedShape = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var reprojectedImage = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(8, reprojectedShape["style"]!["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.64, reprojectedShape["style"]!["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.12, reprojectedShape["style"]!["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(18, reprojectedShape["style"]!["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(90, reprojectedShape["style"]!["reflection"]!["angle"]!.GetValue<double>());
        Assert.Equal(6, reprojectedImage["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.55, reprojectedImage["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.11, reprojectedImage["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(10, reprojectedImage["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(180, reprojectedImage["reflection"]!["angle"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundShapeImageReflectionLeavesStayOpaqueForComplexEffectLists()
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
        shapeStyle["reflection"] = new JsonObject
        {
            ["blur"] = 2,
            ["startOpacity"] = 0.4,
            ["endOpacity"] = 0.1,
            ["distance"] = 1,
            ["angle"] = 30,
        };
        shapeStyle["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 12 };
        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-reflection",
                ["style"] = new JsonObject
                {
                    ["reflection"] = new JsonObject
                    {
                        ["blur"] = 2,
                        ["startOpacity"] = 0.4,
                        ["endOpacity"] = 0.1,
                        ["distance"] = 1,
                        ["angle"] = 30,
                    },
                    ["glow"] = new JsonObject { ["color"] = "#667788", ["radius"] = 6 },
                },
            },
        };
        program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark")["styleRef"] = "evidence-reflection";
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
                SourceUri = "deck.assets/source/shape-image-reflection-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var leaves = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Where(item => item!["nativeRef"]?["leaves"] is not null)
            .SelectMany(item => item!["nativeRef"]!["leaves"]!.AsArray())
            .Select(leaf => leaf!["kind"]!.GetValue<string>())
            .ToArray();
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("shapeReflection", StringComparison.Ordinal));
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("imageReflection", StringComparison.Ordinal));
        var complexShape = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "decision rule")!.AsObject();
        Assert.Equal("shape", complexShape["type"]!.GetValue<string>());
        Assert.DoesNotContain(complexShape["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "shape.reflection"));
        var complexImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "evidence identity")!.AsObject();
        Assert.Equal("opaque", complexImage["type"]!.GetValue<string>());
    }
}
