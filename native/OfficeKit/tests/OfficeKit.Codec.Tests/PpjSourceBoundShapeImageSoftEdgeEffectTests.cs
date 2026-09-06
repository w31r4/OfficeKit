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
    public void PpjSourceBoundShapeImageSoftEdgeEditsOwnersAndReprojects()
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
        shapeStyle["softEdge"] = new JsonObject { ["radius"] = 8 };

        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-soft-edge",
                ["style"] = new JsonObject
                {
                    ["shadow"] = new JsonObject
                    {
                        ["color"] = "#17324D",
                        ["blur"] = 4,
                        ["distance"] = 2,
                        ["angle"] = 90,
                    },
                    ["softEdge"] = new JsonObject { ["radius"] = 4 },
                },
            },
        };
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["styleRef"] = "evidence-soft-edge";
        program["pages"]!.AsArray()[0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "decision-line",
            ["name"] = "decision line",
            ["type"] = "line",
            ["frame"] = new JsonObject { ["x"] = 80, ["y"] = 500, ["width"] = 220, ["height"] = 20 },
            ["viewBox"] = new JsonArray(220, 20),
            ["points"] = "0,10 220,10",
            ["curve"] = "sharp",
            ["stroke"] = new JsonObject { ["color"] = "#17324D", ["width"] = 1 },
            ["softEdge"] = new JsonObject { ["radius"] = 2 },
        });

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
            var shapeSoftEdge = Assert.IsType<A.SoftEdge>(shapeEffects.ChildElements[1]);
            Assert.Equal(101_600U, shapeSoftEdge.Radius!.Value);

            var picture = slideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.OuterShadow>(pictureEffects.ChildElements[0]);
            var pictureSoftEdge = Assert.IsType<A.SoftEdge>(pictureEffects.ChildElements[1]);
            Assert.Equal(50_800U, pictureSoftEdge.Radius!.Value);

            var line = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision line");
            var lineEffects = line.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            var lineSoftEdge = Assert.IsType<A.SoftEdge>(lineEffects.ChildElements[0]);
            Assert.Equal(25_400U, lineSoftEdge.Radius!.Value);
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
                SourceUri = "deck.assets/source/shape-image-soft-edge.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElements = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var projectedShape = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var projectedImage = projectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var projectedLine = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision line");
        Assert.Equal(8, projectedShape["style"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(4, projectedImage["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(2, projectedLine["softEdge"]!["radius"]!.GetValue<double>());

        var shapeLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("shapeSoftEdge", StringComparison.Ordinal))
            .ToArray();
        var imageLeaves = projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("imageSoftEdge", StringComparison.Ordinal))
            .ToArray();
        var lineLeaves = projectedLine["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shapeSoftEdgeRadiusEmu")
            .ToArray();
        Assert.Single(shapeLeaves);
        Assert.Single(imageLeaves);
        Assert.Single(lineLeaves);
        Assert.Equal(101_600, shapeLeaves[0]["value"]!.GetValue<long>());
        Assert.Equal(50_800, imageLeaves[0]["value"]!.GetValue<long>());
        Assert.Equal(25_400, lineLeaves[0]["value"]!.GetValue<long>());

        shapeLeaves[0]["value"] = 127_000;
        imageLeaves[0]["value"] = 76_200;
        lineLeaves[0]["value"] = 38_100;
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
            var shapeEffects = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.OuterShadow>(shapeEffects.ChildElements[0]);
            Assert.Equal(127_000U, Assert.IsType<A.SoftEdge>(shapeEffects.ChildElements[1]).Radius!.Value);

            var picture = package.PresentationPart.SlideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.OuterShadow>(pictureEffects.ChildElements[0]);
            Assert.Equal(76_200U, Assert.IsType<A.SoftEdge>(pictureEffects.ChildElements[1]).Radius!.Value);

            var line = package.PresentationPart.SlideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision line");
            Assert.Equal(38_100U, Assert.IsType<A.SoftEdge>(line.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements[0]).Radius!.Value);
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
                SourceUri = "deck.assets/edited/shape-image-soft-edge.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElements = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var reprojectedShape = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var reprojectedImage = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var reprojectedLine = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision line");
        Assert.Equal(10, reprojectedShape["style"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(6, reprojectedImage["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(3, reprojectedLine["softEdge"]!["radius"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundShapeImageSoftEdgeLeavesStayOpaqueForComplexEffectLists()
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
        shapeStyle["softEdge"] = new JsonObject { ["radius"] = 8 };
        shapeStyle["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 12 };
        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-soft-edge",
                ["style"] = new JsonObject
                {
                    ["softEdge"] = new JsonObject { ["radius"] = 4 },
                    ["glow"] = new JsonObject { ["color"] = "#445566", ["radius"] = 6 },
                },
            },
        };
        program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark")["styleRef"] = "evidence-soft-edge";

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
                SourceUri = "deck.assets/source/shape-image-soft-edge-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var leaves = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Where(item => item!["nativeRef"]?["leaves"] is not null)
            .SelectMany(item => item!["nativeRef"]!["leaves"]!.AsArray())
            .Select(leaf => leaf!["kind"]!.GetValue<string>())
            .ToArray();
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("shapeSoftEdge", StringComparison.Ordinal));
        Assert.DoesNotContain(leaves, kind => kind.StartsWith("imageSoftEdge", StringComparison.Ordinal));
        var complexShape = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "decision rule")!.AsObject();
        Assert.Equal("shape", complexShape["type"]!.GetValue<string>());
        Assert.DoesNotContain(complexShape["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "shape.softEdge"));
        var complexImage = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Single(item => item!["name"]!.GetValue<string>() == "evidence identity")!.AsObject();
        Assert.Equal("opaque", complexImage["type"]!.GetValue<string>());
    }
}
