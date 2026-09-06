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
    public void PpjSourceBoundShapeShadowRotateWithShapeEditsLeafAndReprojects()
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
            ["opacity"] = 0.35,
            ["blur"] = 6,
            ["distance"] = 3,
            ["angle"] = 45,
            ["rotateWithShape"] = true,
        };
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
            ["shadow"] = new JsonObject
            {
                ["color"] = "#445566",
                ["opacity"] = 0.45,
                ["blur"] = 4,
                ["distance"] = 2,
                ["angle"] = 90,
                ["rotateWithShape"] = true,
            },
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
            var shapeShadow = Assert.IsType<A.OuterShadow>(Assert.Single(shape.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.True(shapeShadow.RotateWithShape!.Value);
            Assert.Equal(35_000, shapeShadow.Descendants<A.Alpha>().Single().Val!.Value);

            var line = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision line");
            var lineShadow = Assert.IsType<A.OuterShadow>(Assert.Single(line.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.True(lineShadow.RotateWithShape!.Value);
            Assert.Equal(45_000, lineShadow.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/source/shape-shadow-rotate-with-shape.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElements = projectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var projectedShape = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var projectedLine = projectedElements.Single(item => item["name"]!.GetValue<string>() == "decision line");
        Assert.True(projectedShape["style"]!["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.True(projectedLine["shadow"]!["rotateWithShape"]!.GetValue<bool>());

        var shapeRotateLeaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shadowRotateWithShape")
            .ToArray();
        var lineRotateLeaves = projectedLine["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shadowRotateWithShape")
            .ToArray();
        Assert.Single(shapeRotateLeaves);
        Assert.Single(lineRotateLeaves);
        Assert.True(shapeRotateLeaves[0]["value"]!.GetValue<bool>());
        Assert.True(lineRotateLeaves[0]["value"]!.GetValue<bool>());

        shapeRotateLeaves[0]["value"] = false;
        lineRotateLeaves[0]["value"] = false;
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
            var slideParts = package.PresentationPart!.SlideParts.ToArray();
            var shape = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeShadow = Assert.IsType<A.OuterShadow>(Assert.Single(shape.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.False(shapeShadow.RotateWithShape!.Value);
            Assert.Equal(35_000, shapeShadow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.Equal(76_200U, shapeShadow.BlurRadius!.Value);
            Assert.Equal(38_100U, shapeShadow.Distance!.Value);

            var line = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision line");
            var lineShadow = Assert.IsType<A.OuterShadow>(Assert.Single(line.ShapeProperties!.GetFirstChild<A.EffectList>()!.ChildElements));
            Assert.False(lineShadow.RotateWithShape!.Value);
            Assert.Equal(45_000, lineShadow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.Equal(50_800U, lineShadow.BlurRadius!.Value);
            Assert.Equal(25_400U, lineShadow.Distance!.Value);
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
                SourceUri = "deck.assets/edited/shape-shadow-rotate-with-shape.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElements = reprojectedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .ToArray();
        var reprojectedShape = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision rule");
        var reprojectedLine = reprojectedElements.Single(item => item["name"]!.GetValue<string>() == "decision line");
        Assert.False(reprojectedShape["style"]!["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.False(reprojectedLine["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.False(reprojectedShape["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowRotateWithShape")!["value"]!.GetValue<bool>());
        Assert.False(reprojectedLine["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowRotateWithShape")!["value"]!.GetValue<bool>());
    }
}
