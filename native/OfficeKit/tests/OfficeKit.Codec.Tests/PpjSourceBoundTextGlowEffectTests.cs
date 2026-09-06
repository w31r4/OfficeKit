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
    public void PpjSourceBoundTextGlowEditsDirectRunOwnerAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "examples",
            "ppj",
            "minimum.ppj")))!.AsObject();
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Evidence changed the decision",
                            ["style"] = new JsonObject
                            {
                                ["glow"] = new JsonObject
                                {
                                    ["color"] = "#D9A514",
                                    ["radius"] = 8,
                                    ["opacity"] = 0.42,
                                },
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 90,
                                },
                            },
                        },
                    },
                },
            },
        };

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
            var effects = run.RunProperties!.GetFirstChild<A.EffectList>()!;
            Assert.Collection(
                effects.ChildElements,
                glow =>
                {
                    var native = Assert.IsType<A.Glow>(glow);
                    Assert.Equal(101_600U, native.Radius!.Value);
                    Assert.Equal("D9A514", native.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                    Assert.Equal(42_000, native.Descendants<A.Alpha>().Single().Val!.Value);
                },
                shadow => Assert.IsType<A.OuterShadow>(shadow));
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
                SourceUri = "deck.assets/source/text-glow.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var glow = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["glow"]!.AsObject();
        Assert.Equal("#D9A514", glow["color"]!.GetValue<string>());
        Assert.Equal(8, glow["radius"]!.GetValue<double>());
        Assert.Equal(0.42, glow["opacity"]!.GetValue<double>(), precision: 6);

        var glowLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textGlow", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(3, glowLeaves.Length);
        Assert.Equal(101_600, glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal("#d9a514", glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowColorRgb")["value"]!.GetValue<string>());
        Assert.Equal(42_000, glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowOpacityThousandthPercent")["value"]!.GetValue<long>());

        glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowRadiusEmu")["value"] = 127_000;
        glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowColorRgb")["value"] = "#AABBCC";
        glowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowOpacityThousandthPercent")["value"] = 66_000;

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

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
            var effects = run.RunProperties!.GetFirstChild<A.EffectList>()!;
            Assert.Equal(127_000U, effects.GetFirstChild<A.Glow>()!.Radius!.Value);
            Assert.Equal("AABBCC", effects.GetFirstChild<A.Glow>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, effects.GetFirstChild<A.Glow>()!.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(effects.ChildElements[1]);
        }

        foreach (var path in ZipPartPaths(source).Where(path => !path.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, path), ZipBytes(edited.File.ToByteArray(), path));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/text-glow.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedGlow = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["glow"]!;
        Assert.Equal("#AABBCC", reprojectedGlow["color"]!.GetValue<string>());
        Assert.Equal(10, reprojectedGlow["radius"]!.GetValue<double>());
        Assert.Equal(0.66, reprojectedGlow["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(127_000, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textGlowRadiusEmu")["value"]!.GetValue<long>());
    }

    [Fact]
    public void PpjSourceBoundTextGlowLeavesStayOpaqueForComplexEffectList()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "examples",
            "ppj",
            "minimum.ppj")))!.AsObject();
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Evidence changed the decision",
                            ["style"] = new JsonObject
                            {
                                ["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 8 },
                                ["innerShadow"] = new JsonObject
                                {
                                    ["color"] = "#112233",
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 30,
                                },
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 90,
                                },
                            },
                        },
                    },
                },
            },
        };
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
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
                SourceUri = "deck.assets/source/text-glow-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(element["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["glow"]);
        Assert.DoesNotContain(
            element["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>().StartsWith("textGlow", StringComparison.Ordinal));
    }
}
