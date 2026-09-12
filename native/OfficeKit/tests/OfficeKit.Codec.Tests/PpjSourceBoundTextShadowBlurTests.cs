using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundTextShadowBlurEditsDirectRunOwnerAndReprojects()
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
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["blur"] = 3,
                                    ["distance"] = 1.5,
                                    ["angle"] = 90,
                                    ["opacity"] = 0.72,
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
            var shadow = Assert.IsType<A.OuterShadow>(effects.ChildElements.Single());
            Assert.Equal(38_100, shadow.BlurRadius!.Value);
            Assert.Equal(19_050, shadow.Distance!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal("16324F", shadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(72_000, shadow.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/source/text-shadow-blur.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var shadowStyle = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!.AsObject();
        Assert.Equal("#16324F", shadowStyle["color"]!.GetValue<string>());
        Assert.Equal(3, shadowStyle["blur"]!.GetValue<double>());
        Assert.Equal(1.5, shadowStyle["distance"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, shadowStyle["angle"]!.GetValue<double>());
        Assert.Equal(0.72, shadowStyle["opacity"]!.GetValue<double>(), precision: 6);

        var leaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textShadowBlurRadiusEmu")
            .ToArray();
        var blurLeaf = Assert.Single(leaves);
        Assert.Equal(38_100, blurLeaf["value"]!.GetValue<long>());

        blurLeaf["value"] = 50_800;
        var editedProgramJson = projectedProgram.ToJsonString();
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(editedProgramJson),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
            var shadow = Assert.IsType<A.OuterShadow>(run.RunProperties!.GetFirstChild<A.EffectList>()!.ChildElements.Single());
            Assert.Equal(50_800, shadow.BlurRadius!.Value);
            Assert.Equal(19_050, shadow.Distance!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal("16324F", shadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(72_000, shadow.Descendants<A.Alpha>().Single().Val!.Value);
        }

        var editedBytes = edited.File.ToByteArray();
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
                SourceUri = "deck.assets/edited/text-shadow-blur.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedShadow = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!;
        Assert.Equal(4, reprojectedShadow["blur"]!.GetValue<double>());
        Assert.Equal(1.5, reprojectedShadow["distance"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, reprojectedShadow["angle"]!.GetValue<double>());
        Assert.Equal(0.72, reprojectedShadow["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(50_800, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textShadowBlurRadiusEmu")["value"]!.GetValue<long>());

        var invalidProgram = JsonNode.Parse(editedProgramJson)!.AsObject();
        var invalidLeaf = invalidProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "textShadowBlurRadiusEmu")!.AsObject();
        invalidLeaf["value"] = 12_700_001;
        var invalid = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(invalidProgram.ToJsonString()),
            },
        });
        Assert.False(invalid.Ok);
    }

    [Fact]
    public void PpjSourceBoundTextShadowBlurLeavesStayOpaqueForSiblingEffects()
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
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["blur"] = 3,
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
                SourceUri = "deck.assets/source/text-shadow-blur-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(element["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["shadow"]);
        Assert.DoesNotContain(
            element["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textShadowBlurRadiusEmu");

        foreach (var (label, mutate) in new[]
        {
            ("transform", (Action<A.OuterShadow>)(shadow => shadow.HorizontalRatio = 120_000)),
            ("missing-blur", (Action<A.OuterShadow>)(shadow => shadow.BlurRadius = null)),
            ("malformed-color", (Action<A.OuterShadow>)(shadow => shadow.GetFirstChild<A.RgbColorModelHex>()!.Val = "12345")),
        })
        {
            var unsupported = authored.File.ToByteArray();
            using (var stream = new MemoryStream())
            {
                stream.Write(unsupported);
                stream.Position = 0;
                using (var package = PresentationDocument.Open(stream, true))
                    mutate(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.OuterShadow>().Single());
                unsupported = stream.ToArray();
            }

            var unsupportedProjected = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ProjectPptxToPpj,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(RemoveEmbeddedPpj(unsupported)),
                PresentationProgram = new PresentationProgramRequest
                {
                    SourceUri = $"deck.assets/source/text-shadow-blur-{label}.pptx",
                },
            });
            Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
            var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
            var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
            Assert.DoesNotContain(
                unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
                leaf => leaf!["kind"]!.GetValue<string>() == "textShadowBlurRadiusEmu");
        }
    }
}
