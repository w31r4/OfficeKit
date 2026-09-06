using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundTextReflectionEditsDirectRunOwnerAndReprojects()
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
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45,
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
                shadow => Assert.IsType<A.OuterShadow>(shadow),
                reflection =>
                {
                    var native = Assert.IsType<A.Reflection>(reflection);
                    Assert.Equal(63_500, native.BlurRadius!.Value);
                    Assert.Equal(42_000, native.StartOpacity!.Value);
                    Assert.Equal(0, native.StartPosition!.Value);
                    Assert.Equal(8_000, native.EndAlpha!.Value);
                    Assert.Equal(100_000, native.EndPosition!.Value);
                    Assert.Equal(152_400, native.Distance!.Value);
                    Assert.Equal(2_700_000, native.Direction!.Value);
                });
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
                SourceUri = "deck.assets/source/text-reflection.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(5, reflection["blur"]!.GetValue<double>());
        Assert.Equal(0.42, reflection["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.08, reflection["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(12, reflection["distance"]!.GetValue<double>());
        Assert.Equal(45, reflection["angle"]!.GetValue<double>());

        var reflectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(5, reflectionLeaves.Length);
        Assert.Equal(63_500, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionBlurRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal(42_000, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(8_000, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(152_400, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDistanceEmu")["value"]!.GetValue<long>());
        Assert.Equal(45, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"]!.GetValue<double>());

        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionBlurRadiusEmu")["value"] = 101_600;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartOpacityThousandthPercent")["value"] = 50_000;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndOpacityThousandthPercent")["value"] = 20_000;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDistanceEmu")["value"] = 38_100;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"] = 90;

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
            Assert.IsType<A.OuterShadow>(effects.ChildElements[0]);
            var native = Assert.IsType<A.Reflection>(effects.ChildElements[1]);
            Assert.Equal(101_600, native.BlurRadius!.Value);
            Assert.Equal(50_000, native.StartOpacity!.Value);
            Assert.Equal(20_000, native.EndAlpha!.Value);
            Assert.Equal(0, native.StartPosition!.Value);
            Assert.Equal(100_000, native.EndPosition!.Value);
            Assert.Equal(38_100, native.Distance!.Value);
            Assert.Equal(5_400_000, native.Direction!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(8, reprojectedReflection["blur"]!.GetValue<double>());
        Assert.Equal(0.5, reprojectedReflection["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.2, reprojectedReflection["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(3, reprojectedReflection["distance"]!.GetValue<double>());
        Assert.Equal(90, reprojectedReflection["angle"]!.GetValue<double>());
        Assert.Equal(90, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundTextReflectionLeavesStayOpaqueOutsideFullSpanProfile()
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
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45,
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

        var complex = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-reflection-complex.pptx",
            },
        });
        Assert.True(complex.Ok, Diagnostics(complex));
        var complexProgram = JsonNode.Parse(complex.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var complexElement = complexProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(complexElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            complexElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal));

        var nonFullSpan = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(nonFullSpan);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var reflection = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
                reflection.StartPosition = 1_000;
            }
            nonFullSpan = stream.ToArray();
        }

        var nonFullProjected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(nonFullSpan)),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-reflection-non-full-span.pptx",
            },
        });
        Assert.True(nonFullProjected.Ok, Diagnostics(nonFullProjected));
        var nonFullProgram = JsonNode.Parse(nonFullProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var nonFullElement = nonFullProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(nonFullElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            nonFullElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal));
    }
}
