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
                                    ["angle"] = 45.5,
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
                    Assert.Equal(2_730_000, native.Direction!.Value);
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
        Assert.Equal(45.5, reflection["angle"]!.GetValue<double>());
        Assert.Equal(0, reflection["startPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(1, reflection["endPosition"]!.GetValue<double>(), precision: 6);

        var reflectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(7, reflectionLeaves.Length);
        Assert.Equal(0, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartPosition")["value"]!.GetValue<double>(), precision: 6);
        Assert.Equal(1, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndPosition")["value"]!.GetValue<double>(), precision: 6);
        Assert.Equal(63_500, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionBlurRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal(42_000, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(8_000, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndOpacityThousandthPercent")["value"]!.GetValue<long>());
        Assert.Equal(152_400, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDistanceEmu")["value"]!.GetValue<long>());
        Assert.Equal(45.5, reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"]!.GetValue<double>());

        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionBlurRadiusEmu")["value"] = 101_600;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartOpacityThousandthPercent")["value"] = 50_000;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndOpacityThousandthPercent")["value"] = 20_000;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDistanceEmu")["value"] = 38_100;
        reflectionLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"] = 90.25;

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
            Assert.Equal(5_415_000, native.Direction!.Value);
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
        Assert.Equal(90.25, reprojectedReflection["angle"]!.GetValue<double>());
        Assert.Equal(90.25, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionDirectionDegrees")["value"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundTextReflectionStartPositionEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Variable reflection start",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
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

        var variableStart = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(variableStart);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var nativeReflection = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
                nativeReflection.StartPosition = 20_000;
                nativeReflection.EndPosition = 100_000;
            }
            variableStart = stream.ToArray();
        }

        var source = RemoveEmbeddedPpj(variableStart);
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-reflection-start-position.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(0.2, reflection["startPosition"]!.GetValue<double>(), precision: 6);
        var reflectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal))
            .ToArray();
        var startLeaf = Assert.Single(reflectionLeaves);
        Assert.Equal("textReflectionStartPosition", startLeaf["kind"]!.GetValue<string>());
        Assert.Equal(0.2, startLeaf["value"]!.GetValue<double>(), precision: 6);

        startLeaf["value"] = 0.35;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(35_000, native.StartPosition!.Value);
            Assert.Equal(100_000, native.EndPosition!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-start-position.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(0.35, reprojectedReflection["startPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(1, reprojectedReflection["endPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.35, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionStartPosition")["value"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundTextReflectionEndPositionEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Variable reflection end",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
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

        var variableEnd = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(variableEnd);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
            {
                var nativeReflection = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
                nativeReflection.StartPosition = 0;
                nativeReflection.EndPosition = 80_000;
            }
            variableEnd = stream.ToArray();
        }

        var source = RemoveEmbeddedPpj(variableEnd);
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-reflection-end-position.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(0.8, reflection["endPosition"]!.GetValue<double>(), precision: 6);
        var reflectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textReflection", StringComparison.Ordinal))
            .ToArray();
        var endLeaf = Assert.Single(reflectionLeaves);
        Assert.Equal("textReflectionEndPosition", endLeaf["kind"]!.GetValue<string>());
        Assert.Equal(0.8, endLeaf["value"]!.GetValue<double>(), precision: 6);

        endLeaf["value"] = 0.65;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(0, native.StartPosition!.Value);
            Assert.Equal(65_000, native.EndPosition!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-end-position.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(0.65, reprojectedReflection["endPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.65, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionEndPosition")["value"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundTextReflectionFadeAngleEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Variable reflection fade",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["fadeAngle"] = 45.5,
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(2_730_000, native.FadeDirection!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-fade-angle.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(45.5, reflection["fadeAngle"]!.GetValue<double>(), precision: 6);
        var fadeLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionFadeAngleDegrees");
        Assert.Equal(45.5, fadeLeaf["value"]!.GetValue<double>(), precision: 6);

        fadeLeaf["value"] = 90.25;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(5_415_000, native.FadeDirection!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-fade-angle.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(90.25, reprojectedReflection["fadeAngle"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90.25, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionFadeAngleDegrees")["value"]!.GetValue<double>(), precision: 6);

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().HorizontalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-fade-angle-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionFadeAngleDegrees");
    }


    [Fact]
    public void PpjSourceBoundTextReflectionScaleXEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Variable reflection scale",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["scaleX"] = 1.25,
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(125_000, native.HorizontalRatio!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-scale-x.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(1.25, reflection["scaleX"]!.GetValue<double>(), precision: 6);
        var scaleLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionScaleX");
        Assert.Equal(1.25, scaleLeaf["value"]!.GetValue<double>(), precision: 6);

        scaleLeaf["value"] = 0.75;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(75_000, native.HorizontalRatio!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-scale-x.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(0.75, reprojectedReflection["scaleX"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.75, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionScaleX")["value"]!.GetValue<double>(), precision: 6);

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().HorizontalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-scale-x-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionScaleX");
    }

    [Fact]
    public void PpjSourceBoundTextReflectionScaleYEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Vertical reflection scale",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["scaleY"] = 1.25,
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(125_000, native.VerticalRatio!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-scale-y.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(1.25, reflection["scaleY"]!.GetValue<double>(), precision: 6);
        var scaleLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionScaleY");
        Assert.Equal(1.25, scaleLeaf["value"]!.GetValue<double>(), precision: 6);

        scaleLeaf["value"] = -0.75;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(-75_000, native.VerticalRatio!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-scale-y.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(-0.75, reprojectedReflection["scaleY"]!.GetValue<double>(), precision: 6);
        Assert.Equal(-0.75, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionScaleY")["value"]!.GetValue<double>(), precision: 6);

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().HorizontalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-scale-y-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionScaleY");
    }

    [Fact]
    public void PpjSourceBoundTextReflectionSkewXEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Horizontal reflection skew",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["skewX"] = -12.5,
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(-750_000, native.HorizontalSkew!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-skew-x.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(-12.5, reflection["skewX"]!.GetValue<double>(), precision: 6);
        var skewLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionSkewX");
        Assert.Equal(-12.5, skewLeaf["value"]!.GetValue<double>(), precision: 6);

        skewLeaf["value"] = 20;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(1_200_000, native.HorizontalSkew!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-skew-x.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(20, reprojectedReflection["skewX"]!.GetValue<double>(), precision: 6);
        Assert.Equal(20, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionSkewX")["value"]!.GetValue<double>(), precision: 6);

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().VerticalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-skew-x-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionSkewX");
    }

    [Fact]
    public void PpjSourceBoundTextReflectionSkewYEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Vertical reflection skew",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["skewY"] = -12.5,
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(-750_000, native.VerticalSkew!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-skew-y.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal(-12.5, reflection["skewY"]!.GetValue<double>(), precision: 6);
        var skewLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionSkewY");
        Assert.Equal(-12.5, skewLeaf["value"]!.GetValue<double>(), precision: 6);

        skewLeaf["value"] = 20;
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(1_200_000, native.VerticalSkew!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-skew-y.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal(20, reprojectedReflection["skewY"]!.GetValue<double>(), precision: 6);
        Assert.Equal(20, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionSkewY")["value"]!.GetValue<double>(), precision: 6);

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().HorizontalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-skew-y-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionSkewY");
    }

    [Fact]
    public void PpjSourceBoundTextReflectionAlignmentEditsCanonicalTokenAndReprojects()
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
                            ["text"] = "Reflection alignment",
                            ["style"] = new JsonObject
                            {
                                ["reflection"] = new JsonObject
                                {
                                    ["blur"] = 5,
                                    ["startOpacity"] = 0.42,
                                    ["endOpacity"] = 0.08,
                                    ["distance"] = 12,
                                    ["angle"] = 45.5,
                                    ["alignment"] = "b",
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(A.RectangleAlignmentValues.Bottom, native.Alignment!.Value);
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
                SourceUri = "deck.assets/source/text-reflection-alignment.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflection = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!.AsObject();
        Assert.Equal("b", reflection["alignment"]!.GetValue<string>());
        var alignmentLeaf = Assert.Single(
            projectedElement["nativeRef"]!["leaves"]!.AsArray()
                .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "textReflectionAlignment");
        Assert.Equal("b", alignmentLeaf["value"]!.GetValue<string>());

        alignmentLeaf["value"] = "tr";
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
            var native = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single();
            Assert.Equal(A.RectangleAlignmentValues.TopRight, native.Alignment!.Value);
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
                SourceUri = "deck.assets/edited/text-reflection-alignment.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedReflection = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["reflection"]!;
        Assert.Equal("tr", reprojectedReflection["alignment"]!.GetValue<string>());
        Assert.Equal("tr", reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textReflectionAlignment")["value"]!.GetValue<string>());

        var unsupported = authored.File.ToByteArray();
        using (var stream = new MemoryStream())
        {
            stream.Write(unsupported);
            stream.Position = 0;
            using (var package = PresentationDocument.Open(stream, true))
                package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Reflection>().Single().VerticalSkew = 60_000;
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
                SourceUri = "deck.assets/source/text-reflection-alignment-skew.pptx",
            },
        });
        Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
        var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(unsupportedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["reflection"]);
        Assert.DoesNotContain(
            unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textReflectionAlignment");
    }

    [Fact]
    public void PpjSourceBoundTextReflectionLeavesStayOpaqueOutsideOneVariableEndpointProfile()
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
                                    ["angle"] = 45.5,
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
                reflection.EndPosition = 99_000;
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
