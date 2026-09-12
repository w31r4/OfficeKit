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
    public void PpjSourceBoundTextShadowOpacityEditsDirectRunOwnerAndReprojects()
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
                                    ["opacity"] = 0.72,
                                    ["distance"] = 1.5,
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
                SourceUri = "deck.assets/source/text-shadow-opacity.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var shadowStyle = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!.AsObject();
        Assert.Equal("#16324F", shadowStyle["color"]!.GetValue<string>());
        Assert.Equal(0.72, shadowStyle["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(3, shadowStyle["blur"]!.GetValue<double>());
        Assert.Equal(1.5, shadowStyle["distance"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, shadowStyle["angle"]!.GetValue<double>());
        Assert.Equal(0.72, shadowStyle["opacity"]!.GetValue<double>(), precision: 6);

        var leaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent")
            .ToArray();
        var opacityLeaf = Assert.Single(leaves);
        Assert.Equal(72_000, opacityLeaf["value"]!.GetValue<long>());

        var themeSource = source.ToArray();
        using (var themeStream = new MemoryStream())
        {
            themeStream.Write(themeSource);
            themeStream.Position = 0;
            using (var themePackage = PresentationDocument.Open(themeStream, true))
            {
                var themeRun = themePackage.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
                var themeShadow = Assert.IsType<A.OuterShadow>(themeRun.RunProperties!.GetFirstChild<A.EffectList>()!.ChildElements.Single());
                themeShadow.GetFirstChild<A.RgbColorModelHex>()!.Remove();
                themeShadow.AppendChild(new A.SchemeColor(new A.Alpha { Val = 72_000 }) { Val = A.SchemeColorValues.Accent1 });
            }

            var themeProjected = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ProjectPptxToPpj,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(RemoveEmbeddedPpj(themeStream.ToArray())),
                PresentationProgram = new PresentationProgramRequest
                {
                    SourceUri = "deck.assets/source/text-shadow-opacity-theme.pptx",
                },
            });
            Assert.True(themeProjected.Ok, Diagnostics(themeProjected));
            var themeProgram = JsonNode.Parse(themeProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
            var themeElement = themeProgram["pages"]![0]!["elements"]![0]!.AsObject();
            Assert.Equal(0.72, themeElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!["opacity"]!.GetValue<double>(), precision: 6);
            Assert.Single(themeElement["nativeRef"]!["leaves"]!.AsArray()
                .Where(leaf => leaf!["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent"));
            Assert.DoesNotContain(themeElement["nativeRef"]!["leaves"]!.AsArray(),
                leaf => leaf!["kind"]!.GetValue<string>() == "textShadowColorRgb");
        }

        var staleSource = source.ToArray();
        using (var staleStream = new MemoryStream())
        {
            staleStream.Write(source);
            staleStream.Position = 0;
            using (var stalePackage = PresentationDocument.Open(staleStream, true))
            {
                var staleRun = stalePackage.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
                var staleShadow = Assert.IsType<A.OuterShadow>(staleRun.RunProperties!.GetFirstChild<A.EffectList>()!.ChildElements.Single());
                staleShadow.Descendants<A.Alpha>().Single().Val = 66_000;
            }

            staleSource = staleStream.ToArray();
        }
        var stale = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(staleSource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(projectedProgram.ToJsonString()),
            },
        });
        Assert.False(stale.Ok, Diagnostics(stale));

        opacityLeaf["value"] = 66_000;
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
            Assert.Equal(38_100, shadow.BlurRadius!.Value);
            Assert.Equal(19_050, shadow.Distance!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal("16324F", shadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, shadow.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/edited/text-shadow-opacity.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedShadow = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!;
        Assert.Equal("#16324F", reprojectedShadow["color"]!.GetValue<string>());
        Assert.Equal(3, reprojectedShadow["blur"]!.GetValue<double>());
        Assert.Equal(1.5, reprojectedShadow["distance"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, reprojectedShadow["angle"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.66, reprojectedShadow["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(66_000, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent")["value"]!.GetValue<long>());

        var invalidProgram = JsonNode.Parse(editedProgramJson)!.AsObject();
        var invalidLeaf = invalidProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent")!.AsObject();
        invalidLeaf["value"] = 100001;
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
    public void PpjSourceBoundTextShadowOpacityLeavesStayOpaqueForSiblingEffects()
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
                                    ["opacity"] = 0.72,
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
                SourceUri = "deck.assets/source/text-shadow-opacity-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(element["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["shadow"]);
        Assert.DoesNotContain(
            element["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent");

        foreach (var (label, mutate) in new[]
        {
            ("transform", (Action<A.OuterShadow>)(shadow => shadow.HorizontalRatio = 120_000)),
            ("overflow", (Action<A.OuterShadow>)(shadow => shadow.BlurRadius = 12_700_001)),
            ("missing-alpha", (Action<A.OuterShadow>)(shadow => shadow.Descendants<A.Alpha>().Single().Remove())),
            ("malformed-alpha", (Action<A.OuterShadow>)(shadow => shadow.Descendants<A.Alpha>().Single().Val = 100001)),
            ("missing-color", (Action<A.OuterShadow>)(shadow => shadow.GetFirstChild<A.RgbColorModelHex>()!.Remove())),
            ("theme-color", (Action<A.OuterShadow>)(shadow =>
            {
                var rgb = shadow.GetFirstChild<A.RgbColorModelHex>()!;
                rgb.Remove();
                shadow.AppendChild(new A.SchemeColor(new A.Alpha { Val = 72_000 }) { Val = A.SchemeColorValues.Accent1 });
            })),
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
                    SourceUri = $"deck.assets/source/text-shadow-opacity-{label}.pptx",
                },
            });
            Assert.True(unsupportedProjected.Ok, Diagnostics(unsupportedProjected));
            var unsupportedProgram = JsonNode.Parse(unsupportedProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
            var unsupportedElement = unsupportedProgram["pages"]![0]!["elements"]![0]!.AsObject();
            Assert.DoesNotContain(
                unsupportedElement["nativeRef"]!["leaves"]!.AsArray(),
                leaf => leaf!["kind"]!.GetValue<string>() == "textShadowOpacityThousandthPercent");
        }
    }
}
