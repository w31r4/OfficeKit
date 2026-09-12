using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundTextShadowScaleYEditsDirectRunOwnerAndReprojects()
    {
        var program = ShadowYProgram(scaleY: -0.5);
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
            var shadow = Assert.IsType<A.OuterShadow>(run.RunProperties!.GetFirstChild<A.EffectList>()!.ChildElements.Single());
            Assert.Equal(38_100, shadow.BlurRadius!.Value);
            Assert.Equal(19_050, shadow.Distance!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal(-50_000, shadow.VerticalRatio!.Value);
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
                SourceUri = "deck.assets/source/text-shadow-scale-y.pptx",
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
        Assert.Equal(-0.5, shadowStyle["scaleY"]!.GetValue<double>(), precision: 6);
        var scaleLeaf = Assert.Single(projectedElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textShadowScaleY");
        Assert.Equal(-0.5, scaleLeaf!["value"]!.GetValue<double>(), precision: 6);

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
                    SourceUri = "deck.assets/source/text-shadow-scale-y-theme.pptx",
                },
            });
            Assert.True(themeProjected.Ok, Diagnostics(themeProjected));
            var themeProgram = JsonNode.Parse(themeProjected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
            var themeElement = themeProgram["pages"]![0]!["elements"]![0]!.AsObject();
            Assert.Equal(-0.5, themeElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!["scaleY"]!.GetValue<double>(), precision: 6);
            Assert.Single(themeElement["nativeRef"]!["leaves"]!.AsArray(),
                leaf => leaf!["kind"]!.GetValue<string>() == "textShadowScaleY");
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
                staleShadow.VerticalRatio = 125_000;
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

        scaleLeaf!["value"] = 1.25;
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
            var shadow = Assert.IsType<A.OuterShadow>(run.RunProperties!.GetFirstChild<A.EffectList>()!.ChildElements.Single());
            Assert.Equal(38_100, shadow.BlurRadius!.Value);
            Assert.Equal(19_050, shadow.Distance!.Value);
            Assert.Equal(5_400_000, shadow.Direction!.Value);
            Assert.Equal(125_000, shadow.VerticalRatio!.Value);
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
                SourceUri = "deck.assets/edited/text-shadow-scale-y.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedShadow = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["shadow"]!;
        Assert.Equal(1.25, reprojectedShadow["scaleY"]!.GetValue<double>(), precision: 6);
        Assert.Equal(1.25, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textShadowScaleY")["value"]!.GetValue<double>(), precision: 6);

        var invalidProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        invalidProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "textShadowScaleY")!["value"] = 2_147_483_648;
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
        Assert.False(invalid.Ok, Diagnostics(invalid));

        var noOp = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.ToJsonString()),
            },
        });
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);
    }

    [Fact]
    public void PpjSourceBoundTextShadowScaleYLeavesStayOpaqueForUnsupportedGraphs()
    {
        var authored = CompileScaleYShadow(ShadowYProgram(scaleY: -0.5));
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());

        foreach (var (label, mutate) in new[]
        {
            ("transform", (Action<A.OuterShadow>)(shadow => shadow.HorizontalRatio = 120_000)),
            ("missing-scale", (Action<A.OuterShadow>)(shadow => shadow.VerticalRatio = null)),
            ("invalid-scale", (Action<A.OuterShadow>)(shadow => shadow.SetAttribute(new OpenXmlAttribute("sy", "", "invalid")))),
            ("missing-geometry", (Action<A.OuterShadow>)(shadow => { shadow.BlurRadius = null; shadow.Distance = null; shadow.Direction = null; })),
            ("missing-color", (Action<A.OuterShadow>)(shadow => shadow.GetFirstChild<A.RgbColorModelHex>()!.Remove())),
            ("unknown-descendant", (Action<A.OuterShadow>)(shadow => shadow.AppendChild(new OpenXmlUnknownElement("future", "unknown", "urn:officekit:test")))),
        })
        {
            var unsupported = source.ToArray();
            using (var stream = new MemoryStream())
            {
                stream.Write(unsupported);
                stream.Position = 0;
                using (var package = PresentationDocument.Open(stream, true))
                    mutate(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.OuterShadow>().Single());
                unsupported = stream.ToArray();
            }

            var projected = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ProjectPptxToPpj,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(RemoveEmbeddedPpj(unsupported)),
                PresentationProgram = new PresentationProgramRequest
                {
                    SourceUri = $"deck.assets/source/text-shadow-scale-y-{label}.pptx",
                },
            });
            if (label == "invalid-scale")
            {
                Assert.False(projected.Ok, Diagnostics(projected));
                continue;
            }
            Assert.True(projected.Ok, label + ": " + Diagnostics(projected));
            var element = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!["pages"]![0]!["elements"]![0]!;
            Assert.DoesNotContain(element!["nativeRef"]!["leaves"]!.AsArray(),
                leaf => leaf!["kind"]!.GetValue<string>() == "textShadowScaleY");
        }

        var siblingAuthored = CompileScaleYShadow(ShadowYProgram(scaleY: -0.5, siblingGlow: true));
        Assert.True(siblingAuthored.Ok, Diagnostics(siblingAuthored));
        var siblingProjected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(siblingAuthored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-shadow-scale-y-sibling.pptx",
            },
        });
        Assert.True(siblingProjected.Ok, Diagnostics(siblingProjected));
        var siblingElement = JsonNode.Parse(siblingProjected.PresentationProgram.ProgramJson.ToByteArray())!["pages"]![0]!["elements"]![0]!;
        Assert.DoesNotContain(siblingElement!["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textShadowScaleY");
    }

    private static JsonObject ShadowYProgram(double scaleY, bool siblingGlow = false)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName, "examples", "ppj", "minimum.ppj")))!.AsObject();
        var style = new JsonObject
        {
            ["shadow"] = new JsonObject
            {
                ["color"] = "#16324F",
                ["blur"] = 3,
                ["opacity"] = 0.72,
                ["distance"] = 1.5,
                ["angle"] = 90,
                ["scaleY"] = scaleY,
            },
        };
        if (siblingGlow)
            style["glow"] = new JsonObject { ["color"] = "#D9A514", ["radius"] = 8 };
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["runs"] = new JsonArray
                    {
                        new JsonObject { ["text"] = "Evidence changed the decision", ["style"] = style },
                    },
                },
            },
        };
        return program;
    }

    private static CodecResponse CompileScaleYShadow(JsonObject program) => Invoke(new CodecRequest
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Operation = CodecOperation.CompilePpjToPptx,
        Family = ArtifactFamily.Presentation,
        PresentationProgram = new PresentationProgramRequest
        {
            ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
        },
    });
}
