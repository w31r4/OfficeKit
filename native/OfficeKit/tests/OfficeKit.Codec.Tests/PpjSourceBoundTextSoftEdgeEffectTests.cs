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
    public void PpjSourceBoundTextSoftEdgeEditsDirectRunOwnerAndReprojects()
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
        program["design"]!["theme"]!["colors"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "accent1",
            ["value"] = "#0B8F8F",
        });
        program["design"]!["theme"]!["colors"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "accent2",
            ["value"] = "#8F4B1F",
        });
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["style"] = new JsonObject
                    {
                        ["defaultText"] = new JsonObject
                        {
                            ["softEdge"] = new JsonObject { ["radius"] = 4 },
                        },
                    },
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Evidence changed the decision",
                            ["style"] = new JsonObject
                            {
                                ["softEdge"] = new JsonObject { ["radius"] = 8 },
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["opacity"] = 0.42,
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 90,
                                    ["alignment"] = "br",
                                },
                            },
                        },
                    },
                },
                new JsonObject
                {
                    ["style"] = new JsonObject
                    {
                        ["defaultText"] = new JsonObject
                        {
                            ["glow"] = new JsonObject
                            {
                                ["color"] = "#0B8F8F",
                                ["radius"] = 4,
                                ["opacity"] = 0.34,
                            },
                        },
                    },
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Default glow owner",
                        },
                    },
                },
                new JsonObject
                {
                    ["style"] = new JsonObject
                    {
                        ["defaultText"] = new JsonObject
                        {
                            ["innerShadow"] = new JsonObject
                            {
                                ["color"] = "#8F4B1F",
                                ["blur"] = 3,
                                ["distance"] = 1,
                                ["angle"] = 120,
                                ["opacity"] = 0.36,
                            },
                        },
                    },
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Default inner-shadow owner",
                        },
                    },
                },
                new JsonObject
                {
                    ["style"] = new JsonObject
                    {
                        ["defaultText"] = new JsonObject
                        {
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
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["text"] = "Default reflection owner",
                        },
                    },
                },
            },
        };

        program["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["type"] = "text",
            ["id"] = "themed-inner-shadow",
            ["frame"] = new JsonObject
            {
                ["x"] = 48,
                ["y"] = 150,
                ["width"] = 624,
                ["height"] = 72,
            },
            ["text"] = new JsonObject
            {
                ["paragraphs"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["style"] = new JsonObject
                        {
                            ["defaultText"] = new JsonObject
                            {
                                ["innerShadow"] = new JsonObject
                                {
                                    ["color"] = new JsonObject { ["token"] = "accent1" },
                                    ["blur"] = 3,
                                    ["distance"] = 1,
                                    ["angle"] = 120,
                                },
                            },
                        },
                        ["runs"] = new JsonArray
                        {
                            new JsonObject
                            {
                                ["text"] = "Default themed inner-shadow owner",
                            },
                        },
                    },
                    new JsonObject
                    {
                        ["style"] = new JsonObject
                        {
                            ["defaultText"] = new JsonObject
                            {
                                ["glow"] = new JsonObject
                                {
                                    ["color"] = new JsonObject { ["token"] = "accent1" },
                                    ["radius"] = 4,
                                },
                            },
                        },
                        ["runs"] = new JsonArray
                        {
                            new JsonObject
                            {
                                ["text"] = "Default themed glow owner",
                            },
                        },
                    },
                },
            },
        });

        program["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["type"] = "text",
            ["id"] = "default-outer-shadow",
            ["frame"] = new JsonObject
            {
                ["x"] = 48,
                ["y"] = 240,
                ["width"] = 624,
                ["height"] = 72,
            },
            ["text"] = new JsonObject
            {
                ["paragraphs"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["style"] = new JsonObject
                        {
                            ["defaultText"] = new JsonObject
                            {
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = "#16324F",
                                    ["opacity"] = 0.42,
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 90,
                                    ["alignment"] = "br",
                                    ["rotateWithShape"] = true,
                                },
                            },
                        },
                        ["runs"] = new JsonArray
                        {
                            new JsonObject
                            {
                                ["text"] = "Default outer-shadow owner",
                            },
                        },
                    },
                },
            },
        });

        program["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["type"] = "text",
            ["id"] = "themed-default-outer-shadow",
            ["frame"] = new JsonObject
            {
                ["x"] = 48,
                ["y"] = 330,
                ["width"] = 624,
                ["height"] = 72,
            },
            ["text"] = new JsonObject
            {
                ["paragraphs"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["style"] = new JsonObject
                        {
                            ["defaultText"] = new JsonObject
                            {
                                ["shadow"] = new JsonObject
                                {
                                    ["color"] = new JsonObject { ["token"] = "accent1" },
                                    ["blur"] = 2,
                                    ["distance"] = 1,
                                    ["angle"] = 90,
                                    ["alignment"] = "br",
                                },
                            },
                        },
                        ["runs"] = new JsonArray
                        {
                            new JsonObject
                            {
                                ["text"] = "Themed default outer-shadow owner",
                            },
                        },
                    },
                },
            },
        });

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
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>()
                .Single(item => item.GetFirstChild<A.Text>()?.Text == "Evidence changed the decision");
            var effects = run.RunProperties!.GetFirstChild<A.EffectList>()!;
            Assert.Collection(
                effects.ChildElements,
                shadow => Assert.IsType<A.OuterShadow>(shadow),
                softEdge => Assert.Equal(101_600U, Assert.IsType<A.SoftEdge>(softEdge).Radius!.Value));
            var paragraphs = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().ToArray();
            var defaultRunProperties = paragraphs[0]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultEffects = defaultRunProperties.GetFirstChild<A.EffectList>()!;
            var defaultSoftEdge = Assert.IsType<A.SoftEdge>(Assert.Single(defaultEffects.ChildElements));
            Assert.Equal(50_800U, defaultSoftEdge.Radius!.Value);
            var defaultGlowProperties = paragraphs[1]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultGlowEffects = defaultGlowProperties.GetFirstChild<A.EffectList>()!;
            var defaultGlow = Assert.IsType<A.Glow>(Assert.Single(defaultGlowEffects.ChildElements));
            Assert.Equal(50_800U, defaultGlow.Radius!.Value);
            Assert.Equal("0B8F8F", defaultGlow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(34_000, defaultGlow.Descendants<A.Alpha>().Single().Val!.Value);
            var defaultInnerShadowProperties = paragraphs[2]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultInnerShadowEffects = defaultInnerShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultInnerShadow = Assert.IsType<A.InnerShadow>(Assert.Single(defaultInnerShadowEffects.ChildElements));
            Assert.Equal(38_100U, defaultInnerShadow.BlurRadius!.Value);
            Assert.Equal("8F4B1F", defaultInnerShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(36_000, defaultInnerShadow.Descendants<A.Alpha>().Single().Val!.Value);
            var defaultReflectionProperties = paragraphs[3]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultReflectionEffects = defaultReflectionProperties.GetFirstChild<A.EffectList>()!;
            var defaultReflection = Assert.IsType<A.Reflection>(Assert.Single(defaultReflectionEffects.ChildElements));
            Assert.Equal(63_500, defaultReflection.BlurRadius!.Value);
            Assert.Equal(42_000, defaultReflection.StartOpacity!.Value);
            Assert.Equal(8_000, defaultReflection.EndAlpha!.Value);
            Assert.Equal(0, defaultReflection.StartPosition!.Value);
            Assert.Equal(100_000, defaultReflection.EndPosition!.Value);
            Assert.Equal(152_400, defaultReflection.Distance!.Value);
            Assert.Equal(2_700_000, defaultReflection.Direction!.Value);
            var defaultThemedInnerShadowProperties = paragraphs[4]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedInnerShadowEffects = defaultThemedInnerShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedInnerShadow = Assert.IsType<A.InnerShadow>(Assert.Single(defaultThemedInnerShadowEffects.ChildElements));
            Assert.Equal(A.SchemeColorValues.Accent1, defaultThemedInnerShadow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            var defaultThemedGlowProperties = paragraphs[5]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedGlowEffects = defaultThemedGlowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedGlow = Assert.IsType<A.Glow>(Assert.Single(defaultThemedGlowEffects.ChildElements));
            Assert.Equal(50_800U, defaultThemedGlow.Radius!.Value);
            Assert.Equal(A.SchemeColorValues.Accent1, defaultThemedGlow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            var defaultShadowProperties = paragraphs[6]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultShadowEffects = defaultShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultShadow = Assert.IsType<A.OuterShadow>(Assert.Single(defaultShadowEffects.ChildElements));
            Assert.Equal(25_400U, defaultShadow.BlurRadius!.Value);
            Assert.Equal(12_700U, defaultShadow.Distance!.Value);
            Assert.Equal(5_400_000, defaultShadow.Direction!.Value);
            Assert.Equal(A.RectangleAlignmentValues.BottomRight, defaultShadow.Alignment!.Value);
            Assert.True(defaultShadow.RotateWithShape!.Value);
            Assert.Equal(42_000, defaultShadow.Descendants<A.Alpha>().Single().Val!.Value);
            var defaultThemedShadowProperties = paragraphs[7]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedShadowEffects = defaultThemedShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedShadow = Assert.IsType<A.OuterShadow>(Assert.Single(defaultThemedShadowEffects.ChildElements));
            Assert.Equal(A.SchemeColorValues.Accent1, defaultThemedShadow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
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
                SourceUri = "deck.assets/source/text-soft-edge.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var projectedParagraph = projectedElement["text"]!["paragraphs"]![0]!;
        Assert.Equal(4, projectedParagraph["style"]!["defaultText"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(8, projectedParagraph["runs"]![0]!["style"]!["softEdge"]!["radius"]!.GetValue<double>());
        var projectedGlowParagraph = projectedElement["text"]!["paragraphs"]![1]!;
        Assert.Equal(4, projectedGlowParagraph["style"]!["defaultText"]!["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal("#0B8F8F", projectedGlowParagraph["style"]!["defaultText"]!["glow"]!["color"]!.GetValue<string>());
        var projectedInnerShadowParagraph = projectedElement["text"]!["paragraphs"]![2]!;
        Assert.Equal(3, projectedInnerShadowParagraph["style"]!["defaultText"]!["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.36, projectedInnerShadowParagraph["style"]!["defaultText"]!["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);
        var projectedReflectionParagraph = projectedElement["text"]!["paragraphs"]![3]!;
        Assert.Equal(5, projectedReflectionParagraph["style"]!["defaultText"]!["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.42, projectedReflectionParagraph["style"]!["defaultText"]!["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.08, projectedReflectionParagraph["style"]!["defaultText"]!["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(12, projectedReflectionParagraph["style"]!["defaultText"]!["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(45, projectedReflectionParagraph["style"]!["defaultText"]!["reflection"]!["angle"]!.GetValue<double>());
        var softEdgeLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textSoftEdgeRadiusEmu")
            .ToArray();
        Assert.Single(softEdgeLeaves);
        Assert.Equal(101_600, softEdgeLeaves[0]["value"]!.GetValue<long>());
        var defaultSoftEdgeLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultSoftEdgeRadiusEmu")
            .ToArray();
        Assert.Single(defaultSoftEdgeLeaves);
        Assert.Equal(50_800, defaultSoftEdgeLeaves[0]["value"]!.GetValue<long>());
        var defaultGlowLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultGlowRadiusEmu")
            .ToArray();
        Assert.Single(defaultGlowLeaves);
        Assert.Equal(50_800, defaultGlowLeaves[0]["value"]!.GetValue<long>());
        var defaultGlowColorLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultGlowColorRgb")
            .ToArray();
        Assert.Single(defaultGlowColorLeaves);
        Assert.Equal("#0b8f8f", defaultGlowColorLeaves[0]["value"]!.GetValue<string>());
        Assert.Equal(0.34, projectedGlowParagraph["style"]!["defaultText"]!["glow"]!["opacity"]!.GetValue<double>(), precision: 6);
        var defaultGlowOpacityLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultGlowOpacityThousandthPercent")
            .ToArray();
        Assert.Single(defaultGlowOpacityLeaves);
        Assert.Equal(34_000, defaultGlowOpacityLeaves[0]["value"]!.GetValue<long>());
        var defaultInnerShadowLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowBlurRadiusEmu")
            .ToArray();
        Assert.Single(defaultInnerShadowLeaves);
        Assert.Equal(38_100, defaultInnerShadowLeaves[0]["value"]!.GetValue<long>());
        var defaultInnerShadowDistanceLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowDistanceEmu")
            .ToArray();
        Assert.Single(defaultInnerShadowDistanceLeaves);
        Assert.Equal(12_700, defaultInnerShadowDistanceLeaves[0]["value"]!.GetValue<long>());
        var defaultInnerShadowDirectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowDirectionDegrees")
            .ToArray();
        Assert.Single(defaultInnerShadowDirectionLeaves);
        Assert.Equal(7_200_000, defaultInnerShadowDirectionLeaves[0]["value"]!.GetValue<long>());
        var defaultInnerShadowColorLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowColorRgb")
            .ToArray();
        Assert.Single(defaultInnerShadowColorLeaves);
        Assert.Equal("#8f4b1f", defaultInnerShadowColorLeaves[0]["value"]!.GetValue<string>());
        var defaultInnerShadowOpacityLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowOpacityThousandthPercent")
            .ToArray();
        Assert.Single(defaultInnerShadowOpacityLeaves);
        Assert.Equal(36_000, defaultInnerShadowOpacityLeaves[0]["value"]!.GetValue<long>());
        var projectedThemedElement = projectedProgram["pages"]![0]!["elements"]![1]!.AsObject();
        var projectedThemedInnerShadowParagraph = projectedThemedElement["text"]!["paragraphs"]![0]!;
        Assert.Equal("accent1", projectedThemedInnerShadowParagraph["style"]!["defaultText"]!["innerShadow"]!["color"]!["token"]!.GetValue<string>());
        var defaultInnerShadowColorSchemeLeaves = projectedThemedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultInnerShadowColorScheme")
            .ToArray();
        Assert.Single(defaultInnerShadowColorSchemeLeaves);
        Assert.Equal("accent1", defaultInnerShadowColorSchemeLeaves[0]["value"]!.GetValue<string>());
        var projectedThemedGlowParagraph = projectedThemedElement["text"]!["paragraphs"]![1]!;
        Assert.Equal("accent1", projectedThemedGlowParagraph["style"]!["defaultText"]!["glow"]!["color"]!["token"]!.GetValue<string>());
        var defaultGlowColorSchemeLeaves = projectedThemedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultGlowColorScheme")
            .ToArray();
        Assert.Single(defaultGlowColorSchemeLeaves);
        Assert.Equal("accent1", defaultGlowColorSchemeLeaves[0]["value"]!.GetValue<string>());
        var projectedShadowElement = projectedProgram["pages"]![0]!["elements"]![2]!.AsObject();
        Assert.Equal(2, projectedShadowElement["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["blur"]!.GetValue<double>());
        var defaultShadowLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowBlurRadiusEmu")
            .ToArray();
        Assert.Single(defaultShadowLeaves);
        Assert.Equal(25_400, defaultShadowLeaves[0]["value"]!.GetValue<long>());
        var defaultShadowDistanceLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowDistanceEmu")
            .ToArray();
        Assert.Single(defaultShadowDistanceLeaves);
        Assert.Equal(12_700, defaultShadowDistanceLeaves[0]["value"]!.GetValue<long>());
        var defaultShadowDirectionLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowDirectionDegrees")
            .ToArray();
        Assert.Single(defaultShadowDirectionLeaves);
        Assert.Equal(5_400_000, defaultShadowDirectionLeaves[0]["value"]!.GetValue<long>());
        var defaultShadowAlignmentLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowAlignment")
            .ToArray();
        Assert.Single(defaultShadowAlignmentLeaves);
        Assert.Equal("br", defaultShadowAlignmentLeaves[0]["value"]!.GetValue<string>());
        var defaultShadowColorLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowColorRgb")
            .ToArray();
        Assert.Single(defaultShadowColorLeaves);
        Assert.Equal("#16324f", defaultShadowColorLeaves[0]["value"]!.GetValue<string>());
        var defaultShadowOpacityLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowOpacityThousandthPercent")
            .ToArray();
        Assert.Single(defaultShadowOpacityLeaves);
        Assert.Equal(42_000, defaultShadowOpacityLeaves[0]["value"]!.GetValue<long>());
        var defaultShadowRotateWithShapeLeaves = projectedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowRotateWithShape")
            .ToArray();
        Assert.Single(defaultShadowRotateWithShapeLeaves);
        Assert.True(defaultShadowRotateWithShapeLeaves[0]["value"]!.GetValue<bool>());
        var projectedThemedShadowElement = projectedProgram["pages"]![0]!["elements"]![3]!.AsObject();
        Assert.Equal("accent1", projectedThemedShadowElement["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["color"]!["token"]!.GetValue<string>());
        var defaultShadowColorSchemeLeaves = projectedThemedShadowElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultShadowColorScheme")
            .ToArray();
        Assert.Single(defaultShadowColorSchemeLeaves);
        Assert.Equal("accent1", defaultShadowColorSchemeLeaves[0]["value"]!.GetValue<string>());
        var defaultReflectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultReflectionBlurRadiusEmu")
            .ToArray();
        Assert.Single(defaultReflectionLeaves);
        Assert.Equal(63_500, defaultReflectionLeaves[0]["value"]!.GetValue<long>());
        var defaultReflectionDistanceLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultReflectionDistanceEmu")
            .ToArray();
        Assert.Single(defaultReflectionDistanceLeaves);
        Assert.Equal(152_400, defaultReflectionDistanceLeaves[0]["value"]!.GetValue<long>());
        var defaultReflectionStartOpacityLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultReflectionStartOpacityThousandthPercent")
            .ToArray();
        Assert.Single(defaultReflectionStartOpacityLeaves);
        Assert.Equal(42_000, defaultReflectionStartOpacityLeaves[0]["value"]!.GetValue<long>());
        var defaultReflectionEndOpacityLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultReflectionEndOpacityThousandthPercent")
            .ToArray();
        Assert.Single(defaultReflectionEndOpacityLeaves);
        Assert.Equal(8_000, defaultReflectionEndOpacityLeaves[0]["value"]!.GetValue<long>());
        var defaultReflectionDirectionLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textDefaultReflectionDirectionDegrees")
            .ToArray();
        Assert.Single(defaultReflectionDirectionLeaves);
        Assert.Equal(2_700_000, defaultReflectionDirectionLeaves[0]["value"]!.GetValue<long>());

        softEdgeLeaves[0]["value"] = 127_000;
        defaultSoftEdgeLeaves[0]["value"] = 76_200;
        defaultGlowLeaves[0]["value"] = 76_200;
        defaultGlowColorLeaves[0]["value"] = "#1A2B3C";
        defaultGlowOpacityLeaves[0]["value"] = 66_000;
        defaultInnerShadowLeaves[0]["value"] = 76_200;
        defaultInnerShadowDistanceLeaves[0]["value"] = 25_400;
        defaultInnerShadowDirectionLeaves[0]["value"] = 10_800_000;
        defaultInnerShadowColorLeaves[0]["value"] = "#1A2B3C";
        defaultInnerShadowOpacityLeaves[0]["value"] = 66_000;
        defaultInnerShadowColorSchemeLeaves[0]["value"] = "accent2";
        defaultGlowColorSchemeLeaves[0]["value"] = "accent2";
        defaultShadowLeaves[0]["value"] = 50_800;
        defaultShadowDistanceLeaves[0]["value"] = 25_400;
        defaultShadowDirectionLeaves[0]["value"] = 10_800_000;
        defaultShadowAlignmentLeaves[0]["value"] = "tl";
        defaultShadowColorLeaves[0]["value"] = "#1A2B3C";
        defaultShadowOpacityLeaves[0]["value"] = 66_000;
        defaultShadowRotateWithShapeLeaves[0]["value"] = false;
        defaultShadowColorSchemeLeaves[0]["value"] = "accent2";
        defaultReflectionLeaves[0]["value"] = 101_600;
        defaultReflectionDistanceLeaves[0]["value"] = 254_000;
        defaultReflectionStartOpacityLeaves[0]["value"] = 50_000;
        defaultReflectionEndOpacityLeaves[0]["value"] = 20_000;
        defaultReflectionDirectionLeaves[0]["value"] = 5_400_000;
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
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>()
                .Single(item => item.GetFirstChild<A.Text>()?.Text == "Evidence changed the decision");
            var effects = run.RunProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.OuterShadow>(effects.ChildElements[0]);
            Assert.Equal(127_000U, Assert.IsType<A.SoftEdge>(effects.ChildElements[1]).Radius!.Value);
            var paragraphs = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().ToArray();
            var defaultRunProperties = paragraphs[0]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultEffects = defaultRunProperties.GetFirstChild<A.EffectList>()!;
            Assert.Equal(76_200U, Assert.IsType<A.SoftEdge>(Assert.Single(defaultEffects.ChildElements)).Radius!.Value);
            var defaultGlowProperties = paragraphs[1]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultGlowEffects = defaultGlowProperties.GetFirstChild<A.EffectList>()!;
            var defaultGlow = Assert.IsType<A.Glow>(Assert.Single(defaultGlowEffects.ChildElements));
            Assert.Equal(76_200U, defaultGlow.Radius!.Value);
            Assert.Equal("1A2B3C", defaultGlow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, defaultGlow.Descendants<A.Alpha>().Single().Val!.Value);
            var defaultInnerShadowProperties = paragraphs[2]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultInnerShadowEffects = defaultInnerShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultInnerShadow = Assert.IsType<A.InnerShadow>(Assert.Single(defaultInnerShadowEffects.ChildElements));
            Assert.Equal(76_200U, defaultInnerShadow.BlurRadius!.Value);
            Assert.Equal(25_400U, defaultInnerShadow.Distance!.Value);
            Assert.Equal(10_800_000, defaultInnerShadow.Direction!.Value);
            Assert.Equal("1A2B3C", defaultInnerShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, defaultInnerShadow.Descendants<A.Alpha>().Single().Val!.Value);
            var defaultThemedInnerShadowProperties = paragraphs[4]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedInnerShadowEffects = defaultThemedInnerShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedInnerShadow = Assert.IsType<A.InnerShadow>(Assert.Single(defaultThemedInnerShadowEffects.ChildElements));
            Assert.Equal(A.SchemeColorValues.Accent2, defaultThemedInnerShadow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            var defaultThemedGlowProperties = paragraphs[5]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedGlowEffects = defaultThemedGlowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedGlow = Assert.IsType<A.Glow>(Assert.Single(defaultThemedGlowEffects.ChildElements));
            Assert.Equal(A.SchemeColorValues.Accent2, defaultThemedGlow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            var defaultShadowProperties = paragraphs[6]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultShadowEffects = defaultShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultShadow = Assert.IsType<A.OuterShadow>(Assert.Single(defaultShadowEffects.ChildElements));
            Assert.Equal(50_800U, defaultShadow.BlurRadius!.Value);
            Assert.Equal(25_400U, defaultShadow.Distance!.Value);
            Assert.Equal(10_800_000, defaultShadow.Direction!.Value);
            Assert.Equal(A.RectangleAlignmentValues.TopLeft, defaultShadow.Alignment!.Value);
            Assert.Equal("1A2B3C", defaultShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, defaultShadow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.False(defaultShadow.RotateWithShape!.Value);
            var defaultThemedShadowProperties = paragraphs[7]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultThemedShadowEffects = defaultThemedShadowProperties.GetFirstChild<A.EffectList>()!;
            var defaultThemedShadow = Assert.IsType<A.OuterShadow>(Assert.Single(defaultThemedShadowEffects.ChildElements));
            Assert.Equal(A.SchemeColorValues.Accent2, defaultThemedShadow.GetFirstChild<A.SchemeColor>()!.Val!.Value);
            var defaultReflectionProperties = paragraphs[3]
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultReflectionEffects = defaultReflectionProperties.GetFirstChild<A.EffectList>()!;
            var defaultReflection = Assert.IsType<A.Reflection>(Assert.Single(defaultReflectionEffects.ChildElements));
            Assert.Equal(101_600, defaultReflection.BlurRadius!.Value);
            Assert.Equal(50_000, defaultReflection.StartOpacity!.Value);
            Assert.Equal(20_000, defaultReflection.EndAlpha!.Value);
            Assert.Equal(0, defaultReflection.StartPosition!.Value);
            Assert.Equal(100_000, defaultReflection.EndPosition!.Value);
            Assert.Equal(254_000, defaultReflection.Distance!.Value);
            Assert.Equal(5_400_000, defaultReflection.Direction!.Value);
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
                SourceUri = "deck.assets/edited/text-soft-edge.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedParagraph = reprojectedElement["text"]!["paragraphs"]![0]!;
        Assert.Equal(6, reprojectedParagraph["style"]!["defaultText"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(10, reprojectedParagraph["runs"]![0]!["style"]!["softEdge"]!["radius"]!.GetValue<double>());
        Assert.Equal(6, reprojectedElement["text"]!["paragraphs"]![1]!["style"]!["defaultText"]!["glow"]!["radius"]!.GetValue<double>());
        Assert.Equal("#1A2B3C", reprojectedElement["text"]!["paragraphs"]![1]!["style"]!["defaultText"]!["glow"]!["color"]!.GetValue<string>());
        Assert.Equal(0.66, reprojectedElement["text"]!["paragraphs"]![1]!["style"]!["defaultText"]!["glow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(6, reprojectedElement["text"]!["paragraphs"]![2]!["style"]!["defaultText"]!["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(2, reprojectedElement["text"]!["paragraphs"]![2]!["style"]!["defaultText"]!["innerShadow"]!["distance"]!.GetValue<double>());
        Assert.Equal(180, reprojectedElement["text"]!["paragraphs"]![2]!["style"]!["defaultText"]!["innerShadow"]!["angle"]!.GetValue<double>());
        Assert.Equal("#1A2B3C", reprojectedElement["text"]!["paragraphs"]![2]!["style"]!["defaultText"]!["innerShadow"]!["color"]!.GetValue<string>());
        Assert.Equal(0.66, reprojectedElement["text"]!["paragraphs"]![2]!["style"]!["defaultText"]!["innerShadow"]!["opacity"]!.GetValue<double>(), precision: 6);
        var reprojectedThemedElement = reprojectedProgram["pages"]![0]!["elements"]![1]!.AsObject();
        Assert.Equal("accent2", reprojectedThemedElement["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["innerShadow"]!["color"]!["token"]!.GetValue<string>());
        Assert.Equal("accent2", reprojectedThemedElement["text"]!["paragraphs"]![1]!["style"]!["defaultText"]!["glow"]!["color"]!["token"]!.GetValue<string>());
        Assert.Equal(4, reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(2, reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["distance"]!.GetValue<double>());
        Assert.Equal(180, reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["angle"]!.GetValue<double>());
        Assert.Equal("tl", reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["alignment"]!.GetValue<string>());
        Assert.Equal("#1A2B3C", reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["color"]!.GetValue<string>());
        Assert.Equal(0.66, reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.False(reprojectedProgram["pages"]![0]!["elements"]![2]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["rotateWithShape"]!.GetValue<bool>());
        Assert.Equal("accent2", reprojectedProgram["pages"]![0]!["elements"]![3]!["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["shadow"]!["color"]!["token"]!.GetValue<string>());
        Assert.Equal(8, reprojectedElement["text"]!["paragraphs"]![3]!["style"]!["defaultText"]!["reflection"]!["blur"]!.GetValue<double>());
        Assert.Equal(20, reprojectedElement["text"]!["paragraphs"]![3]!["style"]!["defaultText"]!["reflection"]!["distance"]!.GetValue<double>());
        Assert.Equal(0.5, reprojectedElement["text"]!["paragraphs"]![3]!["style"]!["defaultText"]!["reflection"]!["startOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.2, reprojectedElement["text"]!["paragraphs"]![3]!["style"]!["defaultText"]!["reflection"]!["endOpacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, reprojectedElement["text"]!["paragraphs"]![3]!["style"]!["defaultText"]!["reflection"]!["angle"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundTextSoftEdgeLeavesStayOpaqueForUnsupportedEffectChain()
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
                                ["softEdge"] = new JsonObject { ["radius"] = 8 },
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
                SourceUri = "deck.assets/source/text-soft-edge-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(element["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["softEdge"]);
        Assert.DoesNotContain(
            element["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "textSoftEdgeRadiusEmu");
    }
}
