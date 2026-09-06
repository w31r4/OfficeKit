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
    public void PpjSourceBoundTextInnerShadowEditsDirectRunOwnerAndReprojects()
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
                                ["innerShadow"] = new JsonObject
                                {
                                    ["color"] = "#D9A514",
                                    ["blur"] = 5,
                                    ["distance"] = 2,
                                    ["angle"] = 45,
                                    ["opacity"] = 0.36,
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
                innerShadow =>
                {
                    var native = Assert.IsType<A.InnerShadow>(innerShadow);
                    Assert.Equal(63_500, native.BlurRadius!.Value);
                    Assert.Equal(25_400, native.Distance!.Value);
                    Assert.Equal(2_700_000, native.Direction!.Value);
                    Assert.Equal("D9A514", native.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                    Assert.Equal(36_000, native.Descendants<A.Alpha>().Single().Val!.Value);
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
                SourceUri = "deck.assets/source/text-inner-shadow.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var innerShadow = projectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["innerShadow"]!.AsObject();
        Assert.Equal("#D9A514", innerShadow["color"]!.GetValue<string>());
        Assert.Equal(5, innerShadow["blur"]!.GetValue<double>());
        Assert.Equal(2, innerShadow["distance"]!.GetValue<double>());
        Assert.Equal(45, innerShadow["angle"]!.GetValue<double>());
        Assert.Equal(0.36, innerShadow["opacity"]!.GetValue<double>(), precision: 6);

        var innerShadowLeaves = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>().StartsWith("textInnerShadow", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(5, innerShadowLeaves.Length);
        Assert.Equal(63_500, innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowBlurRadiusEmu")["value"]!.GetValue<long>());
        Assert.Equal(25_400, innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowDistanceEmu")["value"]!.GetValue<long>());
        Assert.Equal(45, innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowDirectionDegrees")["value"]!.GetValue<double>());
        Assert.Equal("#d9a514", innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowColorRgb")["value"]!.GetValue<string>());
        Assert.Equal(36_000, innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowOpacityThousandthPercent")["value"]!.GetValue<long>());

        innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowBlurRadiusEmu")["value"] = 101_600;
        innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowDistanceEmu")["value"] = 38_100;
        innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowDirectionDegrees")["value"] = 90;
        innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowColorRgb")["value"] = "#AABBCC";
        innerShadowLeaves.Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowOpacityThousandthPercent")["value"] = 66_000;

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
            var native = Assert.IsType<A.InnerShadow>(effects.ChildElements[0]);
            Assert.Equal(101_600, native.BlurRadius!.Value);
            Assert.Equal(38_100, native.Distance!.Value);
            Assert.Equal(5_400_000, native.Direction!.Value);
            Assert.Equal("AABBCC", native.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(66_000, native.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(effects.ChildElements[1]);
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
                SourceUri = "deck.assets/edited/text-inner-shadow.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reprojectedInnerShadow = reprojectedElement["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["innerShadow"]!;
        Assert.Equal("#AABBCC", reprojectedInnerShadow["color"]!.GetValue<string>());
        Assert.Equal(8, reprojectedInnerShadow["blur"]!.GetValue<double>());
        Assert.Equal(3, reprojectedInnerShadow["distance"]!.GetValue<double>());
        Assert.Equal(90, reprojectedInnerShadow["angle"]!.GetValue<double>());
        Assert.Equal(0.66, reprojectedInnerShadow["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(90, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "textInnerShadowDirectionDegrees")["value"]!.GetValue<double>());
    }

    [Fact]
    public void PpjSourceBoundTextInnerShadowLeavesStayOpaqueForComplexEffectList()
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
                SourceUri = "deck.assets/source/text-inner-shadow-complex.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Null(element["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]?["innerShadow"]);
        Assert.DoesNotContain(
            element["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>().StartsWith("textInnerShadow", StringComparison.Ordinal));
    }
}
