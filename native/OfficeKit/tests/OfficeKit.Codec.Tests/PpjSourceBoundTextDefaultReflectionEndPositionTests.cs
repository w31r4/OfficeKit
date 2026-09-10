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
    public void PpjSourceBoundTextDefaultReflectionEndPositionEditsLeafAndReprojects()
    {
        var authored = CompileDefaultReflectionEndPositionProgram();
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        source = RewriteEndReflectionPosition(source, 0, 80_000);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-default-reflection-end-position.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var reflectionJson = projectedElement["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["reflection"]!.AsObject();
        Assert.Equal(0, reflectionJson["startPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.8, reflectionJson["endPosition"]!.GetValue<double>(), precision: 6);
        var leaf = projectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["kind"]!.GetValue<string>() == "textDefaultReflectionEndPosition");
        Assert.Equal(0.8, leaf["value"]!.GetValue<double>(), precision: 6);

        leaf["value"] = 0.65;
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
            var reflection = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().Single()
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!
                .GetFirstChild<A.EffectList>()!.GetFirstChild<A.Reflection>()!;
            Assert.Equal(0, reflection.StartPosition!.Value);
            Assert.Equal(65_000, reflection.EndPosition!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/text-default-reflection-end-position.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedElement = reprojectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Equal(0.65, reprojectedElement["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["reflection"]!["endPosition"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.65, reprojectedElement["nativeRef"]!["leaves"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["kind"]!.GetValue<string>() == "textDefaultReflectionEndPosition")["value"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundTextDefaultReflectionEndPositionStaysOpaqueForPartialSpan()
    {
        var authored = CompileDefaultReflectionEndPositionProgram();
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        source = RewriteEndReflectionPosition(source, 20_000, 80_000);

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-default-reflection-partial-span.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedElement = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.DoesNotContain(projectedElement["nativeRef"]!["leaves"]!.AsArray(), item =>
            item!["kind"]!.GetValue<string>() == "textDefaultReflectionEndPosition");
    }

    private static CodecResponse CompileDefaultReflectionEndPositionProgram()
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
                    ["style"] = new JsonObject
                    {
                        ["defaultText"] = new JsonObject
                        {
                            ["reflection"] = new JsonObject
                            {
                                ["startPosition"] = 0,
                                ["endPosition"] = 1,
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
                        new JsonObject { ["text"] = "Default reflection end position" },
                    },
                },
            },
        };
        return Invoke(new CodecRequest
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

    private static byte[] RewriteEndReflectionPosition(byte[] source, int startPosition, int endPosition)
    {
        var path = Path.Combine(Path.GetTempPath(), $"officekit-reflection-{Guid.NewGuid():N}.pptx");
        try
        {
            File.WriteAllBytes(path, source);
            using (var package = PresentationDocument.Open(path, true))
            {
                var reflection = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Paragraph>().Single()
                    .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!
                    .GetFirstChild<A.EffectList>()!.GetFirstChild<A.Reflection>()!;
                reflection.StartPosition = startPosition;
                reflection.EndPosition = endPosition;
            }
            return File.ReadAllBytes(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

}
