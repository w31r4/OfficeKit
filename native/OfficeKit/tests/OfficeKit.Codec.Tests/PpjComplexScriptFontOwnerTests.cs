using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
    public void PpjComplexScriptFontAuthorsProjectsAndEditsDirectRunOwner()
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
                                ["fontFamily"] = "Aptos",
                                ["fontFamilyEastAsia"] = "Noto Sans CJK SC",
                                ["fontFamilyComplexScript"] = "Noto Sans Arabic",
                            },
                        },
                    },
                },
            },
        };
        var programBytes = Encoding.UTF8.GetBytes(program.ToJsonString());
        var validation = PpjProgramValidator.Validate(programBytes);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics));

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFrom(programBytes),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
            Assert.Equal("Aptos", run.RunProperties!.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal("Noto Sans CJK SC", run.RunProperties.GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Equal("Noto Sans Arabic", run.RunProperties.GetFirstChild<A.ComplexScriptFont>()!.Typeface!.Value);
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
                SourceUri = "deck.assets/source/complex-script.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var owner = projectedProgram["pages"]![0]!["elements"]![0]!.AsObject();
        var complexLeaf = owner["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "fontFamilyComplexScript")!.AsObject();
        Assert.Equal("Noto Sans Arabic", complexLeaf["value"]!.GetValue<string>());

        complexLeaf["value"] = "Noto Sans Hebrew";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(projectedProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var run = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Run>().Single();
            Assert.Equal("Noto Sans Hebrew", run.RunProperties!.GetFirstChild<A.ComplexScriptFont>()!.Typeface!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/complex-script-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedLeaf = reprojectedProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "fontFamilyComplexScript");
        Assert.Equal("Noto Sans Hebrew", reprojectedLeaf!["value"]!.GetValue<string>());
    }
}
