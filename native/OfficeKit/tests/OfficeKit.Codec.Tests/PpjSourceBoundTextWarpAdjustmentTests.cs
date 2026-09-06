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
    public void PpjSourceBoundTextWarpAdjustmentEditsLeafAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = program["pages"]!.AsArray()[0]!.DeepClone()!.AsObject();
        var textShape = page["elements"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-band");
        textShape["textStyle"] = new JsonObject
        {
            ["textWarpPreset"] = "textArchUp",
            ["textWarpAdjustments"] = new JsonArray(new JsonObject
            {
                ["name"] = "adj",
                ["value"] = 50_000,
            }),
        };
        page["elements"] = new JsonArray(textShape.DeepClone());
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        program["assets"] = new JsonArray();
        program["components"] = new JsonArray();
        program["pages"] = new JsonArray(page);
        program["sections"] = new JsonArray();
        program["customShows"] = new JsonArray();
        program["comments"] = new JsonArray();

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

        long authoredTextLength;
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var bodyProperties = nativeShape.TextBody!.GetFirstChild<A.BodyProperties>()!;
            var textWarp = bodyProperties.GetFirstChild<A.PresetTextWarp>()!;
            Assert.Equal("textArchUp", textWarp.GetAttribute("prst", string.Empty).Value);
            var guide = Assert.Single(textWarp.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>());
            Assert.Equal("adj", guide.Name?.Value);
            Assert.Equal("val 50000", guide.Formula?.Value);
            authoredTextLength = nativeShape.TextBody.InnerText.Length;
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
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
                SourceUri = "deck.assets/source/text-body-warp-adjustment.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal("textArchUp", projectedShape["textStyle"]!["textWarpPreset"]!.GetValue<string>());
        var projectedAdjustment = Assert.Single(projectedShape["textStyle"]!["textWarpAdjustments"]!.AsArray()).AsObject();
        Assert.Equal("adj", projectedAdjustment["name"]!.GetValue<string>());
        Assert.Equal(50_000, projectedAdjustment["value"]!.GetValue<int>());
        var leaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textBodyWarpAdjustment")
            .ToArray();
        Assert.Single(leaves);
        Assert.Equal(50_000, leaves[0]["value"]!.GetValue<long>());
        Assert.Contains(projectedShape["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyWarpPreset");

        leaves[0]["value"] = 60_000;
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

        var editedBytes = edited.File.ToByteArray();
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var bodyProperties = nativeShape.TextBody!.GetFirstChild<A.BodyProperties>()!;
            var textWarp = bodyProperties.GetFirstChild<A.PresetTextWarp>()!;
            Assert.Equal("textArchUp", textWarp.GetAttribute("prst", string.Empty).Value);
            var guide = Assert.Single(textWarp.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>());
            Assert.Equal("adj", guide.Name?.Value);
            Assert.Equal("val 60000", guide.Formula?.Value);
            Assert.Equal(authoredTextLength, nativeShape.TextBody.InnerText.Length);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

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
                SourceUri = "deck.assets/edited/text-body-warp-adjustment.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal(60_000, reprojectedShape["textStyle"]!["textWarpAdjustments"]!.AsArray().Single()!["value"]!.GetValue<int>());
        Assert.Equal(60_000, reprojectedShape["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyWarpAdjustment")["value"]!.GetValue<long>());

        var unsupportedSources = new[]
        {
            ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
                xml.Replace("fmla=\"val 50000\"", "fmla=\"pin 0 1 2\"", StringComparison.Ordinal)),
            ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
            {
                var close = xml.IndexOf("</a:avLst>", StringComparison.Ordinal);
                Assert.True(close >= 0);
                return xml[..close] + "<a:gd name=\"adj\" fmla=\"val 60000\"/>" + xml[close..];
            }),
            ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
            {
                var close = xml.IndexOf("</a:prstTxWarp>", StringComparison.Ordinal);
                Assert.True(close >= 0);
                return xml[..close] + "<a:extLst/>" + xml[close..];
            }),
        };
        for (var index = 0; index < unsupportedSources.Length; index++)
        {
            var unsupported = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ProjectPptxToPpj,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(unsupportedSources[index]),
                PresentationProgram = new PresentationProgramRequest
                {
                    SourceUri = $"deck.assets/source/text-body-warp-adjustment-unsupported-{index}.pptx",
                },
            });
            Assert.True(unsupported.Ok, Diagnostics(unsupported));
            var unsupportedProgram = JsonNode.Parse(unsupported.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
            var unsupportedShape = unsupportedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
            Assert.DoesNotContain(unsupportedShape["nativeRef"]!["leaves"]!.AsArray(), leaf =>
                leaf!["kind"]!.GetValue<string>() == "textBodyWarpAdjustment");
        }
    }
}
