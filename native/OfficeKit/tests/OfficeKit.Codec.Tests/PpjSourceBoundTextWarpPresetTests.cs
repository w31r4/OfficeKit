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
    public void PpjSourceBoundTextWarpPresetEditsLeafAndReprojects()
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
                SourceUri = "deck.assets/source/text-body-warp-preset.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal("textArchUp", projectedShape["textStyle"]!["textWarpPreset"]!.GetValue<string>());
        var leaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textBodyWarpPreset")
            .ToArray();
        Assert.Single(leaves);
        Assert.Equal("textArchUp", leaves[0]["value"]!.GetValue<string>());

        leaves[0]["value"] = "textPlain";
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
            Assert.Equal("textPlain", textWarp.GetAttribute("prst", string.Empty).Value);
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
                SourceUri = "deck.assets/edited/text-body-warp-preset.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal("textPlain", reprojectedShape["textStyle"]!["textWarpPreset"]!.GetValue<string>());
        Assert.Equal("textPlain", reprojectedShape["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyWarpPreset")["value"]!.GetValue<string>());

        var adjustedSource = ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
        {
            var marker = xml.IndexOf("prstTxWarp", StringComparison.Ordinal);
            Assert.True(marker >= 0);
            var start = xml.LastIndexOf('<', marker);
            var end = xml.IndexOf('>', marker);
            Assert.True(start >= 0 && end > start && xml[end - 1] == '/');
            return xml[..(end - 1)] + "><a:avLst/></a:prstTxWarp>" + xml[(end + 1)..];
        });
        var adjusted = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(adjustedSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-body-warp-preset-adjusted.pptx",
            },
        });
        Assert.True(adjusted.Ok, Diagnostics(adjusted));
        var adjustedProgram = JsonNode.Parse(adjusted.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var adjustedShape = adjustedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.DoesNotContain(adjustedShape["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyWarpPreset");

        var absentSource = ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
        {
            var marker = xml.IndexOf("prstTxWarp", StringComparison.Ordinal);
            Assert.True(marker >= 0);
            var start = xml.LastIndexOf('<', marker);
            var end = xml.IndexOf('>', marker);
            Assert.True(start >= 0 && end > start);
            return xml.Remove(start, end - start + 1);
        });
        var absent = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(absentSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/text-body-warp-preset-absent.pptx",
            },
        });
        Assert.True(absent.Ok, Diagnostics(absent));
        var absentProgram = JsonNode.Parse(absent.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var absentShape = absentProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.DoesNotContain(absentShape["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyWarpPreset");
    }
}
