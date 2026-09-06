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
    public void PpjSourceBoundTextFlatTextZEditsLeafAndReprojects()
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
            ["flatTextZ"] = 1200,
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
            var flatText = Assert.Single(bodyProperties.ChildElements.OfType<A.FlatText>());
            Assert.Equal(1200L, flatText.Z!.Value);
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
                SourceUri = "deck.assets/source/text-body-flat-text-z.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal(1200, projectedShape["textStyle"]!["flatTextZ"]!.GetValue<int>());
        var leaves = projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "textBodyFlatTextZ")
            .ToArray();
        Assert.Single(leaves);
        Assert.Equal(1200L, leaves[0]["value"]!.GetValue<long>());

        leaves[0]["value"] = 2400;
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
            var flatText = Assert.Single(bodyProperties.ChildElements.OfType<A.FlatText>());
            Assert.Equal(2400L, flatText.Z!.Value);
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
                SourceUri = "deck.assets/edited/text-body-flat-text-z.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal(2400, reprojectedShape["textStyle"]!["flatTextZ"]!.GetValue<int>());
        Assert.Equal(2400L, reprojectedShape["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "textBodyFlatTextZ")["value"]!.GetValue<long>());
    }
}
