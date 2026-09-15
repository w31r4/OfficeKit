using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using Google.Protobuf;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundThemeAccent2RgbaEditsOnlyAccent2AndReprojects()
    {
        var authoredRequest = ExportRequest();
        var authoredTheme = new PresentationThemeArtifact();
        authoredTheme.AccentRgb.Add(new[] { "112233", "22334480", "334455", "445566", "556677", "667788" });
        authoredRequest.Artifact.Presentation.AuthoredTheme = authoredTheme;
        var authored = Invoke(authoredRequest);
        Assert.True(authored.Ok, Diagnostics(authored));

        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/theme-accent2.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTheme = projectedProgram["design"]!["theme"]!.AsObject();
        var projectedColors = projectedTheme["accentColors"]!.AsObject();
        Assert.Equal("#112233", projectedColors["accent1"]!.GetValue<string>());
        Assert.Equal("#22334480", projectedColors["accent2"]!.GetValue<string>());
        Assert.Equal("#667788", projectedColors["accent6"]!.GetValue<string>());
        var nativeRef = Assert.IsType<JsonObject>(projectedTheme["nativeRef"]);
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeAccent2Color" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "accentColors.accent2"));

        var noOp = Invoke(new CodecRequest
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
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);
        Assert.Equal(source, noOp.File.ToByteArray());

        var editedProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        editedProgram["design"]!["theme"]!["accentColors"]!["accent2"] = "#ABCDEF40";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(editedProgram.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        const string themePartPath = "ppt/slideMasters/theme/theme1.xml";
        Assert.Equal([themePartPath], edited.PresentationProgram.ChangedParts);

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var scheme = Assert.Single(package.PresentationPart!.SlideMasterParts).ThemePart!.Theme!.ThemeElements!.ColorScheme!;
            Assert.Equal("112233", scheme.Accent1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("ABCDEF", scheme.Accent2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(25098, scheme.Accent2Color.RgbColorModelHex.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal("334455", scheme.Accent3Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("445566", scheme.Accent4Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("556677", scheme.Accent5Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("667788", scheme.Accent6Color!.RgbColorModelHex!.Val!.Value);
        }

        var editedBytes = edited.File.ToByteArray();
        foreach (var path in ZipPartPaths(source).Where(path => !path.Equals(themePartPath, StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, path), ZipBytes(editedBytes, path));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(editedBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/theme-accent2.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal("#ABCDEF40", reprojectedProgram["design"]!["theme"]!["accentColors"]!["accent2"]!.GetValue<string>());

        var alphaPresenceProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        alphaPresenceProgram["design"]!["theme"]!["accentColors"]!["accent2"] = "#ABCDEF";
        var alphaPresence = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(alphaPresenceProgram.ToJsonString()),
            },
        });
        Assert.False(alphaPresence.Ok, Diagnostics(alphaPresence));

        var unauthorizedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        unauthorizedProgram["design"]!["theme"]!["accentColors"]!["accent1"] = "#AABBCC";
        var unauthorized = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(unauthorizedProgram.ToJsonString()),
            },
        });
        Assert.False(unauthorized.Ok, Diagnostics(unauthorized));

        var tamperedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        tamperedProgram["design"]!["theme"]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(capability => capability!["operation"]!.GetValue<string>() == "setThemeAccent2Color")!["fields"] = new JsonArray("accentColors.accent1");
        var tampered = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(tamperedProgram.ToJsonString()),
            },
        });
        Assert.False(tampered.Ok, Diagnostics(tampered));
    }
}
