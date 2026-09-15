using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundThemeFontSchemeNameEditsOnlyThemePartAndReprojects()
    {
        var authoredRequest = ExportRequest();
        authoredRequest.Artifact.Presentation.AuthoredTheme = new PresentationThemeArtifact
        {
            Name = "Aurora Theme",
            FontSchemeName = "Aurora Fonts",
        };
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
                SourceUri = "deck.assets/source/font-scheme-name.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTheme = projectedProgram["design"]!["theme"]!.AsObject();
        Assert.Equal("Aurora Theme", projectedTheme["name"]!.GetValue<string>());
        var projectedFontScheme = projectedTheme["fontScheme"]!.AsObject();
        Assert.Equal("Aurora Fonts", projectedFontScheme["name"]!.GetValue<string>());
        var nativeRef = Assert.IsType<JsonObject>(projectedTheme["nativeRef"]);
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeFontSchemeName" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "fontScheme.name"));

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
        editedProgram["design"]!["theme"]!["fontScheme"]!["name"] = "Nebula Fonts";
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
            var theme = Assert.Single(package.PresentationPart!.SlideMasterParts).ThemePart!.Theme!;
            Assert.Equal("Aurora Theme", theme.Name!.Value);
            Assert.Equal("Nebula Fonts", theme.ThemeElements!.FontScheme!.Name!.Value);
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
                SourceUri = "deck.assets/edited/font-scheme-name.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal(
            "Nebula Fonts",
            reprojectedProgram["design"]!["theme"]!["fontScheme"]!["name"]!.GetValue<string>());

        var deletionProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        deletionProgram["design"]!["theme"]!["fontScheme"]!.AsObject().Remove("name");
        var deletion = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(deletionProgram.ToJsonString()),
            },
        });
        Assert.False(deletion.Ok, Diagnostics(deletion));

        var combinedProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        combinedProgram["design"]!["theme"]!["fontScheme"]!["name"] = "Nebula Fonts";
        combinedProgram["design"]!["theme"]!["fontScheme"]!["major"] = "Calibri";
        var combined = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(combinedProgram.ToJsonString()),
            },
        });
        Assert.False(combined.Ok, Diagnostics(combined));

        var unauthorizedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        var nameCapability = unauthorizedProgram["design"]!["theme"]!["nativeRef"]!["capabilities"]!
            .AsArray()
            .Single(capability => capability!["operation"]!.GetValue<string>() == "setThemeFontSchemeName")!;
        nameCapability["fields"] = new JsonArray("fontScheme.major");
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
    }
}
