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
    public void PpjSourceBoundThemeMajorComplexScriptFontEditsOnlyOwnedSlotAndReprojects()
    {
        var authoredRequest = ExportRequest();
        authoredRequest.Artifact.Presentation.AuthoredTheme = new PresentationThemeArtifact
        {
            Name = "Aurora",
            MajorFontFamily = "Aptos Display",
            MinorFontFamily = "Aptos",
            MajorFontFamilyEastAsia = "Noto Sans CJK SC",
            MinorFontFamilyEastAsia = "Noto Sans CJK TC",
            MajorFontFamilyComplexScript = "Noto Sans Arabic",
            MinorFontFamilyComplexScript = "Noto Sans Hebrew",
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
                SourceUri = "deck.assets/source/theme-major-complex-script-font.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTheme = projectedProgram["design"]!["theme"]!.AsObject();
        var projectedFontScheme = projectedTheme["fontScheme"]!.AsObject();
        Assert.Equal("Aptos Display", projectedFontScheme["major"]!.GetValue<string>());
        Assert.Equal("Aptos", projectedFontScheme["minor"]!.GetValue<string>());
        Assert.Equal("Noto Sans CJK SC", projectedFontScheme["majorEastAsia"]!.GetValue<string>());
        Assert.Equal("Noto Sans CJK TC", projectedFontScheme["minorEastAsia"]!.GetValue<string>());
        Assert.Equal("Noto Sans Arabic", projectedFontScheme["majorComplexScript"]!.GetValue<string>());
        Assert.Equal("Noto Sans Hebrew", projectedFontScheme["minorComplexScript"]!.GetValue<string>());
        var nativeRef = Assert.IsType<JsonObject>(projectedTheme["nativeRef"]);
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeMajorFontComplexScript" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "fontScheme.majorComplexScript"));

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
        editedProgram["design"]!["theme"]!["fontScheme"]!["majorComplexScript"] = "Nebula Arabic";
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
            var fontScheme = theme.ThemeElements!.FontScheme!;
            Assert.Equal("Aptos Display", fontScheme.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Aptos", fontScheme.MinorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Noto Sans CJK SC", fontScheme.MajorFont.EastAsianFont!.Typeface!.Value);
            Assert.Equal("Noto Sans CJK TC", fontScheme.MinorFont.EastAsianFont!.Typeface!.Value);
            Assert.Equal("Nebula Arabic", fontScheme.MajorFont.ComplexScriptFont!.Typeface!.Value);
            Assert.Equal("Noto Sans Hebrew", fontScheme.MinorFont.ComplexScriptFont!.Typeface!.Value);
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
                SourceUri = "deck.assets/edited/theme-major-complex-script-font.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal("Nebula Arabic", reprojectedProgram["design"]!["theme"]!["fontScheme"]!["majorComplexScript"]!.GetValue<string>());

        var combinedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        combinedProgram["design"]!["theme"]!["fontScheme"]!["major"] = "Changed Major";
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

        var tamperedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        tamperedProgram["design"]!["theme"]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(capability => capability!["operation"]!.GetValue<string>() == "setThemeMajorFontComplexScript")!["fields"] = new JsonArray("fontScheme.major");
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
