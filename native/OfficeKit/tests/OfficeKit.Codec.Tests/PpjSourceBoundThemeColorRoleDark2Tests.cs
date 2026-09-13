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
    public void PpjSourceBoundThemeColorRoleDark2EditsOnlyDark2AndReprojects()
    {
        var authoredRequest = ExportRequest();
        var authoredTheme = new PresentationThemeArtifact();
        authoredTheme.AccentRgb.Add(new[] { "112233", "223344", "334455", "445566", "556677", "667788" });
        authoredTheme.Dark1Rgb = "101112";
        authoredTheme.Light1Rgb = "202122";
        authoredTheme.Dark2Rgb = "303132";
        authoredTheme.Light2Rgb = "404142";
        authoredTheme.HyperlinkRgb = "505152";
        authoredTheme.FollowedHyperlinkRgb = "606162";
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
                SourceUri = "deck.assets/source/theme-color-role-dark2.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTheme = projectedProgram["design"]!["theme"]!.AsObject();
        var projectedColors = projectedTheme["accentColors"]!.AsObject();
        Assert.Equal("#112233", projectedColors["accent1"]!.GetValue<string>());
        Assert.Equal("#556677", projectedColors["accent5"]!.GetValue<string>());
        Assert.Equal("#667788", projectedColors["accent6"]!.GetValue<string>());
        var projectedColorRoles = projectedTheme["colorRoles"]!.AsObject();
        Assert.Equal("#101112", projectedColorRoles["dark1"]!.GetValue<string>());
        Assert.Equal("#202122", projectedColorRoles["light1"]!.GetValue<string>());
        Assert.Equal("#303132", projectedColorRoles["dark2"]!.GetValue<string>());
        var nativeRef = Assert.IsType<JsonObject>(projectedTheme["nativeRef"]);
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeColorRoleDark2" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "colorRoles.dark2"));
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeColorRoleDark1" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "colorRoles.dark1"));
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeColorRoleLight1" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "colorRoles.light1"));

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
        editedProgram["design"]!["theme"]!["colorRoles"]!["dark2"] = "#ABCDEF";
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
            Assert.Equal("101112", scheme.Dark1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("202122", scheme.Light1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("ABCDEF", scheme.Dark2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("404142", scheme.Light2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("505152", scheme.Hyperlink!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("606162", scheme.FollowedHyperlinkColor!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("112233", scheme.Accent1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("223344", scheme.Accent2Color!.RgbColorModelHex!.Val!.Value);
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
                SourceUri = "deck.assets/edited/theme-color-role-dark2.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal("#ABCDEF", reprojectedProgram["design"]!["theme"]!["colorRoles"]!["dark2"]!.GetValue<string>());

        var unauthorizedProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        unauthorizedProgram["design"]!["theme"]!["colorRoles"]!.AsObject().Remove("dark2");
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

        var unownedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        unownedProgram["design"]!["theme"]!["colorRoles"]!["light2"] = "#202122";
        var unowned = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(unownedProgram.ToJsonString()),
            },
        });
        Assert.False(unowned.Ok, Diagnostics(unowned));

        var tamperedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        tamperedProgram["design"]!["theme"]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(capability => capability!["operation"]!.GetValue<string>() == "setThemeColorRoleDark2")!["fields"] = new JsonArray("colorRoles.light2");
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
