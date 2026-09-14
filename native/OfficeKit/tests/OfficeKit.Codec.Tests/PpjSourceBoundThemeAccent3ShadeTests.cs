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
    public void PpjSourceBoundThemeAccent3ShadeEditsOnlyShadeAndReprojects()
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
        authoredTheme.AccentTransforms.Add(new PresentationThemeColorTransform
        {
            Role = "accent3",
            ShadeThousandth = 25_000,
        });
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
                SourceUri = "deck.assets/source/theme-accent3-shade.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTheme = projectedProgram["design"]!["theme"]!.AsObject();
        Assert.False(projectedTheme.ContainsKey("accentColors"));
        var projectedTransform = projectedTheme["accentTransforms"]!["accent3"]!;
        Assert.Equal(0.25, projectedTransform["shade"]!.GetValue<double>());
        var nativeRef = Assert.IsType<JsonObject>(projectedTheme["nativeRef"]);
        Assert.Contains(nativeRef["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setThemeAccent3Shade" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "accentTransforms.accent3.shade"));

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
        editedProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["shade"] = 0.7;
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
            Assert.Equal("223344", scheme.Accent2Color!.RgbColorModelHex!.Val!.Value);
            var accent3 = scheme.Accent3Color!.RgbColorModelHex!;
            Assert.Equal("334455", accent3.Val!.Value);
            Assert.Equal(70_000, Assert.Single(accent3.Elements<A.Shade>()).Val!.Value);
            Assert.Equal("445566", scheme.Accent4Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("556677", scheme.Accent5Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("667788", scheme.Accent6Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("101112", scheme.Dark1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("202122", scheme.Light1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("303132", scheme.Dark2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("404142", scheme.Light2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("505152", scheme.Hyperlink!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("606162", scheme.FollowedHyperlinkColor!.RgbColorModelHex!.Val!.Value);
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
                SourceUri = "deck.assets/edited/theme-accent3-shade.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal(0.7, reprojectedProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["shade"]!.GetValue<double>());

        var unauthorizedProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        unauthorizedProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["shade"] = 0.8;
        unauthorizedProgram["design"]!["theme"]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(capability => capability!["operation"]!.GetValue<string>() == "setThemeAccent3Shade")!["fields"] = new JsonArray("accentTransforms.accent3.tint");
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

        var deletionProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        deletionProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!.AsObject().Remove("shade");
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

        var siblingProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        siblingProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["tint"] = 0.1;
        var sibling = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(siblingProgram.ToJsonString()),
            },
        });
        Assert.False(sibling.Ok, Diagnostics(sibling));

        var combinedProgram = JsonNode.Parse(editedProgram.ToJsonString())!.AsObject();
        combinedProgram["design"]!["theme"]!["colorRoles"]!["hyperlink"] = "#202122";
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

        var outOfRangeProgram = JsonNode.Parse(projectedProgram.ToJsonString())!.AsObject();
        outOfRangeProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["shade"] = 1.01;
        var outOfRange = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(outOfRangeProgram.ToJsonString()),
            },
        });
        Assert.False(outOfRange.Ok, Diagnostics(outOfRange));
    }
}
