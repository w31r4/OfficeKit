using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
    public void PpjAuthoredThemeAccentLuminanceTransformsWriteAndRecover()
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
        program["design"]!["theme"]!["accentColors"] = new JsonObject
        {
            ["accent1"] = "#102030",
            ["accent2"] = "#203040",
            ["accent3"] = "#304050",
            ["accent4"] = "#405060",
            ["accent5"] = "#506070",
            ["accent6"] = "#607080",
        };
        program["design"]!["theme"]!["accentTransforms"] = new JsonObject
        {
            ["accent2"] = new JsonObject
            {
                ["lumMod"] = 0.8,
                ["lumOff"] = -0.05,
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
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var scheme = Assert.Single(package.PresentationPart!.SlideMasterParts).ThemePart!.Theme!
                .ThemeElements!.ColorScheme!;
            var accent2 = scheme.Accent2Color!.RgbColorModelHex!;
            Assert.Equal("203040", accent2.Val!.Value);
            Assert.Equal(80_000, Assert.Single(accent2.Elements<A.LuminanceModulation>()).Val!.Value);
            Assert.Equal(-5_000, Assert.Single(accent2.Elements<A.LuminanceOffset>()).Val!.Value);
            Assert.Empty(scheme.Accent1Color!.RgbColorModelHex!.Elements<A.LuminanceModulation>());
            Assert.Empty(scheme.Accent1Color.RgbColorModelHex.Elements<A.LuminanceOffset>());
        }

        var recovered = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = authored.File,
            PresentationProgram = new PresentationProgramRequest(),
        });
        Assert.True(recovered.Ok, Diagnostics(recovered));
        Assert.True(recovered.PresentationProgram.RestoredEmbeddedProgram);
        var recoveredProgram = JsonNode.Parse(recovered.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var recoveredTransform = recoveredProgram["design"]!["theme"]!["accentTransforms"]!["accent2"]!;
        Assert.Equal(0.8, recoveredTransform!["lumMod"]!.GetValue<double>());
        Assert.Equal(-0.05, recoveredTransform["lumOff"]!.GetValue<double>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["design"]!["theme"]!["accentTransforms"]!["accent2"]!["lumOff"] = 1.2;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
