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
    public void PpjAuthoredThemeAccentAlphaTransformsRetainAbsoluteAlphaAndRecover()
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
            ["accent3"] = "#30405080",
            ["accent4"] = "#405060",
            ["accent5"] = "#506070",
            ["accent6"] = "#607080",
        };
        program["design"]!["theme"]!["accentTransforms"] = new JsonObject
        {
            ["accent3"] = new JsonObject
            {
                ["tint"] = 0.1,
                ["shade"] = 0.2,
                ["lumMod"] = 0.8,
                ["lumOff"] = -0.05,
                ["satMod"] = 0.7,
                ["satOff"] = -0.2,
                ["alphaMod"] = 0.6,
                ["alphaOff"] = -0.1,
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
            var accent3 = scheme.Accent3Color!.RgbColorModelHex!;
            Assert.Equal("304050", accent3.Val!.Value);
            Assert.Equal(10_000, Assert.Single(accent3.Elements<A.Tint>()).Val!.Value);
            Assert.Equal(20_000, Assert.Single(accent3.Elements<A.Shade>()).Val!.Value);
            Assert.Equal(80_000, Assert.Single(accent3.Elements<A.LuminanceModulation>()).Val!.Value);
            Assert.Equal(-5_000, Assert.Single(accent3.Elements<A.LuminanceOffset>()).Val!.Value);
            Assert.Equal(70_000, Assert.Single(accent3.Elements<A.SaturationModulation>()).Val!.Value);
            Assert.Equal(-20_000, Assert.Single(accent3.Elements<A.SaturationOffset>()).Val!.Value);
            Assert.Equal(50_196, Assert.Single(accent3.Elements<A.Alpha>()).Val!.Value);
            Assert.Equal(60_000, Assert.Single(accent3.Elements<A.AlphaModulation>()).Val!.Value);
            Assert.Equal(-10_000, Assert.Single(accent3.Elements<A.AlphaOffset>()).Val!.Value);
            Assert.Empty(scheme.Accent2Color!.RgbColorModelHex!.Elements<A.AlphaModulation>());
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
        var recoveredTransform = recoveredProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!;
        Assert.Equal(0.6, recoveredTransform!["alphaMod"]!.GetValue<double>());
        Assert.Equal(-0.1, recoveredTransform["alphaOff"]!.GetValue<double>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["design"]!["theme"]!["accentTransforms"]!["accent3"]!["alphaOff"] = 1.2;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
