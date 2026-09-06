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
    public void PpjAuthoredThemeAccentChannelTransformsWriteAndRecover()
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
            ["accent1"] = new JsonObject
            {
                ["redMod"] = 0.8,
                ["blueOff"] = -0.2,
            },
            ["accent4"] = new JsonObject
            {
                ["redMod"] = 0.8,
                ["redOff"] = -0.05,
                ["greenMod"] = 0.6,
                ["greenOff"] = 0.1,
                ["blueMod"] = 0.7,
                ["blueOff"] = -0.2,
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

            var accent1 = scheme.Accent1Color!.RgbColorModelHex!;
            Assert.Equal("102030", accent1.Val!.Value);
            Assert.Equal(80_000, Assert.Single(accent1.Elements<A.RedModulation>()).Val!.Value);
            Assert.Equal(-20_000, Assert.Single(accent1.Elements<A.BlueOffset>()).Val!.Value);
            Assert.Empty(accent1.Elements<A.GreenModulation>());
            Assert.Empty(accent1.Elements<A.GreenOffset>());

            var accent4 = scheme.Accent4Color!.RgbColorModelHex!;
            Assert.Equal("405060", accent4.Val!.Value);
            Assert.Equal(80_000, Assert.Single(accent4.Elements<A.RedModulation>()).Val!.Value);
            Assert.Equal(-5_000, Assert.Single(accent4.Elements<A.RedOffset>()).Val!.Value);
            Assert.Equal(60_000, Assert.Single(accent4.Elements<A.GreenModulation>()).Val!.Value);
            Assert.Equal(10_000, Assert.Single(accent4.Elements<A.GreenOffset>()).Val!.Value);
            Assert.Equal(70_000, Assert.Single(accent4.Elements<A.BlueModulation>()).Val!.Value);
            Assert.Equal(-20_000, Assert.Single(accent4.Elements<A.BlueOffset>()).Val!.Value);
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
        var recoveredAccent1 = recoveredProgram["design"]!["theme"]!["accentTransforms"]!["accent1"]!;
        Assert.Equal(0.8, recoveredAccent1!["redMod"]!.GetValue<double>());
        Assert.Equal(-0.2, recoveredAccent1["blueOff"]!.GetValue<double>());
        Assert.Null(recoveredAccent1["greenMod"]);

        var recoveredAccent4 = recoveredProgram["design"]!["theme"]!["accentTransforms"]!["accent4"]!;
        Assert.Equal(0.8, recoveredAccent4!["redMod"]!.GetValue<double>());
        Assert.Equal(-0.05, recoveredAccent4["redOff"]!.GetValue<double>());
        Assert.Equal(0.6, recoveredAccent4["greenMod"]!.GetValue<double>());
        Assert.Equal(0.1, recoveredAccent4["greenOff"]!.GetValue<double>());
        Assert.Equal(0.7, recoveredAccent4["blueMod"]!.GetValue<double>());
        Assert.Equal(-0.2, recoveredAccent4["blueOff"]!.GetValue<double>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["design"]!["theme"]!["accentTransforms"]!["accent4"]!["greenOff"] = 1.2;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
