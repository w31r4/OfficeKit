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
    public void PpjAuthoredThemeAccentDiscreteTransformsWriteAndRecover()
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
                ["gray"] = true,
            },
            ["accent5"] = new JsonObject
            {
                ["gray"] = true,
                ["comp"] = true,
                ["inv"] = true,
                ["gamma"] = true,
                ["invGamma"] = true,
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
            Assert.Single(accent1.Elements<A.Gray>());
            Assert.Empty(accent1.Elements<A.Complement>());
            Assert.Empty(accent1.Elements<A.Inverse>());
            Assert.Empty(accent1.Elements<A.Gamma>());
            Assert.Empty(accent1.Elements<A.InverseGamma>());

            var accent5 = scheme.Accent5Color!.RgbColorModelHex!;
            Assert.Equal("506070", accent5.Val!.Value);
            Assert.Single(accent5.Elements<A.Gray>());
            Assert.Single(accent5.Elements<A.Complement>());
            Assert.Single(accent5.Elements<A.Inverse>());
            Assert.Single(accent5.Elements<A.Gamma>());
            Assert.Single(accent5.Elements<A.InverseGamma>());
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
        Assert.True(recoveredAccent1!["gray"]!.GetValue<bool>());
        Assert.Null(recoveredAccent1["comp"]);

        var recoveredAccent5 = recoveredProgram["design"]!["theme"]!["accentTransforms"]!["accent5"]!;
        Assert.True(recoveredAccent5!["gray"]!.GetValue<bool>());
        Assert.True(recoveredAccent5["comp"]!.GetValue<bool>());
        Assert.True(recoveredAccent5["inv"]!.GetValue<bool>());
        Assert.True(recoveredAccent5["gamma"]!.GetValue<bool>());
        Assert.True(recoveredAccent5["invGamma"]!.GetValue<bool>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["design"]!["theme"]!["accentTransforms"]!["accent5"]!["inv"] = false;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
