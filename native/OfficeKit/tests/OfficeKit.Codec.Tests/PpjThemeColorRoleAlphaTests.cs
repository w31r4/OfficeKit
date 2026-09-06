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
    public void PpjThemeColorRoleAlphaAuthorsNativeTransformsAndPreservesLegacyRgb()
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
            ["accent1"] = "#10203080",
            ["accent2"] = "#203040",
            ["accent3"] = "#304050",
            ["accent4"] = "#405060",
            ["accent5"] = "#506070",
            ["accent6"] = "#607080",
        };
        program["design"]!["theme"]!["colorRoles"] = new JsonObject
        {
            ["dark1"] = "#10111240",
            ["light1"] = "#202122",
            ["dark2"] = "#303132",
            ["light2"] = "#404142",
            ["hyperlink"] = "#505152",
            ["followedHyperlink"] = "#606162",
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
        Assert.Equal(programBytes, ZipBytes(authored.File.ToByteArray(), "officeKit/program.ppj"));

        using var stream = new MemoryStream(authored.File.ToByteArray(), writable: false);
        using var package = PresentationDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        var colorScheme = Assert.Single(package.PresentationPart!.SlideMasterParts).ThemePart!.Theme!
            .ThemeElements!.ColorScheme!;
        Assert.Equal("102030", colorScheme.Accent1Color!.RgbColorModelHex!.Val!.Value);
        Assert.Equal(50_196, colorScheme.Accent1Color.RgbColorModelHex.GetFirstChild<A.Alpha>()!.Val!.Value);
        Assert.Equal("203040", colorScheme.Accent2Color!.RgbColorModelHex!.Val!.Value);
        Assert.Null(colorScheme.Accent2Color.RgbColorModelHex.GetFirstChild<A.Alpha>());
        Assert.Equal("101112", colorScheme.Dark1Color!.RgbColorModelHex!.Val!.Value);
        Assert.Equal(25_098, colorScheme.Dark1Color.RgbColorModelHex.GetFirstChild<A.Alpha>()!.Val!.Value);
    }
}
