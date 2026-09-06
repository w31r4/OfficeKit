using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjAuthoredTextInnerShadowEffectWritesRunAndDefaultRunOwners()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();

        var claimTitle = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-title");
        var runStyle = claimTitle["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!.AsObject();
        runStyle["glow"] = new JsonObject
        {
            ["color"] = "#D9A514",
            ["radius"] = 8,
            ["opacity"] = 0.42,
        };
        runStyle["innerShadow"] = new JsonObject
        {
            ["color"] = "#AABBCC",
            ["blur"] = 5,
            ["distance"] = 2,
            ["angle"] = 45,
            ["opacity"] = 0.36,
        };
        runStyle["shadow"] = new JsonObject
        {
            ["color"] = "#16324F",
            ["blur"] = 2,
            ["distance"] = 1,
            ["angle"] = 90,
        };
        claimTitle["text"]!["paragraphs"]![0]!["style"] = new JsonObject
        {
            ["defaultText"] = new JsonObject
            {
                ["innerShadow"] = new JsonObject
                {
                    ["color"] = "#0B8F8F",
                    ["blur"] = 4,
                    ["distance"] = 1,
                    ["angle"] = 180,
                    ["opacity"] = 0.27,
                },
            },
        };

        var programBytes = Encoding.UTF8.GetBytes(program.ToJsonString());
        var validation = PpjProgramValidator.Validate(programBytes);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics));
        var assetBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
        var assetSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(assetBytes)).ToLowerInvariant();
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFrom(programBytes),
                Assets =
                {
                    new Asset
                    {
                        Id = "evidence-mark",
                        FileName = "evidence-mark.svg",
                        ContentType = "image/svg+xml",
                        Data = ByteString.CopyFrom(assetBytes),
                        Sha256 = assetSha256,
                    },
                },
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));

        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var slideParts = package.PresentationPart!.SlideParts.ToArray();
            var titleShape = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.TextBody?.Descendants<A.Text>().Any(text => text.Text == "Reduce incident hours ") == true);
            var run = titleShape.TextBody!.Descendants<A.Run>().Single(item =>
                item.GetFirstChild<A.Text>()?.Text == "Reduce incident hours ");
            var runEffects = run.RunProperties!.GetFirstChild<A.EffectList>()!;
            Assert.Collection(
                runEffects.ChildElements,
                glow => Assert.IsType<A.Glow>(glow),
                innerShadow =>
                {
                    var native = Assert.IsType<A.InnerShadow>(innerShadow);
                    Assert.Equal(63_500, native.BlurRadius!.Value);
                    Assert.Equal(25_400, native.Distance!.Value);
                    Assert.Equal(2_700_000, native.Direction!.Value);
                    Assert.Equal("AABBCC", native.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                    Assert.Equal(36_000, native.Descendants<A.Alpha>().Single().Val!.Value);
                },
                shadow => Assert.IsType<A.OuterShadow>(shadow));

            var defaultRunProperties = titleShape.TextBody!.Elements<A.Paragraph>().Single()
                .ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            var defaultEffects = defaultRunProperties.GetFirstChild<A.EffectList>()!;
            var defaultInnerShadow = Assert.IsType<A.InnerShadow>(Assert.Single(defaultEffects.ChildElements));
            Assert.Equal(50_800, defaultInnerShadow.BlurRadius!.Value);
            Assert.Equal(12_700, defaultInnerShadow.Distance!.Value);
            Assert.Equal(10_800_000, defaultInnerShadow.Direction!.Value);
            Assert.Equal("0B8F8F", defaultInnerShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(27_000, defaultInnerShadow.Descendants<A.Alpha>().Single().Val!.Value);
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
        var recoveredRunInnerShadow = recoveredProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-title")["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["innerShadow"]!;
        Assert.Equal(5, recoveredRunInnerShadow["blur"]!.GetValue<double>());
        Assert.Equal(2, recoveredRunInnerShadow["distance"]!.GetValue<double>());
        Assert.Equal(45, recoveredRunInnerShadow["angle"]!.GetValue<double>());
        Assert.Equal(0.36, recoveredRunInnerShadow["opacity"]!.GetValue<double>());
        var recoveredDefaultInnerShadow = recoveredProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-title")["text"]!["paragraphs"]![0]!["style"]!["defaultText"]!["innerShadow"]!;
        Assert.Equal(4, recoveredDefaultInnerShadow["blur"]!.GetValue<double>());
        Assert.Equal(180, recoveredDefaultInnerShadow["angle"]!.GetValue<double>());
        Assert.Equal(0.27, recoveredDefaultInnerShadow["opacity"]!.GetValue<double>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-title")["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["innerShadow"]!["blur"] = -1;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
