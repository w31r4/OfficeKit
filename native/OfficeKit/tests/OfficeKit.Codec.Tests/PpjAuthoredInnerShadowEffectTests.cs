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
    public void PpjAuthoredInnerShadowEffectWritesShapeAndPictureOwners()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();

        var shapeStyle = program["design"]!["styles"]!["shape"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "decision-band")!["style"]!.AsObject();
        shapeStyle["shadow"] = new JsonObject
        {
            ["color"] = "#17324D",
            ["opacity"] = 0.24,
            ["blur"] = 6,
            ["distance"] = 3,
            ["angle"] = 45,
        };
        shapeStyle["glow"] = new JsonObject
        {
            ["color"] = "#D9A514",
            ["radius"] = 12,
            ["opacity"] = 0.42,
        };
        shapeStyle["innerShadow"] = new JsonObject
        {
            ["color"] = "#AABBCC",
            ["blur"] = 5,
            ["distance"] = 2,
            ["angle"] = 45,
            ["opacity"] = 0.36,
        };
        shapeStyle["softEdge"] = new JsonObject { ["radius"] = 8 };

        program["design"]!["styles"]!["image"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "evidence-effects",
                ["style"] = new JsonObject
                {
                    ["glow"] = new JsonObject
                    {
                        ["color"] = "#445566",
                        ["radius"] = 6,
                        ["opacity"] = 0.27,
                    },
                    ["innerShadow"] = new JsonObject
                    {
                        ["color"] = "#667788",
                        ["blur"] = 3,
                        ["distance"] = 1,
                        ["angle"] = 90,
                        ["opacity"] = 0.18,
                    },
                    ["softEdge"] = new JsonObject { ["radius"] = 4 },
                },
            },
        };
        var image = program["pages"]!.AsArray().SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["id"]!.GetValue<string>() == "claim-mark");
        image["styleRef"] = "evidence-effects";

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
            var shape = slideParts.SelectMany(part => part.Slide!.Descendants<P.Shape>()).Single(item =>
                item.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "decision rule");
            var shapeEffects = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.Glow>(shapeEffects.ChildElements[0]);
            var shapeInnerShadow = Assert.IsType<A.InnerShadow>(shapeEffects.ChildElements[1]);
            Assert.Equal(63_500, shapeInnerShadow.BlurRadius!.Value);
            Assert.Equal(25_400, shapeInnerShadow.Distance!.Value);
            Assert.Equal(2_700_000, shapeInnerShadow.Direction!.Value);
            Assert.Equal("AABBCC", shapeInnerShadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(36_000, shapeInnerShadow.Descendants<A.Alpha>().Single().Val!.Value);
            Assert.IsType<A.OuterShadow>(shapeEffects.ChildElements[2]);
            Assert.IsType<A.SoftEdge>(shapeEffects.ChildElements[3]);

            var picture = slideParts.SelectMany(part => part.Slide!.Descendants<P.Picture>()).Single(item =>
                item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var pictureEffects = picture.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.IsType<A.Glow>(pictureEffects.ChildElements[0]);
            var pictureInnerShadow = Assert.IsType<A.InnerShadow>(pictureEffects.ChildElements[1]);
            Assert.Equal(38_100, pictureInnerShadow.BlurRadius!.Value);
            Assert.Equal(12_700, pictureInnerShadow.Distance!.Value);
            Assert.Equal(5_400_000, pictureInnerShadow.Direction!.Value);
            Assert.IsType<A.OuterShadow>(pictureEffects.ChildElements[2]);
            Assert.IsType<A.SoftEdge>(pictureEffects.ChildElements[3]);
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
        var recoveredShapeStyle = recoveredProgram["design"]!["styles"]!["shape"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "decision-band")!["style"]!;
        Assert.Equal(5, recoveredShapeStyle!["innerShadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(0.36, recoveredShapeStyle["innerShadow"]!["opacity"]!.GetValue<double>());
        var recoveredImageStyle = recoveredProgram["design"]!["styles"]!["image"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "evidence-effects")!["style"]!;
        Assert.Equal(3, recoveredImageStyle!["innerShadow"]!["blur"]!.GetValue<double>());

        var invalidProgram = program.DeepClone().AsObject();
        invalidProgram["design"]!["styles"]!["shape"]!.AsArray()
            .Single(item => item!["id"]!.GetValue<string>() == "decision-band")!["style"]!["innerShadow"]!["blur"] = -1;
        var invalid = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidProgram.ToJsonString()));
        Assert.False(invalid.IsValid);
    }
}
