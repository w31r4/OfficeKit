using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Security.Cryptography;
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
    public void PpjSourceBoundPictureShape3dBevelBottomWidthLeafEditsAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        if (root is null)
        {
            root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
                root = root.Parent;
        }
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            fixtureDirectory,
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var assetBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
        var assetSha256 = Convert.ToHexString(SHA256.HashData(assetBytes)).ToLowerInvariant();
        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
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

        var source = AddPictureShape3dBevelBottomWidth(RemoveEmbeddedPpj(authored.File.ToByteArray()), 1_200);
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/picture-shape-3d-bevel-bottom-width.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));

        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedImage = projectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        var target = Assert.Single(projectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Where(leaf => leaf["kind"]!.GetValue<string>() == "shape3dBevelBottomWidthEmu"));
        Assert.Equal(1_200, target["value"]!.GetValue<long>());
        target["value"] = 2_400;

        var edited = Invoke(new CodecRequest
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
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);

        var editedBytes = edited.File.ToByteArray();
        var editedXml = Encoding.UTF8.GetString(ZipBytes(editedBytes, "ppt/slides/slide1.xml"));
        Assert.Contains("<a:bevelB", editedXml, StringComparison.Ordinal);
        Assert.Contains("w=\"2400\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("h=\"800\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("prst=\"angle\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("z=\"1200\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("extrusionH=\"2400\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("prstMaterial=\"metal\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        foreach (var pathName in ZipPartPaths(source).Where(pathName => !pathName.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, pathName), ZipBytes(editedBytes, pathName));

        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var picture = package.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>())
                .Single(item => item.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity");
            var bevel = picture.ShapeProperties!.GetFirstChild<A.Shape3DType>()!.GetFirstChild<A.BevelBottom>();
            Assert.Equal("2400", bevel!.GetAttributes().Single(attribute => attribute.LocalName == "w").Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(editedBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/picture-shape-3d-bevel-bottom-width.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedImage = reprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(item => item!.AsObject())
            .Single(item => item["name"]!.GetValue<string>() == "evidence identity");
        Assert.Equal(2_400, reprojectedImage["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject())
            .Single(leaf => leaf["kind"]!.GetValue<string>() == "shape3dBevelBottomWidthEmu")["value"]!.GetValue<long>());
    }

    private static byte[] AddPictureShape3dBevelBottomWidth(byte[] source, long width)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true))
        {
            var matches = package.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                    .Select(picture => (slidePart, picture)))
                .Where(item => item.picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "evidence identity")
                .ToArray();
            Assert.Single(matches);
            var slidePart = matches[0].slidePart;
            var picture = matches[0].picture;
            var bevel = new A.BevelBottom();
            bevel.SetAttribute(new OpenXmlAttribute("w", string.Empty, width.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            bevel.SetAttribute(new OpenXmlAttribute("h", string.Empty, "800"));
            bevel.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "angle"));
            var shape3d = new A.Shape3DType(bevel);
            shape3d.SetAttribute(new OpenXmlAttribute("z", string.Empty, "1200"));
            shape3d.SetAttribute(new OpenXmlAttribute("extrusionH", string.Empty, "2400"));
            shape3d.SetAttribute(new OpenXmlAttribute("prstMaterial", string.Empty, "metal"));
            picture.ShapeProperties!.AppendChild(shape3d);
            slidePart.Slide.Save();
        }
        return stream.ToArray();
    }
}
