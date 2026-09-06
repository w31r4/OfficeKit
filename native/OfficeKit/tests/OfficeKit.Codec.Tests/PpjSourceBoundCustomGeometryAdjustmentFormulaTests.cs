using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjSourceBoundCustomGeometryAdjustmentFormulaEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var shape = authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape;
        shape.Geometry = "custom";
        shape.CustomAdjustments.Add(new PresentationCustomGeometryGuide { Name = "adj", Formula = "*/ w 1 2" });
        shape.CustomGuides.Add(new PresentationCustomGeometryGuide { Name = "x1", Formula = "*/ w adj 100000" });
        var path = new PresentationCustomGeometryPath
        {
            Width = shape.WidthEmu,
            Height = shape.HeightEmu,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
        };
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            MoveTo = new PresentationCustomGeometryPoint { XReference = "l", YReference = "t" },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            LineTo = new PresentationCustomGeometryPoint { XReference = "x1", YReference = "b" },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        shape.CustomPaths.Add(path);

        var authored = Invoke(authoredRequest);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        long authoredTextLength;
        using (var stream = new MemoryStream(source, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var geometry = nativeShape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!;
            var adjustment = Assert.Single(geometry.GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>());
            Assert.Equal("*/ w 1 2", adjustment.Formula!.Value);
            authoredTextLength = nativeShape.TextBody!.InnerText.Length;
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-geometry-adjustment-formula.pptx",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var projectedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = projectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var formulaLeaf = Assert.Single(projectedShape["nativeRef"]!["leaves"]!.AsArray()
            .Select(leaf => leaf!.AsObject()),
            leaf => leaf["kind"]!.GetValue<string>() == "customGeometryAdjustmentFormula");
        Assert.Equal("*/ w 1 2", formulaLeaf["value"]!.GetValue<string>());
        Assert.Contains(projectedShape["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "customGeometryGuideFormula" &&
            leaf["value"]!.GetValue<string>() == "*/ w adj 100000");

        formulaLeaf["value"] = "*/ w 1 3";
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
        using (var stream = new MemoryStream(editedBytes, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var guides = nativeShape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!
                .GetFirstChild<A.AdjustValueList>()!.Elements<A.ShapeGuide>().ToArray();
            Assert.Equal("*/ w 1 3", guides[0].Formula!.Value);
            Assert.Equal(authoredTextLength, nativeShape.TextBody!.InnerText.Length);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        foreach (var pathName in ZipPartPaths(source).Where(pathName => !pathName.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(ZipBytes(source, pathName), ZipBytes(editedBytes, pathName));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/edited/custom-geometry-adjustment-formula.pptx",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var reprojectedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedShape = reprojectedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal("*/ w 1 3", reprojectedShape["nativeRef"]!["leaves"]!.AsArray().Single(leaf =>
            leaf!["kind"]!.GetValue<string>() == "customGeometryAdjustmentFormula")!["value"]!.GetValue<string>());
    }
}
