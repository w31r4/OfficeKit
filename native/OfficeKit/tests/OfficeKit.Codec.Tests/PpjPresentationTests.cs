using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using AD = DocumentFormat.OpenXml.Office2019.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using OD = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Fact]
    public void PpjV1ValidatesAndExpandsCanonicalPresentationProgram()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var fixturePath = Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj");
        var bytes = File.ReadAllBytes(fixturePath);
        var result = PpjProgramValidator.Validate(bytes);
        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics));
        Assert.NotNull(result.Expansion);
        Assert.Equal(16, result.Expansion.ExpandedElementCount);
        Assert.Contains(result.Expansion.Nodes, node => node.Id == "evidence-rows::evidence-row::hours::row-label");
        Assert.Contains(result.Expansion.Nodes, node => node.Id == "evidence-rows::evidence-row::workload::row-value");
        Assert.Contains(result.Expansion.Pages.SelectMany(page => page.Elements), element =>
            element.Id == "evidence-rows::evidence-row::hours::row-label");

        var repeated = PpjProgramValidator.Validate(bytes);
        Assert.Equal(result.ProgramSha256, repeated.ProgramSha256);
        Assert.Equal(result.Expansion.NodeMapSha256, repeated.Expansion!.NodeMapSha256);

        var widePlaceholderIndex = JsonNode.Parse(bytes)!.AsObject();
        widePlaceholderIndex["pages"]!.AsArray()[0]!.AsObject()["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "source-placeholder",
            ["type"] = "placeholder",
            ["frame"] = new JsonObject { ["x"] = 0, ["y"] = 0, ["width"] = 100, ["height"] = 40 },
            ["placeholderType"] = "other",
            ["index"] = uint.MaxValue,
        });
        var widePlaceholderResult = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(widePlaceholderIndex.ToJsonString()));
        Assert.True(widePlaceholderResult.IsValid, string.Join(Environment.NewLine, widePlaceholderResult.Diagnostics));

        var unknownField = JsonNode.Parse(bytes)!.AsObject();
        unknownField["pages"]!.AsArray()[0]!.AsObject()["elements"]!.AsArray()[0]!.AsObject()["rawOoxml"] = "forbidden";
        var rejectedField = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(unknownField.ToJsonString()));
        Assert.False(rejectedField.IsValid);
        Assert.Contains(rejectedField.Diagnostics, diagnostic =>
            diagnostic.Code == "ppj.schema.unknownField" && diagnostic.Path == "$.pages[0].elements[0].rawOoxml");

        var recursive = JsonNode.Parse(bytes)!.AsObject();
        recursive["components"]!.AsArray()[0]!.AsObject()["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "self",
            ["type"] = "component",
            ["frame"] = new JsonObject { ["x"] = 0, ["y"] = 0, ["width"] = 1, ["height"] = 1 },
            ["component"] = "evidence-row",
        });
        var rejectedCycle = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(recursive.ToJsonString()));
        Assert.False(rejectedCycle.IsValid);
        Assert.Contains(rejectedCycle.Diagnostics, diagnostic => diagnostic.Code == "ppj.component.cycle");
    }

    [Fact]
    public void PpjSourceBoundAccessibilityEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape.Accessibility =
            new PresentationNonVisualAccessibility
            {
                Title = "Before review",
                Description = "Initial decision summary.",
                Decorative = false,
            };
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
                SourceUri = "deck.assets/source/accessibility.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = program["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal("Before review", element["accessibility"]!["title"]!.GetValue<string>());
        Assert.False(element["accessibility"]!["decorative"]!.GetValue<bool>());
        Assert.Contains(element["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setAccessibility" &&
            capability["fields"]!.AsArray().Select(field => field!.GetValue<string>()).SequenceEqual(["accessibility"]));

        var tampered = JsonNode.Parse(program.ToJsonString())!.AsObject();
        var tamperedCapabilities = tampered["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray();
        tamperedCapabilities.Remove(tamperedCapabilities.Single(capability =>
            capability!["operation"]!.GetValue<string>() == "setAccessibility"));
        tampered["pages"]![0]!["elements"]![0]!["accessibility"]!["title"] = "Tampered capability";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(tampered.ToJsonString()),
            },
        });
        Assert.False(rejected.Ok);
        Assert.Equal("ppj.nativeRef.stale", Assert.Single(rejected.Diagnostics).Code);

        element["accessibility"]!["title"] = "After review";
        element["accessibility"]!["description"] = "Updated decision summary.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nonVisual = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single()
                .NonVisualShapeProperties!.NonVisualDrawingProperties!;
            Assert.Equal("After review", nonVisual.Title!.Value);
            Assert.Equal("Updated decision summary.", nonVisual.Description!.Value);
            Assert.False(Assert.Single(nonVisual.Descendants<AD.Decorative>()).Val!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/accessibility-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var roundTrip = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var roundTripAccessibility = roundTrip["pages"]![0]!["elements"]!.AsArray().Single()!["accessibility"]!;
        Assert.Equal("After review", roundTripAccessibility!["title"]!.GetValue<string>());
        Assert.Equal("Updated decision summary.", roundTripAccessibility!["description"]!.GetValue<string>());
        Assert.False(roundTripAccessibility!["decorative"]!.GetValue<bool>());
    }

    [Fact]
    public void PpjSourceBoundMasterAndLayoutBackgroundEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var masterTextStyles = new PresentationMasterTextStyles();
        masterTextStyles.TitleLevels.Add(new PresentationTextParagraph
        {
            Level = 0,
            Alignment = "center",
            DefaultRunProperties = new PresentationTextStyle { FontSizePoints = 28 },
        });
        authoredRequest.Artifact!.Presentation!.Masters.Add(new PresentationMaster
        {
            Id = "master/source-background",
            Name = "Source background master",
            Background = new PresentationBackground { ColorRgb = "EAF2F8", Solid = true },
            TextStyles = masterTextStyles,
        });
        authoredRequest.Artifact.Presentation.Layouts.Add(new PresentationLayout
        {
            Id = "layout/source-background",
            Name = "Source background layout",
            MasterId = "master/source-background",
            Type = "blank",
            Background = new PresentationBackground { ColorRgb = "FDF2E9", Solid = true },
        });
        authoredRequest.Artifact.Presentation.Slides[0].LayoutId = "layout/source-background";
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
                SourceUri = "deck.assets/source/master-background.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var master = program["design"]!["masters"]!.AsArray().Single()!.AsObject();
        var layout = program["design"]!["layouts"]!.AsArray().Single()!.AsObject();
        Assert.Contains(master["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setBackground" &&
            capability["fields"]!.AsArray().Select(field => field!.GetValue<string>()).SequenceEqual(["background"]));
        Assert.Contains(master["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setTextParagraphStyle" &&
            capability["fields"]!.AsArray().Select(field => field!.GetValue<string>()).SequenceEqual(["textStyles"]));
        Assert.Contains(layout["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setBackground" &&
            capability["fields"]!.AsArray().Select(field => field!.GetValue<string>()).SequenceEqual(["background"]));

        var tampered = JsonNode.Parse(program.ToJsonString())!.AsObject();
        tampered["design"]!["masters"]!.AsArray().Single()!["nativeRef"]!["capabilities"]!.AsArray()
            .RemoveAt(0);
        tampered["design"]!["masters"]!.AsArray().Single()!["background"]!["color"] = "#112233";
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(tampered.ToJsonString()),
            },
        });
        Assert.False(rejected.Ok);
        Assert.Equal("ppj.nativeRef.stale", Assert.Single(rejected.Diagnostics).Code);

        master["background"]!["color"] = "#112233";
        master["textStyles"]!["title"]![0]!["alignment"] = "right";
        layout["background"]!["color"] = "#445566";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains("ppt/slideMasters/slideMaster1.xml", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideLayouts/slideLayout1.xml", edited.PresentationProgram.ChangedParts);
        Assert.DoesNotContain("ppt/slides/slide1.xml", edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var masterBackground = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .CommonSlideData!.Background!.Descendants<A.RgbColorModelHex>().Single();
            Assert.Equal("112233", masterBackground.Val!.Value);
            var layoutBackground = package.PresentationPart.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!
                .CommonSlideData!.Background!.Descendants<A.RgbColorModelHex>().Single();
            Assert.Equal("445566", layoutBackground.Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/master-background-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var roundTrip = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Equal("#112233", roundTrip["design"]!["masters"]!.AsArray().Single()!["background"]!["color"]!.GetValue<string>());
        Assert.Equal("right", roundTrip["design"]!["masters"]!.AsArray().Single()!["textStyles"]!["title"]![0]!["alignment"]!.GetValue<string>());
        Assert.Equal("#445566", roundTrip["design"]!["layouts"]!.AsArray().Single()!["background"]!["color"]!.GetValue<string>());
    }

    [Fact]
    public void PpjSourceBoundMasterAndLayoutImageBackgroundCropOpacityEditAndReprojects()
    {
        var request = ExportRequest();
        var masterBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        var layoutBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=");
        var masterAssetId = AddPictureAsset(request.Artifact, masterBytes, "image/png");
        var layoutAssetId = AddPictureAsset(request.Artifact, layoutBytes, "image/png");
        request.Artifact.Presentation.Masters.Add(new PresentationMaster
        {
            Id = "master/source-image-background",
            Name = "Source image background master",
            Background = new PresentationBackground
            {
                ImagePaint = new PresentationImagePaint
                {
                    AssetId = masterAssetId,
                    Mode = PresentationImagePaint.Types.Mode.Stretch,
                    Crop = new PresentationImageCrop { LeftThousandthPercent = 5_000, BottomThousandthPercent = 10_000 },
                    OpacityThousandthPercent = 88_000,
                },
            },
        });
        request.Artifact.Presentation.Layouts.Add(new PresentationLayout
        {
            Id = "layout/source-image-background",
            Name = "Source image background layout",
            MasterId = "master/source-image-background",
            Type = "blank",
            Background = new PresentationBackground
            {
                ImagePaint = new PresentationImagePaint
                {
                    AssetId = layoutAssetId,
                    Mode = PresentationImagePaint.Types.Mode.Stretch,
                    Crop = new PresentationImageCrop { TopThousandthPercent = 7_000, RightThousandthPercent = 12_000 },
                    OpacityThousandthPercent = 77_000,
                },
            },
        });
        request.Artifact.Presentation.Slides[0].LayoutId = "layout/source-image-background";
        var authored = Invoke(request);
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
                SourceUri = "master-layout-image-background/source.pptx",
                AssetRootUri = "master-layout-image-background/assets",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var master = program["design"]!["masters"]!.AsArray().Single()!.AsObject();
        var layout = program["design"]!["layouts"]!.AsArray().Single()!.AsObject();
        Assert.Equal("image", master["background"]!["type"]!.GetValue<string>());
        Assert.Equal("image", layout["background"]!["type"]!.GetValue<string>());
        Assert.Equal(0.88, master["background"]!["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.77, layout["background"]!["opacity"]!.GetValue<double>(), precision: 6);

        master["background"]!["crop"] = new JsonObject { ["left"] = 0.12, ["bottom"] = 0.02 };
        master["background"]!["opacity"] = 0.61;
        layout["background"]!["crop"] = new JsonObject { ["top"] = 0.03, ["right"] = 0.18 };
        layout["background"]!["opacity"] = 0.49;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains("ppt/slideMasters/slideMaster1.xml", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideLayouts/slideLayout1.xml", edited.PresentationProgram.ChangedParts);
        Assert.DoesNotContain(edited.PresentationProgram.ChangedParts, path => path.Contains("/media/", StringComparison.Ordinal));
        Assert.DoesNotContain(edited.PresentationProgram.ChangedParts, path => path.Contains("/_rels/", StringComparison.Ordinal));

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var masterFill = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>()!;
            var masterCrop = masterFill.GetFirstChild<A.SourceRectangle>()!;
            Assert.Equal(12_000, masterCrop.Left!.Value);
            Assert.Equal(2_000, masterCrop.Bottom!.Value);
            Assert.Equal(61_000, masterFill.GetFirstChild<A.Blip>()!.GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
            var layoutFill = package.PresentationPart.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!
                .CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>()!;
            var layoutCrop = layoutFill.GetFirstChild<A.SourceRectangle>()!;
            Assert.Equal(3_000, layoutCrop.Top!.Value);
            Assert.Equal(18_000, layoutCrop.Right!.Value);
            Assert.Equal(49_000, layoutFill.GetFirstChild<A.Blip>()!.GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "master-layout-image-background/edited.pptx",
                AssetRootUri = "master-layout-image-background/assets",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var output = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var outputMaster = output["design"]!["masters"]!.AsArray().Single()!["background"]!;
        var outputLayout = output["design"]!["layouts"]!.AsArray().Single()!["background"]!;
        Assert.Equal(0.12, outputMaster["crop"]!["left"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.02, outputMaster["crop"]!["bottom"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.61, outputMaster["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.03, outputLayout["crop"]!["top"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.18, outputLayout["crop"]!["right"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.49, outputLayout["opacity"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects()
    {
        var request = ExportRequest();
        var masterSourceBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMQ+A8AASIBEMsUe3wAAAAASUVORK5CYII=");
        var layoutSourceBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNQ+A8AAUIBIIyHQh0AAAAASUVORK5CYII=");
        var masterReplacementBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNw+A8AAYIBQKns11MAAAAASUVORK5CYII=");
        var layoutReplacementBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNo+A8AAgIBgG5WixMAAAAASUVORK5CYII=");
        var masterSourceAssetId = AddPictureAsset(request.Artifact, masterSourceBytes, "image/png");
        var layoutSourceAssetId = AddPictureAsset(request.Artifact, layoutSourceBytes, "image/png");
        request.Artifact.Presentation.Masters.Add(new PresentationMaster
        {
            Id = "master/source-image-background-replacement",
            Name = "Source image background replacement master",
            Background = new PresentationBackground
            {
                ImagePaint = new PresentationImagePaint
                {
                    AssetId = masterSourceAssetId,
                    Mode = PresentationImagePaint.Types.Mode.Stretch,
                    OpacityThousandthPercent = 89_000,
                },
            },
        });
        request.Artifact.Presentation.Layouts.Add(new PresentationLayout
        {
            Id = "layout/source-image-background-replacement",
            Name = "Source image background replacement layout",
            MasterId = "master/source-image-background-replacement",
            Type = "blank",
            Background = new PresentationBackground
            {
                ImagePaint = new PresentationImagePaint
                {
                    AssetId = layoutSourceAssetId,
                    Mode = PresentationImagePaint.Types.Mode.Stretch,
                    OpacityThousandthPercent = 79_000,
                },
            },
        });
        request.Artifact.Presentation.Slides[0].LayoutId = "layout/source-image-background-replacement";
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());

        string masterSourceRelationshipId;
        string layoutSourceRelationshipId;
        string masterImagePartPath;
        string layoutImagePartPath;
        using (var stream = new MemoryStream(source, writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var masterPart = Assert.Single(package.PresentationPart!.SlideMasterParts);
            var layoutPart = Assert.Single(masterPart.SlideLayoutParts);
            var masterBlip = masterPart.SlideMaster!.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!.GetFirstChild<A.Blip>()!;
            var layoutBlip = layoutPart.SlideLayout!.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!.GetFirstChild<A.Blip>()!;
            masterSourceRelationshipId = masterBlip.Embed!.Value!;
            layoutSourceRelationshipId = layoutBlip.Embed!.Value!;
            masterImagePartPath = masterPart.GetPartById(masterSourceRelationshipId).Uri.OriginalString.TrimStart('/');
            layoutImagePartPath = layoutPart.GetPartById(layoutSourceRelationshipId).Uri.OriginalString.TrimStart('/');
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "master-layout-image-background-replacement/source.pptx",
                AssetRootUri = "master-layout-image-background-replacement/assets",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var masterBackground = program["design"]!["masters"]!.AsArray().Single()!["background"]!.AsObject();
        var layoutBackground = program["design"]!["layouts"]!.AsArray().Single()!["background"]!.AsObject();
        var projectedMasterAssetId = masterBackground["asset"]!.GetValue<string>();
        var projectedLayoutAssetId = layoutBackground["asset"]!.GetValue<string>();
        var assets = program["assets"]!.AsArray();
        var masterReplacementDeclaration = assets
            .Select(asset => asset?.AsObject() ?? throw new InvalidOperationException("Projected asset declaration is null."))
            .Single(asset => asset["id"]!.GetValue<string>() == projectedMasterAssetId)
            .DeepClone()
            .AsObject();
        var layoutReplacementDeclaration = assets
            .Select(asset => asset?.AsObject() ?? throw new InvalidOperationException("Projected asset declaration is null."))
            .Single(asset => asset["id"]!.GetValue<string>() == projectedLayoutAssetId)
            .DeepClone()
            .AsObject();
        var masterReplacementAssetId = "background-master-replacement";
        var layoutReplacementAssetId = "background-layout-replacement";
        var masterReplacementHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(masterReplacementBytes)).ToLowerInvariant();
        var layoutReplacementHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(layoutReplacementBytes)).ToLowerInvariant();
        masterReplacementDeclaration["id"] = masterReplacementAssetId;
        masterReplacementDeclaration["uri"] = "master-layout-image-background-replacement/assets/master-replacement.png";
        masterReplacementDeclaration["sha256"] = masterReplacementHash;
        layoutReplacementDeclaration["id"] = layoutReplacementAssetId;
        layoutReplacementDeclaration["uri"] = "master-layout-image-background-replacement/assets/layout-replacement.png";
        layoutReplacementDeclaration["sha256"] = layoutReplacementHash;
        assets.Add(masterReplacementDeclaration);
        assets.Add(layoutReplacementDeclaration);
        masterBackground["asset"] = masterReplacementAssetId;
        masterBackground["opacity"] = 0.41;
        layoutBackground["asset"] = layoutReplacementAssetId;
        layoutBackground["opacity"] = 0.37;

        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
                Assets =
                {
                    new Asset
                    {
                        Id = masterReplacementAssetId,
                        FileName = "master-replacement.png",
                        ContentType = "image/png",
                        Data = ByteString.CopyFrom(masterReplacementBytes),
                        Sha256 = masterReplacementHash,
                    },
                    new Asset
                    {
                        Id = layoutReplacementAssetId,
                        FileName = "layout-replacement.png",
                        ContentType = "image/png",
                        Data = ByteString.CopyFrom(layoutReplacementBytes),
                        Sha256 = layoutReplacementHash,
                    },
                },
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains("ppt/slideMasters/slideMaster1.xml", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideMasters/_rels/slideMaster1.xml.rels", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideLayouts/slideLayout1.xml", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideLayouts/_rels/slideLayout1.xml.rels", edited.PresentationProgram.ChangedParts);
        Assert.Contains(edited.PresentationProgram.ChangedParts, path => path.StartsWith("ppt/media/", StringComparison.Ordinal));
        var editedPaths = ZipPartPaths(edited.File.ToByteArray());
        Assert.DoesNotContain(masterImagePartPath, editedPaths, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain(layoutImagePartPath, editedPaths, StringComparer.OrdinalIgnoreCase);

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
            var masterPart = Assert.Single(package.PresentationPart!.SlideMasterParts);
            var layoutPart = Assert.Single(masterPart.SlideLayoutParts);
            var masterBlip = masterPart.SlideMaster!.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!.GetFirstChild<A.Blip>()!;
            var layoutBlip = layoutPart.SlideLayout!.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!.GetFirstChild<A.Blip>()!;
            var masterOutputRelationshipId = masterBlip.Embed?.Value ?? throw new InvalidOperationException("Replaced master background relationship is missing.");
            var layoutOutputRelationshipId = layoutBlip.Embed?.Value ?? throw new InvalidOperationException("Replaced layout background relationship is missing.");
            Assert.NotEqual(masterSourceRelationshipId, masterOutputRelationshipId);
            Assert.NotEqual(layoutSourceRelationshipId, layoutOutputRelationshipId);
            var masterReplacementPart = Assert.IsType<ImagePart>(masterPart.GetPartById(masterOutputRelationshipId));
            var layoutReplacementPart = Assert.IsType<ImagePart>(layoutPart.GetPartById(layoutOutputRelationshipId));
            using (var replacementStream = masterReplacementPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var replacementMemory = new MemoryStream())
            {
                replacementStream.CopyTo(replacementMemory);
                Assert.Equal(masterReplacementBytes, replacementMemory.ToArray());
            }
            using (var replacementStream = layoutReplacementPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var replacementMemory = new MemoryStream())
            {
                replacementStream.CopyTo(replacementMemory);
                Assert.Equal(layoutReplacementBytes, replacementMemory.ToArray());
            }
            Assert.Equal(41_000, masterBlip.GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
            Assert.Equal(37_000, layoutBlip.GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "master-layout-image-background-replacement/edited.pptx",
                AssetRootUri = "master-layout-image-background-replacement/assets",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var output = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var outputMasterBackground = output["design"]!["masters"]!.AsArray().Single()!["background"]!.AsObject();
        var outputLayoutBackground = output["design"]!["layouts"]!.AsArray().Single()!["background"]!.AsObject();
        var outputMasterAsset = output["assets"]!.AsArray()
            .Select(asset => asset?.AsObject() ?? throw new InvalidOperationException("Reprojected asset declaration is null."))
            .Single(asset => asset["id"]!.GetValue<string>() == outputMasterBackground["asset"]!.GetValue<string>());
        var outputLayoutAsset = output["assets"]!.AsArray()
            .Select(asset => asset?.AsObject() ?? throw new InvalidOperationException("Reprojected asset declaration is null."))
            .Single(asset => asset["id"]!.GetValue<string>() == outputLayoutBackground["asset"]!.GetValue<string>());
        Assert.Equal(masterReplacementHash, outputMasterAsset["sha256"]!.GetValue<string>());
        Assert.Equal(layoutReplacementHash, outputLayoutAsset["sha256"]!.GetValue<string>());
        Assert.Equal(0.41, outputMasterBackground["opacity"]!.GetValue<double>(), precision: 6);
        Assert.Equal(0.37, outputLayoutBackground["opacity"]!.GetValue<double>(), precision: 6);
    }

    [Fact]
    public void PpjSourceBoundMasterAndLayoutPlaceholderEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var master = new PresentationMaster
        {
            Id = "master/source-placeholder",
            Name = "Source placeholder master",
        };
        master.Placeholders.Add(new PresentationPlaceholder
        {
            Id = "master/source-placeholder/title",
            Name = "Master title",
            Type = "title",
            Index = 0,
            DirectFrame = new PresentationPlaceholderFrame
            {
                LeftEmu = 571_500,
                TopEmu = 381_000,
                WidthEmu = 8_191_500,
                HeightEmu = 666_750,
                RotationAngle60000 = 9 * 60_000,
                FlipHorizontal = false,
                FlipVertical = true,
            },
            TextBody = new PresentationTextBody
            {
                BodyProperties = new PresentationTextBodyProperties
                {
                    VerticalAnchor = "center",
                    RotationAngle60000 = 6 * 60_000,
                },
                Paragraphs =
                {
                    new PresentationTextParagraph
                    {
                        Runs = { new PresentationTextRun { Text = "Master title" } },
                    },
                },
            },
        });
        authoredRequest.Artifact!.Presentation!.Masters.Add(master);

        var layout = new PresentationLayout
        {
            Id = "layout/source-placeholder",
            Name = "Source placeholder layout",
            MasterId = master.Id,
            Type = "titleOnly",
        };
        layout.Placeholders.Add(new PresentationPlaceholder
        {
            Id = "layout/source-placeholder/title",
            Name = "Layout title",
            Type = "title",
            Index = 0,
            DirectFrame = new PresentationPlaceholderFrame
            {
                LeftEmu = 571_500,
                TopEmu = 1_333_500,
                WidthEmu = 8_191_500,
                HeightEmu = 666_750,
            },
            TextBody = new PresentationTextBody
            {
                BodyProperties = new PresentationTextBodyProperties
                {
                    VerticalAnchor = "bottom",
                    VerticalOverflowMode = "ellipsis",
                },
                Paragraphs =
                {
                    new PresentationTextParagraph
                    {
                        Runs = { new PresentationTextRun { Text = "Layout title" } },
                    },
                },
            },
        });
        authoredRequest.Artifact.Presentation.Layouts.Add(layout);
        authoredRequest.Artifact.Presentation.Slides[0].LayoutId = layout.Id;

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
                SourceUri = "deck.assets/source/master-placeholder.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var masterJson = program["design"]!["masters"]!.AsArray().Single()!.AsObject();
        var layoutJson = program["design"]!["layouts"]!.AsArray().Single()!.AsObject();
        var masterPlaceholder = masterJson["placeholders"]!.AsArray().Single()!.AsObject();
        var layoutPlaceholder = layoutJson["placeholders"]!.AsArray().Single()!.AsObject();
        Assert.Contains(masterPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setFrame");
        Assert.Contains(masterPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "replaceText");
        Assert.Contains(masterPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setTextBodyStyle");
        Assert.Contains(layoutPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setFrame");
        Assert.Contains(layoutPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "replaceText");
        Assert.Contains(layoutPlaceholder["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setTextBodyStyle");
        Assert.Equal(9, masterPlaceholder["frame"]!["rotation"]!.GetValue<double>(), 3);
        Assert.False(masterPlaceholder["frame"]!["flipH"]!.GetValue<bool>());
        Assert.True(masterPlaceholder["frame"]!["flipV"]!.GetValue<bool>());
        Assert.Equal(6, masterPlaceholder["style"]!["rotation"]!.GetValue<double>(), 3);
        Assert.Equal("ellipsis", layoutPlaceholder["style"]!["verticalOverflow"]!.GetValue<string>());
        masterPlaceholder["style"]!["rotation"] = 20;
        masterPlaceholder["style"]!["verticalAlignment"] = "bottom";
        layoutPlaceholder["style"]!["verticalOverflow"] = "clip";

        var tampered = JsonNode.Parse(program.ToJsonString())!.AsObject();
        tampered["design"]!["masters"]!.AsArray().Single()!["placeholders"]!.AsArray().Single()!["nativeRef"]!["capabilities"]!
            .AsArray().RemoveAt(0);
        tampered["design"]!["masters"]!.AsArray().Single()!["placeholders"]!.AsArray().Single()!["frame"]!["x"] = 80;
        var rejected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(tampered.ToJsonString()),
            },
        });
        Assert.False(rejected.Ok);
        Assert.Equal("ppj.nativeRef.stale", Assert.Single(rejected.Diagnostics).Code);

        masterPlaceholder["frame"]!["x"] = 80;
        masterPlaceholder["text"]!["paragraphs"]![0]!["runs"]![0]!["text"] = "Updated master";
        layoutPlaceholder["frame"]!["y"] = 120;
        layoutPlaceholder["text"]!["paragraphs"]![0]!["runs"]![0]!["text"] = "Updated layout";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains("ppt/slideMasters/slideMaster1.xml", edited.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slideLayouts/slideLayout1.xml", edited.PresentationProgram.ChangedParts);
        Assert.DoesNotContain("ppt/slides/slide1.xml", edited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeMasterPlaceholder = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single(shape =>
                    shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!.Type!.InnerText == "title");
            Assert.Equal(80 * 12_700, nativeMasterPlaceholder.ShapeProperties!.Transform2D!.Offset!.X!.Value);
            Assert.Contains(nativeMasterPlaceholder.TextBody!.Descendants<A.Text>(), text => text.Text == "Updated master");
            Assert.Equal(9 * 60_000, nativeMasterPlaceholder.ShapeProperties.Transform2D!.Rotation!.Value);
            Assert.False(nativeMasterPlaceholder.ShapeProperties.Transform2D.HorizontalFlip!.Value);
            Assert.True(nativeMasterPlaceholder.ShapeProperties.Transform2D.VerticalFlip!.Value);
            Assert.Equal(20 * 60_000, nativeMasterPlaceholder.TextBody.GetFirstChild<A.BodyProperties>()!.Rotation!.Value);
            Assert.Equal(A.TextAnchoringTypeValues.Bottom, nativeMasterPlaceholder.TextBody.GetFirstChild<A.BodyProperties>()!.Anchor!.Value);
            var nativeLayoutPlaceholder = package.PresentationPart.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout!
                .CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single(shape =>
                    shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!.Type!.InnerText == "title");
            Assert.Equal(120 * 12_700, nativeLayoutPlaceholder.ShapeProperties!.Transform2D!.Offset!.Y!.Value);
            Assert.Contains(nativeLayoutPlaceholder.TextBody!.Descendants<A.Text>(), text => text.Text == "Updated layout");
            Assert.Equal(A.TextVerticalOverflowValues.Clip, nativeLayoutPlaceholder.TextBody.GetFirstChild<A.BodyProperties>()!.VerticalOverflow!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/master-placeholder-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var roundTrip = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var roundTripMasterPlaceholder = roundTrip["design"]!["masters"]![0]!["placeholders"]![0]!;
        var roundTripLayoutPlaceholder = roundTrip["design"]!["layouts"]![0]!["placeholders"]![0]!;
        Assert.Equal(80, roundTripMasterPlaceholder!["frame"]!["x"]!.GetValue<double>());
        Assert.Equal("Updated master", roundTripMasterPlaceholder["text"]!["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>());
        Assert.Equal(9, roundTripMasterPlaceholder["frame"]!["rotation"]!.GetValue<double>(), 3);
        Assert.False(roundTripMasterPlaceholder["frame"]!["flipH"]!.GetValue<bool>());
        Assert.True(roundTripMasterPlaceholder["frame"]!["flipV"]!.GetValue<bool>());
        Assert.Equal(20, roundTripMasterPlaceholder["style"]!["rotation"]!.GetValue<double>(), 3);
        Assert.Equal("bottom", roundTripMasterPlaceholder["style"]!["verticalAlignment"]!.GetValue<string>());
        Assert.Equal(120, roundTripLayoutPlaceholder!["frame"]!["y"]!.GetValue<double>());
        Assert.Equal("Updated layout", roundTripLayoutPlaceholder["text"]!["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("clip", roundTripLayoutPlaceholder["style"]!["verticalOverflow"]!.GetValue<string>());

        // Optional native transform attributes are independently editable:
        // remove rotation/vertical flip while retaining an explicit false
        // horizontal flip, then prove the second projection preserves that
        // presence distinction.
        roundTripMasterPlaceholder["frame"]!.AsObject().Remove("rotation");
        roundTripMasterPlaceholder["frame"]!.AsObject().Remove("flipV");
        var transformEdited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(roundTrip.ToJsonString()),
            },
        });
        Assert.True(transformEdited.Ok, Diagnostics(transformEdited));
        Assert.Equal(["ppt/slideMasters/slideMaster1.xml"], transformEdited.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(transformEdited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeMasterPlaceholder = package.PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single(shape =>
                    shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!.Type!.InnerText == "title");
            var transform = nativeMasterPlaceholder.ShapeProperties!.Transform2D!;
            Assert.Null(transform.Rotation);
            Assert.False(transform.HorizontalFlip!.Value);
            Assert.Null(transform.VerticalFlip);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var transformReprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = transformEdited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/master-placeholder-transform-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(transformReprojected.Ok, Diagnostics(transformReprojected));
        var transformRoundTrip = JsonNode.Parse(transformReprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var transformFrame = transformRoundTrip["design"]!["masters"]![0]!["placeholders"]![0]!["frame"]!.AsObject();
        Assert.False(transformFrame.ContainsKey("rotation"));
        Assert.False(transformFrame.ContainsKey("flipV"));
        Assert.False(transformFrame["flipH"]!.GetValue<bool>());
    }

    [Fact]
    public void PpjParagraphTabStopsAuthorAndReproject()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var fixture = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var title = fixture["pages"]![0]!["elements"]!.AsArray()
            .Select(node => node!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-title");
        title["text"]!["paragraphs"]![0]!["style"] = new JsonObject
        {
            ["alignment"] = "distributed",
            ["tabStops"] = new JsonArray(
                new JsonObject { ["position"] = 90, ["alignment"] = "left" },
                new JsonObject { ["position"] = 180, ["alignment"] = "decimal" }),
        };
        var page = fixture["pages"]![0]!.DeepClone().AsObject();
        page["elements"] = new JsonArray(title.DeepClone());
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        fixture["assets"] = new JsonArray();
        fixture["components"] = new JsonArray();
        fixture["pages"] = new JsonArray(page);
        fixture["sections"] = new JsonArray();
        fixture["customShows"] = new JsonArray();
        fixture["comments"] = new JsonArray();

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(fixture.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
                .Single(shape => shape.TextBody is not null);
            var nativeParagraph = nativeShape.TextBody!.Elements<A.Paragraph>().Single();
            var tabStops = nativeParagraph.ParagraphProperties!.GetFirstChild<A.TabStopList>()!.Elements<A.TabStop>().ToArray();
            Assert.Equal(A.TextAlignmentTypeValues.Distributed, nativeParagraph.ParagraphProperties.Alignment!.Value);
            Assert.Equal(new[] { 1_143_000, 2_286_000 }, tabStops.Select(tab => tab.Position!.Value));
            Assert.Equal(A.TextTabAlignmentValues.Decimal, tabStops[1].Alignment!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/tab-stops.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedTitle = state["pages"]![0]!["elements"]!.AsArray()
            .Select(node => node!.AsObject())
            .Single();
        var projectedTabs = projectedTitle["text"]!["paragraphs"]![0]!["style"]!["tabStops"]!.AsArray();
        Assert.Equal(2, projectedTabs.Count);
        Assert.Equal("distributed", projectedTitle["text"]!["paragraphs"]![0]!["style"]!["alignment"]!.GetValue<string>());
        Assert.Equal(90, projectedTabs[0]! ["position"]!.GetValue<double>());
        Assert.Equal("decimal", projectedTabs[1]! ["alignment"]!.GetValue<string>());

        projectedTabs[0]!.AsObject()["position"] = 120;
        projectedTabs[1]!.AsObject()["alignment"] = "center";
        projectedTitle["text"]!["paragraphs"]![0]!["style"]!["alignment"] = "center";
        var sourceBound = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
            },
        });
        Assert.True(sourceBound.Ok, Diagnostics(sourceBound));
        Assert.Equal(["ppt/slides/slide1.xml"], sourceBound.PresentationProgram.ChangedParts);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = sourceBound.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/tab-stops-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedState = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedTabs = editedState["pages"]![0]!["elements"]!.AsArray().Single()!
            ["text"]!["paragraphs"]![0]!["style"]!["tabStops"]!.AsArray();
        Assert.Equal(120, editedTabs[0]! ["position"]!.GetValue<double>());
        Assert.Equal("center", editedTabs[1]! ["alignment"]!.GetValue<string>());
        Assert.Equal("center", editedState["pages"]![0]!["elements"]!.AsArray().Single()!
            ["text"]!["paragraphs"]![0]!["style"]!["alignment"]!.GetValue<string>());
    }

    [Fact]
    public void PpjModernCommentsProjectAndReprojectTextAndStatus()
    {
        var request = ModernCommentExportRequest();
        request.Artifact!.Presentation!.Slides[0].ModernComments[0].Replies.Add(new PresentationModernComment
        {
            Id = "{55555555-5555-4555-8555-555555555555}",
            AuthorId = "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
            Author = "Evidence Reviewer",
            Initials = "ER",
            UserId = "evidence.reviewer@example.test",
            ProviderId = "None",
            Text = "I will verify the source binding.",
            CreatedAt = "2026-07-19T05:05:00Z",
            Status = "active",
        });
        var authored = Invoke(request);
        Assert.True(authored.Ok, Diagnostics(authored));
        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/modern-comments.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        using var projectedJson = JsonDocument.Parse(projected.PresentationProgram.ProgramJson.ToByteArray());
        var comments = projectedJson.RootElement.GetProperty("comments");
        Assert.Equal(2, comments.GetArrayLength());
        Assert.All(comments.EnumerateArray(), comment => Assert.Equal("modern", comment.GetProperty("kind").GetString()));
        Assert.Equal("element", comments[0].GetProperty("anchor").GetProperty("kind").GetString());
        Assert.Equal("spMk", comments[0].GetProperty("anchor").GetProperty("moniker").GetString());
        Assert.Equal(comments[0].GetProperty("id").GetString(), comments[1].GetProperty("parent").GetString());
        Assert.Contains(comments[0].GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
            capability.GetProperty("operation").GetString() == "setCommentStatus" &&
            capability.GetProperty("fields").EnumerateArray().Select(field => field.GetString()).SequenceEqual(["status", "resolved"]));

        var editedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedComments = editedProgram["comments"]!.AsArray();
        editedComments[0]!["text"] = "Customer evidence confirmed.";
        editedComments[0]!["status"] = "resolved";
        editedComments[0]!["resolved"] = true;
        editedComments[1]!["text"] = "Recorded in the decision log.";
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(editedProgram.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/comments/modernComment.xml"], edited.PresentationProgram.ChangedParts);
        Assert.Contains(edited.PresentationProgram.ChangedNodeIds, id => id == comments[0].GetProperty("id").GetString());
        Assert.Contains(edited.PresentationProgram.ChangedNodeIds, id => id == comments[1].GetProperty("id").GetString());

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/modern-comments-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        using var reprojectedJson = JsonDocument.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray());
        var reprojectedComments = reprojectedJson.RootElement.GetProperty("comments");
        Assert.Equal("Customer evidence confirmed.", reprojectedComments[0].GetProperty("text").GetString());
        Assert.Equal("resolved", reprojectedComments[0].GetProperty("status").GetString());
        Assert.True(reprojectedComments[0].GetProperty("resolved").GetBoolean());
        Assert.Equal("Recorded in the decision log.", reprojectedComments[1].GetProperty("text").GetString());
    }

    [Fact]
    public void PpjSourceFreeModernCommentsAuthorAndReproject()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = program["pages"]![0]!.DeepClone()!.AsObject();
        page["elements"] = new JsonArray(page["elements"]!.AsArray()
            .Where(element => element!["type"]!.GetValue<string>() != "image")
            .Select(element => element!.DeepClone())
            .ToArray());
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        var pageId = page["id"]!.GetValue<string>();
        program["pages"] = new JsonArray(page);
        program["assets"] = new JsonArray();
        program["components"] = new JsonArray();
        program["sections"] = new JsonArray();
        program["customShows"] = new JsonArray();
        program["comments"] = new JsonArray(
            new JsonObject
            {
                ["id"] = "modern-root",
                ["page"] = pageId,
                ["kind"] = "modern",
                ["target"] = "claim-band",
                ["author"] = "Review Owner",
                ["text"] = "Confirm this decision.",
                ["createdAt"] = "2026-09-03T04:00:00Z",
                ["resolved"] = false,
                ["status"] = "active",
                ["position"] = new JsonObject { ["x"] = 24, ["y"] = 18 },
                ["anchor"] = new JsonObject
                {
                    ["kind"] = "textRange",
                    ["moniker"] = "spMk",
                    ["textStart"] = 0,
                    ["textLength"] = 8,
                },
            },
            new JsonObject
            {
                ["id"] = "modern-reply",
                ["page"] = pageId,
                ["kind"] = "modern",
                ["parent"] = "modern-root",
                ["author"] = "Evidence Owner",
                ["text"] = "The evidence is attached.",
                ["createdAt"] = "2026-09-03T04:05:00Z",
                ["resolved"] = false,
                ["status"] = "active",
            });

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            Assert.Single(package.PresentationPart!.Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointAuthorsPart>());
            Assert.Single(Assert.Single(package.PresentationPart.SlideParts).Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/modern-comments-authored.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        using var projectedJson = JsonDocument.Parse(projected.PresentationProgram.ProgramJson.ToByteArray());
        var comments = projectedJson.RootElement.GetProperty("comments");
        var projectedTarget = projectedJson.RootElement.GetProperty("pages")[0]
            .GetProperty("elements")
            .EnumerateArray()
            .ElementAt(2)
            .GetProperty("id")
            .GetString();
        Assert.Equal(2, comments.GetArrayLength());
        Assert.Equal("modern", comments[0].GetProperty("kind").GetString());
        Assert.Equal("textRange", comments[0].GetProperty("anchor").GetProperty("kind").GetString());
        Assert.Equal(projectedTarget, comments[0].GetProperty("target").GetString());
        Assert.Equal(comments[0].GetProperty("id").GetString(), comments[1].GetProperty("parent").GetString());
        Assert.Equal("Evidence Owner", comments[1].GetProperty("author").GetString());
    }

    [Fact]
    public void PpjSourceBoundTranslucentSolidBackgroundEditsAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = program["pages"]![0]!.DeepClone()!.AsObject();
        page["elements"] = new JsonArray(page["elements"]!.AsArray()
            .Where(element => element!["type"]!.GetValue<string>() != "image")
            .Select(element => element!.DeepClone())
            .ToArray());
        page["background"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = "#F7F4EC",
        };
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        program["pages"] = new JsonArray(page);
        program["assets"] = new JsonArray();
        program["components"] = new JsonArray();
        program["sections"] = new JsonArray();
        program["customShows"] = new JsonArray();
        program.Remove("comments");
        program["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["paperSurface"] = new JsonObject
            {
                ["kind"] = "color",
                ["value"] = "#102030",
            },
        };

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/solid-background-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var editedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        editedProgram["design"]!["grammar"]!["tokens"] = new JsonObject
        {
            ["paperSurface"] = new JsonObject
            {
                ["kind"] = "color",
                ["value"] = "#102030",
            },
        };
        var editedShape = editedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .First(element => element["type"]?.GetValue<string>() == "shape" &&
                element["style"]?["fill"]?["type"]?.GetValue<string>() == "solid");
        var editedShapeId = editedShape["id"]!.GetValue<string>();
        var gradientShape = editedProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .First(element => element["type"]?.GetValue<string>() == "shape" &&
                element["id"]?.GetValue<string>() != editedShapeId);
        var gradientShapeId = gradientShape["id"]!.GetValue<string>();
        gradientShape["style"] ??= new JsonObject();
        gradientShape["style"]!["fill"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "linear",
            ["angle"] = 18,
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = new JsonObject { ["token"] = "paperSurface" } },
                new JsonObject { ["offset"] = 1, ["color"] = "#FFFFFF", ["opacity"] = 0.7 },
            },
        };
        editedShape["style"]!["fill"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = new JsonObject { ["token"] = "paperSurface" },
        };
        editedShape["style"]!["stroke"] = new JsonObject
        {
            ["color"] = new JsonObject { ["token"] = "paperSurface" },
            ["width"] = 1.25,
            ["opacity"] = 0.4,
        };
        editedProgram["pages"]![0]!["background"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = new JsonObject { ["token"] = "paperSurface" },
            ["opacity"] = 0.33,
        };
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(editedProgram.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains("ppt/slides/slide1.xml", edited.PresentationProgram.ChangedParts);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/solid-background-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        using var json = JsonDocument.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray());
        var background = json.RootElement.GetProperty("pages")[0].GetProperty("background");
        Assert.Equal("solid", background.GetProperty("type").GetString());
        Assert.Equal("#102030", background.GetProperty("color").GetString());
        Assert.Equal(0.33, background.GetProperty("opacity").GetDouble(), 3);
        var shapeFill = json.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
            .Single(element => element.GetProperty("id").GetString() == editedShapeId)
            .GetProperty("style").GetProperty("fill");
        Assert.Equal("#102030", shapeFill.GetProperty("color").GetString());
        var shapeStroke = json.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
            .Single(element => element.GetProperty("id").GetString() == editedShapeId)
            .GetProperty("style").GetProperty("stroke");
        Assert.Equal("#102030", shapeStroke.GetProperty("color").GetString());
        var gradientFill = json.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
            .Single(element => element.GetProperty("id").GetString() == gradientShapeId)
            .GetProperty("style").GetProperty("fill");
        Assert.Equal("gradient", gradientFill.GetProperty("type").GetString());
        Assert.Equal("#102030", gradientFill.GetProperty("stops")[0].GetProperty("color").GetString());
    }

    [Fact]
    public void PpjRichTableTextBodyStyleAuthorsAndProjects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var fixture = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = fixture["pages"]!.AsArray()[1]!.DeepClone()!.AsObject();
        var table = page["elements"]!.AsArray()
            .Select(node => node!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "method-table-main");
        var cell = table["rows"]![0]! ["cells"]![0]!.AsObject();
        cell["text"] = new JsonObject
        {
            ["style"] = new JsonObject
            {
                ["verticalAlignment"] = "middle",
                ["wrap"] = "square",
                ["autoFit"] = "shrink-text",
                ["margins"] = new JsonObject
                {
                    ["left"] = 10,
                    ["top"] = 2,
                    ["right"] = 3,
                    ["bottom"] = 4,
                },
                ["columns"] = 2,
                ["columnGap"] = 4,
                ["columnDirection"] = "right-to-left",
                ["verticalText"] = "vertical",
            },
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["runs"] = new JsonArray
                    {
                        new JsonObject { ["text"] = "Protocol" },
                    },
                },
            },
        };
        page["elements"] = new JsonArray(table.DeepClone());
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        fixture["assets"] = new JsonArray();
        fixture["components"] = new JsonArray();
        fixture["pages"] = new JsonArray(page);
        fixture["sections"] = new JsonArray();
        fixture["customShows"] = new JsonArray();
        fixture["comments"] = new JsonArray();

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(fixture.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeCell = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>())
                .Elements<A.TableRow>().First()
                .Elements<A.TableCell>().First();
            var bodyProperties = nativeCell.GetFirstChild<A.TextBody>()!.GetFirstChild<A.BodyProperties>()!;
            Assert.Equal(A.TextAnchoringTypeValues.Center, bodyProperties.Anchor!.Value);
            Assert.Equal(A.TextWrappingValues.Square, bodyProperties.Wrap!.Value);
            Assert.NotNull(bodyProperties.GetFirstChild<A.NormalAutoFit>());
            Assert.Equal(10 * 12_700, bodyProperties.LeftInset!.Value);
            Assert.Equal(2 * 12_700, bodyProperties.TopInset!.Value);
            Assert.Equal(3 * 12_700, bodyProperties.RightInset!.Value);
            Assert.Equal(4 * 12_700, bodyProperties.BottomInset!.Value);
            Assert.Equal(2, bodyProperties.ColumnCount!.Value);
            Assert.Equal(4 * 12_700, bodyProperties.ColumnSpacing!.Value);
            Assert.True(bodyProperties.RightToLeftColumns!.Value);
            Assert.Equal(A.TextVerticalValues.Vertical, bodyProperties.Vertical!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/table-body-style.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedCell = state["pages"]![0]!["elements"]!.AsArray()
            .Single()!["rows"]![0]!["cells"]![0]!.AsObject();
        var projectedStyle = projectedCell["text"]!["style"]!;
        Assert.Equal("middle", projectedStyle["verticalAlignment"]!.GetValue<string>());
        Assert.Equal("square", projectedStyle["wrap"]!.GetValue<string>());
        Assert.Equal("shrink-text", projectedStyle["autoFit"]!.GetValue<string>());
        Assert.Equal(10, projectedStyle["margins"]!["left"]!.GetValue<double>(), 3);
        Assert.Equal(2, projectedStyle["margins"]!["top"]!.GetValue<double>(), 3);
        Assert.Equal(3, projectedStyle["margins"]!["right"]!.GetValue<double>(), 3);
        Assert.Equal(4, projectedStyle["margins"]!["bottom"]!.GetValue<double>(), 3);
        Assert.Equal(2, projectedStyle["columns"]!.GetValue<int>());
        Assert.Equal(4, projectedStyle["columnGap"]!.GetValue<double>(), 3);
        Assert.Equal("right-to-left", projectedStyle["columnDirection"]!.GetValue<string>());
        Assert.Equal("vertical", projectedStyle["verticalText"]!.GetValue<string>());
        Assert.Equal("Protocol", projectedCell["text"]!["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var fixture = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = fixture["pages"]!.AsArray()[0]!.DeepClone()!.AsObject();
        var band = page["elements"]!.AsArray()
            .Select(node => node!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-band");
        band["textStyle"] = new JsonObject
        {
            ["verticalAlignment"] = "middle",
            ["wrap"] = "square",
            ["autoFit"] = "shrink-text",
            ["normalAutoFit"] = new JsonObject
            {
                ["fontScale"] = 80.125,
                ["lineSpacingReduction"] = 12.5,
            },
            ["margins"] = new JsonObject
            {
                ["left"] = 18,
                ["top"] = 8,
                ["right"] = 18,
                ["bottom"] = 8,
            },
            ["columns"] = 1,
            ["columnGap"] = 2,
            ["columnDirection"] = "left-to-right",
            ["verticalText"] = "horizontal",
            ["rotation"] = 12,
            ["verticalOverflow"] = "ellipsis",
            ["horizontalOverflow"] = "clip",
            ["upright"] = true,
        };
        page["elements"] = new JsonArray(band.DeepClone());
        page.Remove("notes");
        page.Remove("transition");
        page.Remove("animations");
        page.Remove("sourceClone");
        fixture["assets"] = new JsonArray();
        fixture["components"] = new JsonArray();
        fixture["pages"] = new JsonArray(page);
        fixture["sections"] = new JsonArray();
        fixture["customShows"] = new JsonArray();
        fixture["comments"] = new JsonArray();

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(fixture.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        using (var stream = new MemoryStream(authored.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var bodyProperties = nativeShape.TextBody!.GetFirstChild<A.BodyProperties>()!;
            Assert.Equal(A.TextAnchoringTypeValues.Center, bodyProperties.Anchor!.Value);
            Assert.Equal(A.TextWrappingValues.Square, bodyProperties.Wrap!.Value);
            Assert.Equal(A.TextVerticalValues.Horizontal, bodyProperties.Vertical!.Value);
            var normalAutoFit = bodyProperties.GetFirstChild<A.NormalAutoFit>();
            Assert.NotNull(normalAutoFit);
            Assert.Equal(80_125, normalAutoFit!.FontScale!.Value);
            Assert.Equal(12_500, normalAutoFit.LineSpaceReduction!.Value);
            Assert.Equal(18 * 12_700, bodyProperties.LeftInset!.Value);
            Assert.Equal(8 * 12_700, bodyProperties.TopInset!.Value);
            Assert.Equal(1, bodyProperties.ColumnCount!.Value);
            Assert.Equal(12 * 60_000, bodyProperties.Rotation!.Value);
            Assert.Equal(A.TextVerticalOverflowValues.Ellipsis, bodyProperties.VerticalOverflow!.Value);
            Assert.Equal(A.TextHorizontalOverflowValues.Clip, bodyProperties.HorizontalOverflow!.Value);
            Assert.True(bodyProperties.UpRight!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/text-body-style.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = state["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var capabilities = projectedShape["nativeRef"]!["capabilities"]!.AsArray();
        Assert.Contains(capabilities, capability => capability!["operation"]!.GetValue<string>() == "setTextBodyStyle");
        var projectedStyle = projectedShape["textStyle"]!.AsObject();
        Assert.Equal(80.125, projectedStyle["normalAutoFit"]!["fontScale"]!.GetValue<double>(), 3);
        Assert.Equal(12.5, projectedStyle["normalAutoFit"]!["lineSpacingReduction"]!.GetValue<double>(), 3);
        projectedStyle["verticalAlignment"] = "bottom";
        projectedStyle["wrap"] = "none";
        projectedStyle["autoFit"] = "shrink-text";
        projectedStyle["normalAutoFit"] = new JsonObject
        {
            ["fontScale"] = 72.5,
            ["lineSpacingReduction"] = 8.125,
        };
        projectedStyle["margins"]!["left"] = 24;
        projectedStyle["margins"]!["top"] = 3;
        projectedStyle["columns"] = 2;
        projectedStyle["columnGap"] = 4;
        projectedStyle["columnDirection"] = "right-to-left";
        projectedStyle["verticalText"] = "vertical270";
        Assert.Equal(12, projectedStyle["rotation"]!.GetValue<double>(), 3);
        Assert.Equal("ellipsis", projectedStyle["verticalOverflow"]!.GetValue<string>());
        Assert.Equal("clip", projectedStyle["horizontalOverflow"]!.GetValue<string>());
        Assert.True(projectedStyle["upright"]!.GetValue<bool>());
        projectedStyle["rotation"] = -18;
        projectedStyle["verticalOverflow"] = "overflow";
        projectedStyle["horizontalOverflow"] = "overflow";
        projectedStyle["upright"] = false;

        var sourceBound = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(authored.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
            },
        });
        Assert.True(sourceBound.Ok, Diagnostics(sourceBound));
        Assert.Equal(["ppt/slides/slide1.xml"], sourceBound.PresentationProgram.ChangedParts);
        using (var stream = new MemoryStream(sourceBound.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var nativeShape = package.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            var bodyProperties = nativeShape.TextBody!.GetFirstChild<A.BodyProperties>()!;
            Assert.Equal(A.TextAnchoringTypeValues.Bottom, bodyProperties.Anchor!.Value);
            Assert.Equal(A.TextWrappingValues.None, bodyProperties.Wrap!.Value);
            var normalAutoFit = bodyProperties.GetFirstChild<A.NormalAutoFit>();
            Assert.NotNull(normalAutoFit);
            Assert.Equal(72_500, normalAutoFit!.FontScale!.Value);
            Assert.Equal(8_125, normalAutoFit.LineSpaceReduction!.Value);
            Assert.Equal(24 * 12_700, bodyProperties.LeftInset!.Value);
            Assert.Equal(3 * 12_700, bodyProperties.TopInset!.Value);
            Assert.Equal(2, bodyProperties.ColumnCount!.Value);
            Assert.Equal(4 * 12_700, bodyProperties.ColumnSpacing!.Value);
            Assert.True(bodyProperties.RightToLeftColumns!.Value);
            Assert.Equal(A.TextVerticalValues.Vertical270, bodyProperties.Vertical!.Value);
            Assert.Equal(-18 * 60_000, bodyProperties.Rotation!.Value);
            Assert.Equal(A.TextVerticalOverflowValues.Overflow, bodyProperties.VerticalOverflow!.Value);
            Assert.Equal(A.TextHorizontalOverflowValues.Overflow, bodyProperties.HorizontalOverflow!.Value);
            Assert.False(bodyProperties.UpRight!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = sourceBound.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/text-body-style-edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedState = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedStyle = editedState["pages"]![0]!["elements"]!.AsArray().Single()!["textStyle"]!;
        Assert.Equal("bottom", editedStyle!["verticalAlignment"]!.GetValue<string>());
        Assert.Equal("none", editedStyle!["wrap"]!.GetValue<string>());
        Assert.Equal("shrink-text", editedStyle!["autoFit"]!.GetValue<string>());
        Assert.Equal(72.5, editedStyle!["normalAutoFit"]!["fontScale"]!.GetValue<double>(), 3);
        Assert.Equal(8.125, editedStyle!["normalAutoFit"]!["lineSpacingReduction"]!.GetValue<double>(), 3);
        Assert.Equal(24, editedStyle!["margins"]!["left"]!.GetValue<double>(), 3);
        Assert.Equal(3, editedStyle!["margins"]!["top"]!.GetValue<double>(), 3);
        Assert.Equal(2, editedStyle!["columns"]!.GetValue<int>());
        Assert.Equal(4, editedStyle!["columnGap"]!.GetValue<double>(), 3);
        Assert.Equal("right-to-left", editedStyle!["columnDirection"]!.GetValue<string>());
        Assert.Equal("vertical270", editedStyle!["verticalText"]!.GetValue<string>());
        Assert.Equal(-18, editedStyle!["rotation"]!.GetValue<double>(), 3);
        Assert.Equal("overflow", editedStyle!["verticalOverflow"]!.GetValue<string>());
        Assert.Equal("overflow", editedStyle!["horizontalOverflow"]!.GetValue<string>());
        Assert.False(editedStyle!["upright"]!.GetValue<bool>());
    }

    [Fact]
    public void PpjSourceBoundNormalAutoFitLeavesEditAndReproject()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);

        var fixture = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var page = fixture["pages"]!.AsArray()[0]!.DeepClone()!.AsObject();
        var band = page["elements"]!.AsArray()
            .Select(node => node!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-band");
        band["textStyle"] = new JsonObject
        {
            ["autoFit"] = "shrink-text",
            ["normalAutoFit"] = new JsonObject
            {
                ["fontScale"] = 80.125,
                ["lineSpacingReduction"] = 12.5,
            },
        };
        page["elements"] = new JsonArray(band.DeepClone());
        foreach (var field in new[] { "notes", "transition", "animations", "sourceClone" }) page.Remove(field);
        fixture["assets"] = new JsonArray();
        fixture["components"] = new JsonArray();
        fixture["pages"] = new JsonArray(page);
        fixture["sections"] = new JsonArray();
        fixture["customShows"] = new JsonArray();
        fixture["comments"] = new JsonArray();

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(fixture.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        var source = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/normal-autofit-leaves.pptx" },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var projectedShape = state["pages"]![0]!["elements"]![0]!.AsObject();
        var leaves = projectedShape["nativeRef"]!["leaves"]!.AsArray();
        var fontScale = leaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "textBodyNormalAutoFitFontScale")!.AsObject();
        var lineSpacing = leaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "textBodyNormalAutoFitLineSpacingReduction")!.AsObject();
        Assert.Equal(80.125, fontScale["value"]!.GetValue<double>(), 3);
        Assert.Equal(12.5, lineSpacing["value"]!.GetValue<double>(), 3);

        fontScale["value"] = 72.5;
        lineSpacing["value"] = 8.125;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        var editedXml = Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), "ppt/slides/slide1.xml"));
        Assert.Contains("fontScale=\"72500\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("lnSpcReduction=\"8125\"", editedXml, StringComparison.Ordinal);

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest { SourceUri = "deck.assets/source/normal-autofit-leaves-edited.pptx" },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedState = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedShape = editedState["pages"]![0]!["elements"]![0]!.AsObject();
        var editedLeaves = editedShape["nativeRef"]!["leaves"]!.AsArray();
        Assert.Equal(72.5, editedLeaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "textBodyNormalAutoFitFontScale")!["value"]!.GetValue<double>(), 3);
        Assert.Equal(8.125, editedLeaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "textBodyNormalAutoFitLineSpacingReduction")!["value"]!.GetValue<double>(), 3);
        Assert.Equal(72.5, editedShape["textStyle"]!["normalAutoFit"]!["fontScale"]!.GetValue<double>(), 3);
        Assert.Equal(8.125, editedShape["textStyle"]!["normalAutoFit"]!["lineSpacingReduction"]!.GetValue<double>(), 3);
    }

    [Fact]
    public void PpjSourceBoundLiteralCustomGeometryAdjustmentLeafEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var shape = authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape;
        shape.Geometry = "custom";
        shape.CustomAdjustments.Add(new PresentationCustomGeometryGuide { Name = "adjX", Formula = "val 25000" });
        shape.CustomGuides.Add(new PresentationCustomGeometryGuide { Name = "x1", Formula = "*/ w adjX 100000" });
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

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-geometry-adjustment.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = program["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        var adjustment = element["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "customGeometryAdjustment")!.AsObject();
        Assert.Equal(25000, adjustment["value"]!.GetValue<long>());

        adjustment["value"] = 30000;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        var editedXml = Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), "ppt/slides/slide1.xml"));
        Assert.Contains("fmla=\"val 30000\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-geometry-adjustment-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedAdjustment = editedProgram["pages"]![0]!["elements"]!.AsArray().Single()!
            ["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "customGeometryAdjustment");
        Assert.Equal(30000, editedAdjustment!["value"]!.GetValue<long>());
    }

    [Fact]
    public void PpjSourceBoundPresetGeometryAdjustmentLeafEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var shape = authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape;
        shape.Geometry = "roundRect";
        shape.PresetAdjustments.Add(25000);
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
                SourceUri = "deck.assets/source/preset-geometry-adjustment.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = program["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal(25000, element["geometry"]!["adjustments"]![0]!.GetValue<int>());
        var adjustment = element["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment")!.AsObject();
        Assert.Equal(25000, adjustment["value"]!.GetValue<long>());

        adjustment["value"] = 30000;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        var editedXml = Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), "ppt/slides/slide1.xml"));
        Assert.Contains("prst=\"roundRect\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("fmla=\"val 30000\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/preset-geometry-adjustment-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedElement = editedProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.Equal(30000, editedElement["geometry"]!["adjustments"]![0]!.GetValue<int>());
        Assert.Equal(30000, editedElement["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment")!["value"]!.GetValue<long>());

        var formulaSource = ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
        {
            Assert.Contains("fmla=\"val 25000\"", xml, StringComparison.Ordinal);
            return xml.Replace("fmla=\"val 25000\"", "fmla=\"*/ w 1 2\"", StringComparison.Ordinal);
        });
        var formulaProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(formulaSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/preset-geometry-adjustment-formula.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(formulaProjection.Ok, Diagnostics(formulaProjection));
        var formulaProgram = JsonNode.Parse(formulaProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var formulaElement = formulaProgram["pages"]![0]!["elements"]!.AsArray().Single()!.AsObject();
        Assert.DoesNotContain(formulaElement["nativeRef"]!["leaves"]!.AsArray(),
            leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment");
    }

    [Fact]
    public void PpjSourceBoundPartialPresetGeometryLiteralSiblingEditsAndReprojects()
    {
        var authoredRequest = ExportRequest();
        var shape = authoredRequest.Artifact!.Presentation!.Slides[0].Elements[0].Shape;
        shape.Geometry = "accentBorderCallout1";
        shape.PresetAdjustments.Add([1000, 2000, 3000, 4000]);
        Assert.Equal(4, shape.PresetAdjustments.Count);
        var authored = Invoke(authoredRequest);
        Assert.True(authored.Ok, Diagnostics(authored));
        var completeSource = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var source = ReplaceZipText(completeSource, "ppt/slides/slide1.xml", xml =>
        {
            const string geometryStart = "<a:prstGeom prst=\"accentBorderCallout1\"";
            const string listStart = "<a:avLst>";
            const string end = "</a:avLst></a:prstGeom>";
            var startIndex = xml.IndexOf(geometryStart, StringComparison.Ordinal);
            Assert.True(startIndex >= 0, "authored preset geometry was not found");
            var contentStart = xml.IndexOf(listStart, startIndex, StringComparison.Ordinal);
            var endIndex = xml.IndexOf(end, contentStart, StringComparison.Ordinal);
            Assert.True(endIndex > contentStart, "authored preset adjustment list was not found");
            var firstGuideEnd = xml.IndexOf("/>", contentStart, StringComparison.Ordinal);
            Assert.True(firstGuideEnd > contentStart && firstGuideEnd < endIndex, "authored first adjustment was not found");
            var secondGuideEnd = xml.IndexOf("/>", firstGuideEnd + 2, StringComparison.Ordinal);
            Assert.True(secondGuideEnd > firstGuideEnd && secondGuideEnd < endIndex, "authored second adjustment was not found");
            var partialGuides = xml[contentStart..(secondGuideEnd + 2)];
            return xml[..contentStart] + partialGuides + xml[endIndex..];
        });

        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/partial-preset-geometry.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var element = program["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.False(element["geometry"]!.AsObject().ContainsKey("adjustments"));
        var leaves = element["nativeRef"]!["leaves"]!.AsArray();
        var adjustment = leaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment" &&
            leaf!["value"]!.GetValue<long>() == 1000)!.AsObject();
        Assert.Equal(1000, adjustment["value"]!.GetValue<long>());

        adjustment["value"] = 1500;
        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        var editedXml = Encoding.UTF8.GetString(ZipBytes(edited.File.ToByteArray(), "ppt/slides/slide1.xml"));
        Assert.Contains("name=\"adj1\" fmla=\"val 1500\"", editedXml, StringComparison.Ordinal);
        Assert.Contains("name=\"adj2\" fmla=\"val 2000\"", editedXml, StringComparison.Ordinal);
        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var reprojected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/partial-preset-geometry-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reprojected.Ok, Diagnostics(reprojected));
        var editedProgram = JsonNode.Parse(reprojected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editedLeaves = editedProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray();
        Assert.Equal(1500, editedLeaves.Single(leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment" &&
            leaf!["value"]!.GetValue<long>() == 1500)!["value"]!.GetValue<long>());
        Assert.Equal(2, editedLeaves.Count(leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment"));

        var formulaSource = ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
        {
            Assert.Contains("name=\"adj1\" fmla=\"val 1000\"", xml, StringComparison.Ordinal);
            return xml.Replace("name=\"adj1\" fmla=\"val 1000\"", "name=\"adj1\" fmla=\"*/ w 1 2\"", StringComparison.Ordinal);
        });
        var formulaProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(formulaSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/partial-preset-geometry-formula.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(formulaProjection.Ok, Diagnostics(formulaProjection));
        var formulaProgram = JsonNode.Parse(formulaProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var formulaLeaves = formulaProgram["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray();
        Assert.DoesNotContain(formulaLeaves, leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment" &&
            leaf!["value"]!.GetValue<long>() == 1000);
        Assert.Contains(formulaLeaves, leaf => leaf!["kind"]!.GetValue<string>() == "presetGeometryAdjustment" &&
            leaf!["value"]!.GetValue<long>() == 2000);
    }

    [Fact]
    public void PpjSourceBoundProgramReusesOneProvenSlide()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixture = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        var sourcePage = fixture["pages"]![0]!.DeepClone().AsObject();
        sourcePage.Remove("notes");
        sourcePage.Remove("transition");
        sourcePage.Remove("animations");
        sourcePage["elements"] = new JsonArray(
            sourcePage["elements"]![0]!.DeepClone(),
            sourcePage["elements"]![1]!.DeepClone());
        fixture["assets"] = new JsonArray();
        fixture["components"] = new JsonArray();
        fixture["pages"] = new JsonArray(sourcePage);
        fixture["sections"] = new JsonArray();
        fixture["customShows"] = new JsonArray();
        fixture["comments"] = new JsonArray();

        var authored = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(fixture.ToJsonString()),
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        var sourceBytes = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var origin = state["pages"]![0]!.AsObject();
        var duplicate = origin["nativeRef"]!["capabilities"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(capability => capability["operation"]!.GetValue<string>() == "duplicate");
        Assert.Contains(duplicate["fields"]!.AsArray(), field => field!.GetValue<string>() == "pageClone");
        var cloneId = "page-source-clone";
        state["pages"]!.AsArray().Insert(1, new JsonObject
        {
            ["id"] = cloneId,
            ["role"] = "source continuation",
            ["elements"] = new JsonArray(),
            ["sourceClone"] = new JsonObject
            {
                ["page"] = origin["id"]!.GetValue<string>(),
                ["capability"] = duplicate["id"]!.GetValue<string>(),
            },
        });

        var cloned = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(cloned.Ok, Diagnostics(cloned));
        Assert.Contains(cloneId, cloned.PresentationProgram.ChangedNodeIds);
        using (var sourceStream = new MemoryStream(sourceBytes, writable: false))
        using (var sourcePackage = PresentationDocument.Open(sourceStream, false))
        using (var outputStream = new MemoryStream(cloned.File.ToByteArray(), writable: false))
        using (var outputPackage = PresentationDocument.Open(outputStream, false))
        {
            var sourceSlide = Assert.Single(OrderedSlides(sourcePackage));
            var outputSlides = OrderedSlides(outputPackage).ToArray();
            Assert.Equal(2, outputSlides.Length);
            Assert.Equal(sourceSlide.Uri, outputSlides[0].Uri);
            Assert.NotEqual(outputSlides[0].Uri, outputSlides[1].Uri);
            Assert.Equal(sourceSlide.Slide!.OuterXml, outputSlides[0].Slide!.OuterXml);
            Assert.Equal(sourceSlide.Slide!.OuterXml, outputSlides[1].Slide!.OuterXml);
        }

        var componentState = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var componentOrigin = componentState["pages"]![0]!.AsObject();
        var componentDuplicate = componentOrigin["nativeRef"]!["capabilities"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(capability => capability["operation"]!.GetValue<string>() == "duplicate");
        var retainedElement = componentOrigin["elements"]![1]!.AsObject();
        var removedSibling = componentOrigin["elements"]![0]!.AsObject();
        Assert.Contains(removedSibling["nativeRef"]!["capabilities"]!.AsArray(), item =>
            item!["operation"]!.GetValue<string>() == "delete" &&
            item["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "element"));
        var componentCloneId = "page-source-component";
        componentState["pages"]!.AsArray().Insert(1, new JsonObject
        {
            ["id"] = componentCloneId,
            ["role"] = "source component continuation",
            ["elements"] = new JsonArray(),
            ["sourceClone"] = new JsonObject
            {
                ["page"] = componentOrigin["id"]!.GetValue<string>(),
                ["capability"] = componentDuplicate["id"]!.GetValue<string>(),
                ["retainElement"] = retainedElement["id"]!.GetValue<string>(),
            },
        });

        var componentClone = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(componentState.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(componentClone.Ok, Diagnostics(componentClone));
        Assert.Contains(componentCloneId, componentClone.PresentationProgram.ChangedNodeIds);
        using (var sourceStream = new MemoryStream(sourceBytes, writable: false))
        using (var sourcePackage = PresentationDocument.Open(sourceStream, false))
        using (var componentStream = new MemoryStream(componentClone.File.ToByteArray(), writable: false))
        using (var componentPackage = PresentationDocument.Open(componentStream, false))
        {
            var sourceSlide = Assert.Single(OrderedSlides(sourcePackage));
            var componentSlides = OrderedSlides(componentPackage).ToArray();
            Assert.Equal(2, componentSlides.Length);
            Assert.Equal(sourceSlide.Slide!.OuterXml, componentSlides[0].Slide!.OuterXml);
            Assert.Equal(2, componentSlides[0].Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Count());
            Assert.Single(componentSlides[1].Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>());
            Assert.Contains("Reduce incident hours", componentSlides[1].Slide!.InnerText, StringComparison.Ordinal);
        }

        var componentProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = componentClone.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/component.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(componentProjection.Ok, Diagnostics(componentProjection));
        using (var componentJson = JsonDocument.Parse(componentProjection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var projectedComponentPage = componentJson.RootElement.GetProperty("pages")[1];
            Assert.False(projectedComponentPage.TryGetProperty("sourceClone", out _));
            var projectedComponent = Assert.Single(projectedComponentPage.GetProperty("elements").EnumerateArray());
            Assert.Equal("text", projectedComponent.GetProperty("type").GetString());
            Assert.True(projectedComponent.TryGetProperty("nativeRef", out _));
        }

        var motionState = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var motionPage = motionState["pages"]![0]!.AsObject();
        var motionCapability = motionPage["nativeRef"]!["capabilities"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(capability => capability["operation"]!.GetValue<string>() == "setAnimations");
        Assert.Contains(motionCapability["fields"]!.AsArray(), field => field!.GetValue<string>() == "animations");
        var motionTargetId = motionPage["elements"]![1]!["id"]!.GetValue<string>();
        motionPage["animations"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "source-title-reveal",
                ["target"] = motionTargetId,
                ["phase"] = "entrance",
                ["effect"] = "wipe",
                ["direction"] = "up",
                ["start"] = "afterPrevious",
                ["durationMs"] = 650,
                ["textBuild"] = "paragraph",
            },
        };
        var motion = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(sourceBytes),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(motionState.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(motion.Ok, Diagnostics(motion));
        Assert.Contains("source-title-reveal", motion.PresentationProgram.ChangedNodeIds);
        var changedMotionPart = Assert.Single(motion.PresentationProgram.ChangedParts);
        using (var sourceStream = new MemoryStream(sourceBytes, writable: false))
        using (var sourcePackage = PresentationDocument.Open(sourceStream, false))
        using (var motionStream = new MemoryStream(motion.File.ToByteArray(), writable: false))
        using (var motionPackage = PresentationDocument.Open(motionStream, false))
        {
            var sourceSlide = Assert.Single(OrderedSlides(sourcePackage));
            var motionSlide = Assert.Single(OrderedSlides(motionPackage));
            Assert.Equal(motionSlide.Uri.OriginalString.TrimStart('/'), changedMotionPart);
            Assert.Null(sourceSlide.Slide!.Timing);
            Assert.NotNull(motionSlide.Slide!.Timing);
        }
        var motionProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = motion.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/motion.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(motionProjection.Ok, Diagnostics(motionProjection));
        using (var motionJson = JsonDocument.Parse(motionProjection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var projectedMotionPage = motionJson.RootElement.GetProperty("pages")[0];
            var projectedAnimation = Assert.Single(projectedMotionPage.GetProperty("animations").EnumerateArray());
            Assert.Equal(motionTargetId, projectedAnimation.GetProperty("target").GetString());
            Assert.Equal("wipe", projectedAnimation.GetProperty("effect").GetString());
            Assert.Equal("paragraph", projectedAnimation.GetProperty("textBuild").GetString());
            Assert.Contains(projectedMotionPage.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setAnimations");
        }

        var reopened = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = cloned.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/reopened.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reopened.Ok, Diagnostics(reopened));
        using var reopenedJson = JsonDocument.Parse(reopened.PresentationProgram.ProgramJson.ToByteArray());
        Assert.Equal(2, reopenedJson.RootElement.GetProperty("pages").GetArrayLength());
        Assert.All(reopenedJson.RootElement.GetProperty("pages").EnumerateArray(), page =>
        {
            Assert.False(page.TryGetProperty("sourceClone", out _));
            Assert.NotEmpty(page.GetProperty("elements").EnumerateArray());
        });

        var continuedState = JsonNode.Parse(reopened.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var continuedPage = continuedState["pages"]![1]!.AsObject();
        var appendCapability = continuedPage["nativeRef"]!["capabilities"]!.AsArray()
            .Select(item => item!.AsObject())
            .Single(capability => capability["operation"]!.GetValue<string>() == "appendElement");
        Assert.Contains(appendCapability["fields"]!.AsArray(), field => field!.GetValue<string>() == "elements");
        var overlayId = "reviewed-source-overlay";
        continuedPage["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = overlayId,
            ["type"] = "text",
            ["role"] = "source continuation label",
            ["frame"] = new JsonObject { ["x"] = 560, ["y"] = 440, ["width"] = 300, ["height"] = 48 },
            ["text"] = "Continued through PPJ",
        });
        var iconOverlayId = "reviewed-source-icon";
        continuedPage["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = iconOverlayId,
            ["type"] = "icon",
            ["role"] = "source continuation lightbulb",
            ["frame"] = new JsonObject { ["x"] = 880, ["y"] = 440, ["width"] = 36, ["height"] = 36 },
            ["iconName"] = "fas:lightbulb",
            ["style"] = new JsonObject
            {
                ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#F2C14E" },
            },
            ["accessibility"] = new JsonObject { ["decorative"] = true },
        });
        var continued = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = cloned.File,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(continuedState.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(continued.Ok, Diagnostics(continued));
        Assert.Contains(overlayId, continued.PresentationProgram.ChangedNodeIds);
        Assert.Contains(iconOverlayId, continued.PresentationProgram.ChangedNodeIds);
        var changedOverlayPart = Assert.Single(continued.PresentationProgram.ChangedParts);
        using (var cloneStream = new MemoryStream(cloned.File.ToByteArray(), writable: false))
        using (var clonePackage = PresentationDocument.Open(cloneStream, false))
        using (var continuedStream = new MemoryStream(continued.File.ToByteArray(), writable: false))
        using (var continuedPackage = PresentationDocument.Open(continuedStream, false))
        {
            var cloneSlides = OrderedSlides(clonePackage).ToArray();
            var continuedSlides = OrderedSlides(continuedPackage).ToArray();
            Assert.Equal(continuedSlides[1].Uri.OriginalString.TrimStart('/'), changedOverlayPart);
            Assert.Equal(cloneSlides[0].Slide!.OuterXml, continuedSlides[0].Slide!.OuterXml);
            Assert.DoesNotContain("Continued through PPJ", cloneSlides[1].Slide!.InnerText);
            Assert.Contains("Continued through PPJ", continuedSlides[1].Slide!.InnerText);
            var nativeIcon = continuedSlides[1].Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "source continuation lightbulb");
            Assert.NotNull(nativeIcon.ShapeProperties!.GetFirstChild<A.CustomGeometry>());
            Assert.Equal("F2C14E", nativeIcon.ShapeProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        var continuedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = continued.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/continued.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(continuedProjection.Ok, Diagnostics(continuedProjection));
        using var continuedJson = JsonDocument.Parse(continuedProjection.PresentationProgram.ProgramJson.ToByteArray());
        var projectedOverlay = continuedJson.RootElement.GetProperty("pages")[1].GetProperty("elements")
            .EnumerateArray()
            .Single(element => element.GetProperty("type").GetString() == "text" &&
                element.GetProperty("text").GetRawText().Contains("Continued through PPJ", StringComparison.Ordinal));
        Assert.True(projectedOverlay.TryGetProperty("nativeRef", out _));
        var projectedIcon = continuedJson.RootElement.GetProperty("pages")[1].GetProperty("elements")
            .EnumerateArray()
            .Single(element => element.TryGetProperty("name", out var name) &&
                name.GetString() == "source continuation lightbulb");
        Assert.Equal("shape", projectedIcon.GetProperty("type").GetString());
        Assert.False(projectedIcon.TryGetProperty("iconName", out _));
        Assert.True(projectedIcon.TryGetProperty("nativeRef", out _));
    }

    [Fact]
    public void AuthoredPpjSmartArtBuildsNativePartsAndReprojectsWithoutEmbeddedPpj()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json"))) root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        program["assets"] = new JsonArray();
        program.Remove("sections");
        program.Remove("comments");
        program.Remove("customShows");
        program["pages"] = new JsonArray(new JsonObject
        {
            ["id"] = "smartart-page",
            ["name"] = "SmartArt",
            ["role"] = "Native SmartArt proof",
            ["elements"] = new JsonArray(new JsonObject
            {
                ["id"] = "decision-process",
                ["name"] = "Decision process",
                ["type"] = "smartArt",
                ["role"] = "Process",
                ["mode"] = "authored",
                ["layout"] = "process",
                ["frame"] = new JsonObject { ["x"] = 72, ["y"] = 120, ["width"] = 816, ["height"] = 240 },
                ["shapeStyleRef"] = "decision-band",
                ["textStyleRef"] = "body",
                ["nodeGeometry"] = new JsonObject { ["kind"] = "preset", ["preset"] = "roundRect" },
                ["connector"] = new JsonObject
                {
                    ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 1.5 },
                    ["endArrow"] = "triangle",
                },
                ["nodes"] = new JsonArray(
                    new JsonObject { ["id"] = "observe", ["text"] = "Observe" },
                    new JsonObject { ["id"] = "decide", ["text"] = "Decide" },
                    new JsonObject { ["id"] = "act", ["text"] = "Act" }),
                ["connections"] = new JsonArray(
                    new JsonObject { ["id"] = "observe-decide", ["from"] = "observe", ["to"] = "decide", ["role"] = "sequence", ["order"] = 0 },
                    new JsonObject { ["id"] = "decide-act", ["from"] = "decide", ["to"] = "act", ["role"] = "sequence", ["order"] = 1 }),
            }),
        });

        var built = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()),
            },
        });
        Assert.True(built.Ok, Diagnostics(built));
        using (var stream = new MemoryStream(built.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var slide = package.PresentationPart!.SlideParts.Single();
            Assert.Single(slide.Slide!.Descendants<P.GraphicFrame>(), frame => frame.Descendants<Dgm.RelationshipIds>().Any());
            Assert.Single(slide.DiagramDataParts);
            Assert.Single(slide.DiagramLayoutDefinitionParts);
            Assert.Single(slide.DiagramStyleParts);
            Assert.Single(slide.DiagramColorsParts);
            Assert.Single(slide.Parts, pair => pair.OpenXmlPart is DiagramPersistLayoutPart);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));
        }

        var nativeOnly = RemoveEmbeddedPpj(built.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/native-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        using var json = JsonDocument.Parse(projected.PresentationProgram.ProgramJson.ToByteArray());
        var smartArt = Assert.Single(json.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray());
        Assert.Equal("smartArt", smartArt.GetProperty("type").GetString());
        Assert.Equal("source-bound", smartArt.GetProperty("mode").GetString());
        Assert.Equal("process", smartArt.GetProperty("layout").GetString());
        Assert.Equal(["observe", "decide", "act"], smartArt.GetProperty("nodes").EnumerateArray().Select(node => node.GetProperty("id").GetString()));
        Assert.Equal(["observe-decide", "decide-act"], smartArt.GetProperty("connections").EnumerateArray().Select(connection => connection.GetProperty("id").GetString()));

        var noOp = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = projected.PresentationProgram.ProgramJson,
            },
        });
        Assert.True(noOp.Ok, Diagnostics(noOp));
        Assert.Empty(noOp.PresentationProgram.ChangedParts);

        var locallyEditedState = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var locallyEditedSmartArt = locallyEditedState["pages"]![0]!["elements"]![0]!.AsObject();
        locallyEditedSmartArt["nodes"]![1]!["text"] = "Decide better";
        locallyEditedSmartArt["connections"]![0]!["to"] = "act";
        locallyEditedSmartArt["connections"]![0]!["role"] = "association";
        locallyEditedSmartArt["frame"]!["x"] = 84;
        locallyEditedSmartArt["frame"]!["width"] = 792;
        var locallyEdited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(locallyEditedState.ToJsonString()),
            },
        });
        Assert.True(locallyEdited.Ok, Diagnostics(locallyEdited));
        using (var sourceStream = new MemoryStream(nativeOnly))
        using (var sourcePackage = PresentationDocument.Open(sourceStream, false))
        using (var editedStream = new MemoryStream(locallyEdited.File.ToByteArray()))
        using (var editedPackage = PresentationDocument.Open(editedStream, false))
        {
            var sourceSlide = Assert.Single(sourcePackage.PresentationPart!.SlideParts);
            var editedSlide = Assert.Single(editedPackage.PresentationPart!.SlideParts);
            Assert.Equal(Assert.Single(sourceSlide.DiagramLayoutDefinitionParts).LayoutDefinition!.OuterXml,
                Assert.Single(editedSlide.DiagramLayoutDefinitionParts).LayoutDefinition!.OuterXml);
            Assert.Equal(Assert.Single(sourceSlide.DiagramStyleParts).StyleDefinition!.OuterXml,
                Assert.Single(editedSlide.DiagramStyleParts).StyleDefinition!.OuterXml);
            Assert.Equal(Assert.Single(sourceSlide.DiagramColorsParts).ColorsDefinition!.OuterXml,
                Assert.Single(editedSlide.DiagramColorsParts).ColorsDefinition!.OuterXml);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(editedPackage));
        }
        var locallyEditedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = locallyEdited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/locally-edited-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(locallyEditedProjection.Ok, Diagnostics(locallyEditedProjection));
        using (var locallyEditedJson = JsonDocument.Parse(locallyEditedProjection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var local = Assert.Single(locallyEditedJson.RootElement.GetProperty("pages")[0]
                .GetProperty("elements").EnumerateArray());
            Assert.Contains("Decide better", local.GetProperty("nodes")[1].GetProperty("text").GetRawText(), StringComparison.Ordinal);
            Assert.Equal("act", local.GetProperty("connections")[0].GetProperty("to").GetString());
            Assert.Equal("association", local.GetProperty("connections")[0].GetProperty("role").GetString());
            Assert.Equal(84, local.GetProperty("frame").GetProperty("x").GetDouble());
            Assert.Equal(792, local.GetProperty("frame").GetProperty("width").GetDouble());
        }

        var renamedState = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        renamedState["pages"]![0]!["name"] = "Renamed around native SmartArt";
        var renamed = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(renamedState.ToJsonString()),
            },
        });
        Assert.True(renamed.Ok, Diagnostics(renamed));
        Assert.Equal(["ppt/slides/slide1.xml"], renamed.PresentationProgram.ChangedParts);
        var renamedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = renamed.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/renamed-native-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(renamedProjection.Ok, Diagnostics(renamedProjection));
        using (var renamedJson = JsonDocument.Parse(renamedProjection.PresentationProgram.ProgramJson.ToByteArray()))
            Assert.Equal(
                "smartArt",
                Assert.Single(renamedJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray())
                    .GetProperty("type").GetString());

        var detachedState = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var detachable = detachedState["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Contains(detachable["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "detachSmartArt");
        detachable["detachToShapes"] = true;
        var detached = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(detachedState.ToJsonString()),
            },
        });
        Assert.True(detached.Ok, Diagnostics(detached));
        Assert.Contains(detached.Diagnostics, diagnostic =>
            diagnostic.Severity == DiagnosticSeverity.Warning && diagnostic.Code == "ppj.smartArt.detachedToShapes");
        using (var detachedStream = new MemoryStream(detached.File.ToByteArray()))
        using (var detachedPackage = PresentationDocument.Open(detachedStream, false))
        {
            var detachedSlide = Assert.Single(detachedPackage.PresentationPart!.SlideParts);
            Assert.DoesNotContain(detachedSlide.Slide!.Descendants<P.GraphicFrame>(), frame => frame.Descendants<Dgm.RelationshipIds>().Any());
            Assert.True(detachedSlide.Slide.Descendants<P.GroupShape>().Any());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(detachedPackage));
        }
        var detachedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = detached.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/detached-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(detachedProjection.Ok, Diagnostics(detachedProjection));
        using (var detachedJson = JsonDocument.Parse(detachedProjection.PresentationProgram.ProgramJson.ToByteArray()))
            Assert.Equal("group", Assert.Single(detachedJson.RootElement.GetProperty("pages")[0]
                .GetProperty("elements").EnumerateArray()).GetProperty("type").GetString());

        var definitionBytes = Encoding.UTF8.GetBytes("""
            {
              "schema": "office-kit/smartart-definition/v1",
              "layout": { "id": "decision-process", "profile": "process" },
              "style": { "id": "basic" },
              "colors": { "id": "accent" }
            }
            """);
        var definitionSha256 = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(definitionBytes)).ToLowerInvariant();
        var authoredSmartArt = program["pages"]![0]!["elements"]![0]!.AsObject();
        authoredSmartArt.Remove("layout");
        authoredSmartArt["definitionAsset"] = "decision-process-definition";
        program["assets"] = new JsonArray(new JsonObject
        {
            ["id"] = "decision-process-definition",
            ["uri"] = "deck.assets/smartart/decision-process.json",
            ["mimeType"] = "application/vnd.officekit.smartart-definition+json",
            ["sha256"] = definitionSha256,
            ["rights"] = new JsonObject { ["status"] = "internal" },
            ["accessibility"] = new JsonObject { ["decorative"] = true },
        });
        var definitionRequest = new CodecRequest
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
                        Id = "decision-process-definition",
                        FileName = "decision-process.json",
                        ContentType = "application/vnd.officekit.smartart-definition+json",
                        Data = ByteString.CopyFrom(definitionBytes),
                        Sha256 = definitionSha256,
                    },
                },
            },
        };
        var definitionBuilt = Invoke(definitionRequest);
        Assert.True(definitionBuilt.Ok, Diagnostics(definitionBuilt));
        using var definitionStream = new MemoryStream(definitionBuilt.File.ToByteArray());
        using var definitionPackage = PresentationDocument.Open(definitionStream, false);
        var definitionSlide = Assert.Single(definitionPackage.PresentationPart!.SlideParts);
        Assert.Contains(
            "urn:officekit:smartart:v1:layout:process",
            Assert.Single(definitionSlide.DiagramLayoutDefinitionParts).LayoutDefinition!.OuterXml,
            StringComparison.Ordinal);

        var executableDefinitionBytes = Encoding.UTF8.GetBytes("""
            {
              "schema": "office-kit/smartart-definition/v1",
              "layout": {
                "id": "custom-process",
                "profile": "process",
                "operators": [
                  { "id": "custom-algorithm", "kind": "algorithm", "input": "nodes", "arguments": { "placement": "square-grid" } },
                  { "id": "custom-rule", "kind": "rule", "arguments": { "columns": 2, "reverse": true } },
                  { "id": "custom-gap", "kind": "constraint", "arguments": { "gapPoints": 20 } }
                ]
              },
              "style": { "id": "emphasis" },
              "colors": { "id": "cool" }
            }
            """);
        var executableDefinitionSha256 = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(executableDefinitionBytes)).ToLowerInvariant();
        program["assets"]![0]!["sha256"] = executableDefinitionSha256;
        definitionRequest.PresentationProgram.ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString());
        definitionRequest.PresentationProgram.Assets[0].Data = ByteString.CopyFrom(executableDefinitionBytes);
        definitionRequest.PresentationProgram.Assets[0].Sha256 = executableDefinitionSha256;
        var executableDefinition = Invoke(definitionRequest);
        Assert.True(executableDefinition.Ok, Diagnostics(executableDefinition));
        using (var executableStream = new MemoryStream(executableDefinition.File.ToByteArray()))
        using (var executablePackage = PresentationDocument.Open(executableStream, false))
        {
            var executableSlide = Assert.Single(executablePackage.PresentationPart!.SlideParts);
            Assert.Contains(
                "urn:officekit:smartart:v1:quickstyle:emphasis",
                Assert.Single(executableSlide.DiagramStyleParts).StyleDefinition!.OuterXml,
                StringComparison.Ordinal);
            Assert.Contains(
                "urn:officekit:smartart:v1:colors:cool",
                Assert.Single(executableSlide.DiagramColorsParts).ColorsDefinition!.OuterXml,
                StringComparison.Ordinal);
            var drawing = XDocument.Parse(Assert.Single(
                executableSlide.Parts,
                pair => pair.OpenXmlPart is DiagramPersistLayoutPart).OpenXmlPart.RootElement!.OuterXml);
            XNamespace dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var offsets = drawing.Descendants(dsp + "sp").ToDictionary(
                shape => shape.Element(dsp + "nvSpPr")!.Element(dsp + "cNvPr")!.Attribute("name")!.Value,
                shape => shape.Element(dsp + "spPr")!.Element(a + "xfrm")!.Element(a + "off")!,
                StringComparer.Ordinal);
            Assert.Equal((0L, 0L), ((long)offsets["act"].Attribute("x")!, (long)offsets["act"].Attribute("y")!));
            Assert.Equal((5_308_600L, 0L), ((long)offsets["decide"].Attribute("x")!, (long)offsets["decide"].Attribute("y")!));
            Assert.Equal((0L, 1_651_000L), ((long)offsets["observe"].Attribute("x")!, (long)offsets["observe"].Attribute("y")!));
        }
        var projectedDefinition = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(RemoveEmbeddedPpj(executableDefinition.File.ToByteArray())),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-definition-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projectedDefinition.Ok, Diagnostics(projectedDefinition));
        var projectedDefinitionAsset = Assert.Single(projectedDefinition.PresentationProgram.Assets,
            asset => asset.ContentType == "application/vnd.officekit.smartart-definition+json");
        Assert.Equal(executableDefinitionSha256, projectedDefinitionAsset.Sha256);
        Assert.Equal(executableDefinitionBytes, projectedDefinitionAsset.Data.ToByteArray());
        using (var projectedDefinitionJson = JsonDocument.Parse(projectedDefinition.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var projectedElement = Assert.Single(projectedDefinitionJson.RootElement.GetProperty("pages")[0]
                .GetProperty("elements").EnumerateArray());
            Assert.False(projectedElement.TryGetProperty("layout", out _));
            var projectedAssetId = projectedElement.GetProperty("definitionAsset").GetString();
            var declaration = Assert.Single(projectedDefinitionJson.RootElement.GetProperty("assets").EnumerateArray(),
                asset => asset.GetProperty("id").GetString() == projectedAssetId);
            Assert.Equal(executableDefinitionSha256, declaration.GetProperty("sha256").GetString());
        }

        var unsupportedDefinitionBytes = Encoding.UTF8.GetBytes("""
            {
              "schema": "office-kit/smartart-definition/v1",
              "layout": {
                "id": "custom-process",
                "profile": "process",
                "operators": [{ "id": "unsupported-shape", "kind": "shape", "arguments": { "geometry": "hexagon" } }]
              },
              "style": { "id": "basic" },
              "colors": { "id": "accent" }
            }
            """);
        var unsupportedDefinitionSha256 = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(unsupportedDefinitionBytes)).ToLowerInvariant();
        program["assets"]![0]!["sha256"] = unsupportedDefinitionSha256;
        definitionRequest.PresentationProgram.ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString());
        definitionRequest.PresentationProgram.Assets[0].Data = ByteString.CopyFrom(unsupportedDefinitionBytes);
        definitionRequest.PresentationProgram.Assets[0].Sha256 = unsupportedDefinitionSha256;
        var unsupportedDefinition = Invoke(definitionRequest);
        Assert.False(unsupportedDefinition.Ok);
        var unsupportedDiagnostic = Assert.Single(unsupportedDefinition.Diagnostics);
        Assert.Equal("unsupported_ppj_compile_feature", unsupportedDiagnostic.Code);
        Assert.Equal("$.assets[decision-process-definition].layout.operators[0].kind", unsupportedDiagnostic.SourcePath);
    }

    [Fact]
    public void SourceBoundPictureSmartArtCanReplaceAnExistingNodeAsset()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(
            root!.FullName,
            "test",
            "fixtures",
            "presentation",
            "evidence-ledger-canonical.ppj")))!.AsObject();
        program.Remove("sections");
        program.Remove("comments");
        program.Remove("customShows");

        var originalBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        var secondBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=");
        var originalSha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(originalBytes)).ToLowerInvariant();
        var secondSha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(secondBytes)).ToLowerInvariant();
        var originalAssetId = "smartart-image-original";
        var secondAssetId = "smartart-image-second";
        program["assets"] = new JsonArray(
            new JsonObject
            {
                ["id"] = originalAssetId,
                ["uri"] = "deck.assets/smartart/original.png",
                ["mimeType"] = "image/png",
                ["sha256"] = originalSha,
                ["rights"] = new JsonObject { ["status"] = "internal" },
                ["accessibility"] = new JsonObject { ["decorative"] = false },
            },
            new JsonObject
            {
                ["id"] = secondAssetId,
                ["uri"] = "deck.assets/smartart/second.png",
                ["mimeType"] = "image/png",
                ["sha256"] = secondSha,
                ["rights"] = new JsonObject { ["status"] = "internal" },
                ["accessibility"] = new JsonObject { ["decorative"] = false },
            });
        program["pages"] = new JsonArray(new JsonObject
        {
            ["id"] = "smartart-picture-page",
            ["name"] = "SmartArt pictures",
            ["role"] = "SmartArt picture replacement proof",
            ["elements"] = new JsonArray(new JsonObject
            {
                ["id"] = "picture-process",
                ["name"] = "Picture process",
                ["type"] = "smartArt",
                ["mode"] = "authored",
                ["layout"] = "picture",
                ["frame"] = new JsonObject { ["x"] = 72, ["y"] = 120, ["width"] = 816, ["height"] = 360 },
                ["shapeStyleRef"] = "decision-band",
                ["textStyleRef"] = "body",
                ["nodes"] = new JsonArray(
                    new JsonObject
                    {
                        ["id"] = "observe",
                        ["text"] = "Observe",
                        ["asset"] = originalAssetId,
                        ["image"] = new JsonObject
                        {
                            ["fit"] = "stretch",
                            ["crop"] = new JsonObject { ["left"] = 0.1, ["right"] = 0.2 },
                            ["opacity"] = 0.65,
                        },
                    },
                    new JsonObject { ["id"] = "decide", ["text"] = "Decide", ["asset"] = secondAssetId }),
            }),
        });

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
                    new Asset { Id = originalAssetId, FileName = "original.png", ContentType = "image/png", Data = ByteString.CopyFrom(originalBytes), Sha256 = originalSha },
                    new Asset { Id = secondAssetId, FileName = "second.png", ContentType = "image/png", Data = ByteString.CopyFrom(secondBytes), Sha256 = secondSha },
                },
            },
        });
        Assert.True(authored.Ok, Diagnostics(authored));
        var nativeOnly = RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/picture-smartart.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        var state = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var smartArt = state["pages"]![0]!["elements"]![0]!.AsObject();
        Assert.Equal("picture", smartArt["layout"]!.GetValue<string>());
        Assert.Contains(smartArt["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setSmartArtImage" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "smartArt.nodes[].asset"));
        Assert.Contains(smartArt["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setSmartArtImagePaint" &&
            capability["fields"]!.AsArray().Any(field => field!.GetValue<string>() == "smartArt.nodes[].image"));
        var projectedNodes = smartArt["nodes"]!.AsArray().Select(node => node!.AsObject()).ToArray();
        var originalProjectedAsset = projectedNodes[0]["asset"]!.GetValue<string>();
        var secondProjectedAsset = projectedNodes[1]["asset"]!.GetValue<string>();
        Assert.NotEqual(originalProjectedAsset, secondProjectedAsset);
        var projectedImage = projectedNodes[0]["image"]!.AsObject();
        Assert.Equal("stretch", projectedImage["fit"]!.GetValue<string>());
        Assert.Equal(0.1, projectedImage["crop"]!["left"]!.GetValue<double>(), 6);
        Assert.Equal(0.2, projectedImage["crop"]!["right"]!.GetValue<double>(), 6);
        Assert.Equal(0.65, projectedImage["opacity"]!.GetValue<double>(), 6);

        var replacementBytes = secondBytes;
        var replacementSha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(replacementBytes)).ToLowerInvariant();
        var replacementAssetId = "smartart-image-replacement";
        state["assets"]!.AsArray().Add(new JsonObject
        {
            ["id"] = replacementAssetId,
            ["uri"] = "deck.assets/smartart/replacement.png",
            ["mimeType"] = "image/png",
            ["sha256"] = replacementSha,
            ["rights"] = new JsonObject { ["status"] = "internal" },
            ["accessibility"] = new JsonObject { ["decorative"] = false },
        });
        smartArt["nodes"]![0]!["asset"] = replacementAssetId;
        var editedImage = smartArt["nodes"]![0]!["image"]!.AsObject();
        editedImage["crop"]!.AsObject()["left"] = 0.2;
        editedImage["opacity"] = 0.35;

        var edited = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeOnly),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(state.ToJsonString()),
                Assets =
                {
                    new Asset { Id = replacementAssetId, FileName = "replacement.png", ContentType = "image/png", Data = ByteString.CopyFrom(replacementBytes), Sha256 = replacementSha },
                },
            },
        });
        Assert.True(edited.Ok, Diagnostics(edited));
        Assert.Contains(edited.PresentationProgram.ChangedParts, path => path.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(edited.PresentationProgram.ChangedParts, path => path.Contains("diagrams", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(edited.PresentationProgram.ChangedParts, path =>
            path.Contains("diagrams/_rels", StringComparison.OrdinalIgnoreCase) &&
            path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(edited.PresentationProgram.ChangedParts, path => path.Contains("media", StringComparison.OrdinalIgnoreCase));

        using (var stream = new MemoryStream(edited.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package));

        var rebound = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = edited.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/picture-smartart-edited.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(rebound.Ok, Diagnostics(rebound));
        using var reboundJson = JsonDocument.Parse(rebound.PresentationProgram.ProgramJson.ToByteArray());
        var reboundSmartArt = reboundJson.RootElement.GetProperty("pages")[0].GetProperty("elements")[0];
        var reboundAssetId = reboundSmartArt.GetProperty("nodes")[0].GetProperty("asset").GetString()!;
        var reboundAsset = reboundJson.RootElement.GetProperty("assets").EnumerateArray()
            .Single(asset => asset.GetProperty("id").GetString() == reboundAssetId);
        Assert.Equal(replacementSha, reboundAsset.GetProperty("sha256").GetString());
        var reboundImage = reboundSmartArt.GetProperty("nodes")[0].GetProperty("image");
        Assert.Equal("stretch", reboundImage.GetProperty("fit").GetString());
        Assert.Equal(0.2, reboundImage.GetProperty("crop").GetProperty("left").GetDouble(), 6);
        Assert.Equal(0.2, reboundImage.GetProperty("crop").GetProperty("right").GetDouble(), 6);
        Assert.Equal(0.35, reboundImage.GetProperty("opacity").GetDouble(), 6);
    }

    [Fact]
    public void PpjV1CompilesCanonicalPresentationProgramDeterministically()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json")))
            root = root.Parent;
        Assert.NotNull(root);
        var fixtureDirectory = Path.Combine(root!.FullName, "test", "fixtures", "presentation");
        var authoredProgram = JsonNode.Parse(File.ReadAllBytes(Path.Combine(fixtureDirectory, "evidence-ledger-canonical.ppj")))!.AsObject();
        authoredProgram["design"]!["masters"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "master-evidence",
                ["name"] = "Evidence master",
                ["background"] = new JsonObject { ["type"] = "solid", ["color"] = "#F8F6EF" },
                ["textStyles"] = new JsonObject
                {
                    ["title"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["level"] = 0,
                            ["alignment"] = "left",
                            ["defaultText"] = new JsonObject
                            {
                                ["font"] = "sans",
                                ["size"] = 28,
                                ["bold"] = true,
                                ["color"] = "#14324A",
                            },
                        },
                    },
                },
                ["placeholders"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "master-title",
                        ["name"] = "Master title",
                        ["placeholderType"] = "title",
                        ["index"] = 1,
                        ["frame"] = new JsonObject { ["x"] = 48, ["y"] = 36, ["width"] = 624, ["height"] = 64 },
                    },
                },
            },
        };
        authoredProgram["design"]!["layouts"] = new JsonArray
        {
            new JsonObject
            {
                ["id"] = "layout-evidence",
                ["name"] = "Evidence layout",
                ["master"] = "master-evidence",
                ["layoutType"] = "titleOnly",
                ["placeholders"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "layout-title",
                        ["name"] = "Layout title",
                        ["placeholderType"] = "title",
                        ["index"] = 1,
                        ["frame"] = new JsonObject { ["x"] = 48, ["y"] = 36, ["width"] = 624, ["height"] = 64 },
                    },
                },
            },
        };
        foreach (var page in authoredProgram["pages"]!.AsArray()) page!["layout"] = "layout-evidence";
        authoredProgram["pages"]![0]!["background"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "linear",
            ["angle"] = 24,
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#F8F6EF" },
                new JsonObject { ["offset"] = 0.62, ["color"] = "#DCEFEA", ["opacity"] = 0.86 },
                new JsonObject { ["offset"] = 1, ["color"] = "#FFFFFF" },
            },
        };
        var authoredTitle = authoredProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-title");
        authoredTitle["hidden"] = true;
        authoredTitle["locked"] = true;
        var authoredImage = authoredProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-mark");
        authoredImage["fit"] = "tile";
        authoredImage["mask"] = new JsonObject
        {
            ["kind"] = "preset",
            ["preset"] = "roundRect",
            ["adjustments"] = new JsonArray(24000),
        };
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "custom-mask-image-main",
            ["type"] = "image",
            ["role"] = "irregular editorial crop",
            ["frame"] = new JsonObject { ["x"] = 520, ["y"] = 300, ["width"] = 80, ["height"] = 80 },
            ["asset"] = "evidence-mark",
            ["fit"] = "cover",
            ["mask"] = new JsonObject
            {
                ["kind"] = "custom",
                ["viewBox"] = new JsonObject { ["x"] = 0, ["y"] = 0, ["width"] = 160, ["height"] = 160 },
                ["paths"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["fill"] = true,
                        ["stroke"] = false,
                        ["commands"] = new JsonArray
                        {
                            new JsonObject { ["op"] = "moveTo", ["x"] = 20, ["y"] = 0 },
                            new JsonObject { ["op"] = "lineTo", ["x"] = 160, ["y"] = 0 },
                            new JsonObject { ["op"] = "lineTo", ["x"] = 140, ["y"] = 160 },
                            new JsonObject { ["op"] = "lineTo", ["x"] = 0, ["y"] = 120 },
                            new JsonObject { ["op"] = "close" },
                        },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Evidence mark clipped by an irregular native path.",
            },
        });
        authoredProgram["pages"]![1]!["background"] = new JsonObject
        {
            ["type"] = "image",
            ["asset"] = "evidence-mark",
            ["fit"] = "cover",
            ["opacity"] = 0.72,
        };
        var authoredImageFillShape = authoredProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "decision-flow-start");
        authoredImageFillShape["style"]!["fill"] = new JsonObject
        {
            ["type"] = "image",
            ["asset"] = "evidence-mark",
            ["fit"] = "contain",
            ["opacity"] = 0.66,
        };
        authoredImageFillShape["style"]!["stroke"] = new JsonObject
        {
            ["color"] = "#16324F",
            ["width"] = 1.5,
            ["opacity"] = 0.8,
        };
        authoredImageFillShape["style"]!["shadow"] = new JsonObject
        {
            ["color"] = "#16324F",
            ["blur"] = 3,
            ["distance"] = 2,
            ["angle"] = 45,
            ["opacity"] = 0.6,
        };
        authoredImageFillShape["style"]!["opacity"] = 0.5;
        authoredTitle["style"] = new JsonObject
        {
            ["wrap"] = "square",
            ["columnDirection"] = "left-to-right",
            ["verticalText"] = "horizontal",
        };
        var authoredParagraph = authoredTitle["text"]!["paragraphs"]![0]!.AsObject();
        authoredParagraph["style"] = new JsonObject
        {
            ["lineSpacingMultiplier"] = 1.1,
            ["spaceAfterMultiplier"] = 0.2,
            ["defaultText"] = new JsonObject
            {
                ["gradient"] = new JsonObject
                {
                    ["kind"] = "radial",
                    ["stops"] = new JsonArray
                    {
                        new JsonObject { ["offset"] = 0, ["color"] = "#FFFFFF" },
                        new JsonObject { ["offset"] = 1, ["color"] = "#DCEFEA", ["opacity"] = 0.7 },
                    },
                },
                ["shadow"] = new JsonObject
                {
                    ["color"] = "#16324F80",
                    ["blur"] = 2,
                    ["distance"] = 1,
                    ["angle"] = 45,
                },
            },
            ["bullet"] = new JsonObject
            {
                ["type"] = "character",
                ["character"] = "•",
                ["color"] = new JsonObject { ["token"] = "signal", ["alpha"] = 0.5 },
            },
        };
        var authoredRunStyle = authoredParagraph["runs"]![0]!["style"]!.AsObject();
        authoredRunStyle["fontFamilyEastAsia"] = "Arial";
        authoredRunStyle.Remove("color");
        authoredRunStyle["gradient"] = new JsonObject
        {
            ["kind"] = "linear",
            ["angle"] = 18,
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#16324F" },
                new JsonObject { ["offset"] = 1, ["color"] = "#0B8F8F", ["opacity"] = 0.8 },
            },
        };
        authoredRunStyle["shadow"] = new JsonObject
        {
            ["color"] = "#16324F66",
            ["blur"] = 3,
            ["distance"] = 1.5,
            ["angle"] = 90,
        };
        authoredRunStyle["highlight"] = "#FFF2CC";
        authoredRunStyle["underline"] = "single";
        authoredRunStyle["strike"] = false;
        authoredRunStyle["kerning"] = 12;
        authoredRunStyle["letterSpacing"] = 0.4;
        authoredRunStyle["baseline"] = 0;
        authoredRunStyle["capitalization"] = "none";
        authoredRunStyle["language"] = "zh-CN";
        var authoredCustomShape = authoredProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-rule");
        authoredCustomShape["geometry"] = new JsonObject
        {
            ["kind"] = "custom",
            ["viewBox"] = new JsonObject { ["x"] = 10, ["y"] = 20, ["width"] = 100, ["height"] = 100 },
            ["paths"] = new JsonArray
            {
                new JsonObject
                {
                    ["fill"] = true,
                    ["stroke"] = true,
                    ["commands"] = new JsonArray
                    {
                        new JsonObject { ["op"] = "moveTo", ["x"] = 10, ["y"] = 70 },
                        new JsonObject { ["op"] = "arcTo", ["radiusX"] = 50, ["radiusY"] = 50, ["startAngle"] = 180, ["sweepAngle"] = 180 },
                        new JsonObject { ["op"] = "lineTo", ["x"] = 110, ["y"] = 75 },
                        new JsonObject { ["op"] = "quadraticTo", ["x1"] = 110, ["y1"] = 100, ["x"] = 60, ["y"] = 120 },
                        new JsonObject { ["op"] = "cubicTo", ["x1"] = 35, ["y1"] = 120, ["x2"] = 10, ["y2"] = 95, ["x"] = 10, ["y"] = 70 },
                        new JsonObject { ["op"] = "close" },
                    },
                },
            },
        };
        authoredCustomShape["style"]!["stroke"] = new JsonObject
        {
            ["color"] = "#0B8F8F",
            ["width"] = 2,
            ["opacity"] = 0.42,
            ["dash"] = "dash",
            ["cap"] = "round",
            ["join"] = "round",
        };
        authoredCustomShape["style"]!["fill"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "radial",
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#F2C14E" },
                new JsonObject { ["offset"] = 1, ["color"] = "#0B8F8F", ["opacity"] = 0.35 },
            },
        };
        authoredCustomShape["style"]!["opacity"] = 0.5;
        var authoredConnector = authoredProgram["pages"]!
            .AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "decision-flow-link");
        authoredConnector["stroke"]!["opacity"] = 0.58;
        var authoredTable = authoredProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "method-table-main");
        authoredTable["frame"]!["rotation"] = -4;
        authoredTable["frame"]!["flipV"] = true;
        var authoredTableStyle = authoredProgram["design"]!["styles"]!["table"]![0]!["style"]!.AsObject();
        authoredTableStyle["headerRows"] = 2;
        authoredTableStyle["headerCellFill"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = "#B7DEE8",
            ["opacity"] = 0.64,
        };
        authoredTableStyle["headerTextStyle"] = new JsonObject
        {
            ["verticalAlignment"] = "middle",
            ["defaultText"] = new JsonObject
            {
                ["font"] = "sans",
                ["size"] = 11,
                ["bold"] = true,
                ["color"] = new JsonObject { ["token"] = "signal", ["alpha"] = 0.88 },
            },
        };
        authoredTableStyle["defaultCellFill"] = new JsonObject
        {
            ["type"] = "image",
            ["asset"] = "evidence-mark",
            ["fit"] = "cover",
            ["opacity"] = 0.22,
        };
        authoredTableStyle["cellStyle"] = new JsonObject
        {
            ["borders"] = new JsonObject
            {
                ["left"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 0.5 },
            },
        };
        authoredTableStyle["bodyStyles"] = new JsonArray
        {
            new JsonObject
            {
                ["borders"] = new JsonObject
                {
                    ["right"] = new JsonObject { ["color"] = "#C1121F", ["width"] = 0.75 },
                },
            },
        };
        authoredTableStyle["firstRowStyle"] = new JsonObject
        {
            ["borders"] = new JsonObject
            {
                ["top"] = new JsonObject { ["color"] = "#F2C14E", ["width"] = 1 },
            },
        };
        authoredTableStyle["lastRowStyle"] = new JsonObject
        {
            ["borders"] = new JsonObject
            {
                ["bottom"] = new JsonObject { ["color"] = "#16324F", ["width"] = 2 },
            },
        };
        authoredTableStyle["firstColumnStyle"] = new JsonObject
        {
            ["textStyle"] = new JsonObject
            {
                ["defaultText"] = new JsonObject { ["italic"] = true },
            },
        };
        authoredTableStyle["rowOverColumn"] = true;
        var authoredHeaderCell = authoredTable["rows"]![0]!["cells"]![0]!.AsObject();
        authoredHeaderCell["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#DCEFEA", ["opacity"] = 0.8 };
        authoredHeaderCell["textStyle"] = new JsonObject
        {
            ["defaultText"] = new JsonObject { ["color"] = "#C1121F" },
        };
        authoredTable["rows"]![0]!["cells"]![1]!["fill"] = new JsonObject
        {
            ["type"] = "image",
            ["asset"] = "evidence-mark",
            ["fit"] = "tile",
            ["opacity"] = 0.55,
        };
        authoredHeaderCell["borders"] = new JsonObject
        {
            ["bottom"] = new JsonObject
            {
                ["color"] = new JsonObject { ["token"] = "signal", ["alpha"] = 0.65 },
                ["width"] = 1.5,
                ["dash"] = "solid",
                ["cap"] = "round",
                ["join"] = "round",
            },
        };
        var authoredTableTextStyle = authoredTableStyle["defaultTextStyle"]!["defaultText"]!.AsObject();
        authoredTableTextStyle["color"] = new JsonObject { ["token"] = "ink", ["alpha"] = 0.72 };
        var authoredChart = authoredProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-chart-main");
        authoredChart["title"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["id"] = "chart-title",
                    ["runs"] = new JsonArray
                    {
                        new JsonObject { ["id"] = "chart-title-label", ["text"] = "Measured profile: " },
                        new JsonObject
                        {
                            ["id"] = "chart-title-delta",
                            ["text"] = "−38% incidents",
                            ["style"] = new JsonObject
                            {
                                ["bold"] = true,
                                ["color"] = "#C1121F",
                                ["fontFamilyEastAsia"] = "Noto Serif CJK SC",
                            },
                        },
                    },
                },
            },
        };
        authoredChart["frame"]!["rotation"] = 6;
        authoredChart["frame"]!["flipH"] = true;
        authoredChart["xAxis"] = new JsonObject
        {
            ["visible"] = true,
            ["reverse"] = true,
            ["title"] = "Half-year",
            ["tickLabelInterval"] = 1,
            ["axisLine"] = new JsonObject { ["color"] = "#16324F", ["width"] = 1.25, ["dash"] = "dash" },
            ["axisLineArrow"] = new JsonObject { ["start"] = "open", ["end"] = "triangle" },
            ["gridLine"] = false,
            ["textStyle"] = new JsonObject
            {
                ["fontSize"] = 9,
                ["fontFamily"] = "Aptos",
                ["fontFamilyEastAsia"] = "Noto Sans CJK SC",
                ["bold"] = true,
                ["italic"] = false,
                ["color"] = "#16324FCC",
            },
            ["titleTextStyle"] = new JsonObject
            {
                ["fontSize"] = 11,
                ["fontFamily"] = "Georgia",
                ["bold"] = true,
                ["color"] = "#0B8F8F",
            },
        };
        authoredChart["yAxis"] = new JsonObject
        {
            ["visible"] = true,
            ["title"] = "Incident hours",
            ["numberFormat"] = "0",
            ["min"] = 0,
            ["max"] = 80,
            ["majorUnit"] = 20,
            ["gridLine"] = new JsonObject { ["color"] = "#DCEFEA", ["width"] = 0.75, ["dash"] = "dot" },
            ["textStyle"] = new JsonObject { ["fontSize"] = 9 },
        };
        authoredChart["secondaryXAxis"] = new JsonObject { ["visible"] = false };
        authoredChart["secondaryYAxis"] = new JsonObject
        {
            ["visible"] = true,
            ["title"] = "Workload index",
            ["min"] = 90,
            ["max"] = 130,
            ["majorUnit"] = 10,
        };
        var invalidAxisArrowProgram = authoredProgram.DeepClone().AsObject();
        invalidAxisArrowProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-chart-main")["xAxis"]!["axisLine"] = false;
        var invalidAxisArrow = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidAxisArrowProgram.ToJsonString()));
        Assert.False(invalidAxisArrow.IsValid);
        Assert.Contains(invalidAxisArrow.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.axisArrowHiddenLine");
        var authoredChartSeries = authoredChart["data"]!["series"]![1]!.AsObject();
        authoredChartSeries["color"] = "#F2C14E80";
        authoredChartSeries["marker"] = new JsonObject
        {
            ["symbol"] = "circle",
            ["size"] = 8,
            ["fill"] = "#F2C14E80",
            ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 1 },
        };
        authoredChartSeries["trendlines"] = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "linear",
                ["name"] = "Growth trend",
                ["displayEquation"] = true,
                ["displayRSquared"] = true,
                ["stroke"] = new JsonObject { ["color"] = "#F2C14E", ["width"] = 1.25, ["dash"] = "dash" },
            },
        };
        authoredChartSeries["errorBars"] = new JsonObject
        {
            ["direction"] = "y",
            ["type"] = "both",
            ["valueType"] = "standard-error",
            ["noEndCap"] = true,
            ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 0.75 },
        };
        authoredChartSeries["dataLabels"] = new JsonObject
        {
            ["showValue"] = true,
            ["numberFormat"] = "0.0",
            ["textStyle"] = new JsonObject { ["fontSize"] = 9, ["color"] = "#16324F" },
            ["points"] = new JsonArray
            {
                new JsonObject { ["index"] = 2, ["showValue"] = false },
                new JsonObject { ["index"] = 7, ["showValue"] = true, ["position"] = "top", ["numberFormat"] = "0.0x" },
            },
        };
        var authoredChartStyle = authoredProgram["design"]!["styles"]!["chart"]!.AsArray()
            .Select(style => style!.AsObject())
            .Single(style => style["id"]!.GetValue<string>() == "evidence-chart")["style"]!.AsObject();
        authoredChartStyle.Remove("showCategoryAxis");
        authoredChartStyle.Remove("showValueAxis");
        authoredChartStyle.Remove("showDataLabels");
        authoredChartStyle.Remove("dataLabelPosition");
        authoredChartStyle["dataLabels"] = new JsonObject
        {
            ["showValue"] = true,
            ["showSeries"] = false,
            ["position"] = "outside-end",
            ["numberFormat"] = "#,##0",
            ["textStyle"] = new JsonObject
            {
                ["fontSize"] = 8.5,
                ["bold"] = true,
                ["color"] = "#16324FCC",
            },
        };
        authoredChartStyle["legendTextStyle"] = new JsonObject
        {
            ["fontSize"] = 10,
            ["fontFamily"] = "Aptos",
            ["color"] = "#16324F",
        };
        authoredChartStyle["titleTextStyle"] = new JsonObject
        {
            ["fontSize"] = 14,
            ["fontFamily"] = "Aptos",
            ["fontFamilyEastAsia"] = "Noto Sans CJK SC",
            ["bold"] = false,
            ["color"] = "#16324F",
        };
        authoredChartStyle["chartAreaFill"] = new JsonObject { ["type"] = "none" };
        authoredChartStyle["plotAreaFill"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "radial",
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#FFFFFF" },
                new JsonObject { ["offset"] = 1, ["color"] = "#DCEFEA", ["opacity"] = 0.6 },
            },
        };
        authoredChart["data"]!["series"]![0]!["fill"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "linear",
            ["angle"] = 90,
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#0B8F8F" },
                new JsonObject { ["offset"] = 1, ["color"] = "#F2C14E", ["opacity"] = 0.7 },
            },
        };
        authoredChart["data"]!["series"]![0]!["pointStyles"] = new JsonArray
        {
            new JsonObject
            {
                ["index"] = 7,
                ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#F2C14E" },
                ["stroke"] = new JsonObject { ["color"] = "#16324F", ["width"] = 1.25 },
            },
        };
        authoredChart["data"]!["series"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "confidence-band",
            ["name"] = "Expected operating band",
            ["values"] = new JsonArray(22, 24, 26, 28, 31, 34, 37, 40),
            ["chartType"] = "area",
            ["axis"] = "primary",
            ["fill"] = new JsonObject
            {
                ["type"] = "solid",
                ["color"] = "#0B8F8F",
                ["opacity"] = 0.18,
            },
            ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 0.75, ["opacity"] = 0.45 },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-bubble-main",
            ["type"] = "chart",
            ["role"] = "numeric relationship evidence",
            ["frame"] = new JsonObject { ["x"] = 620, ["y"] = 300, ["width"] = 280, ["height"] = 160 },
            ["chartType"] = "bubble",
            ["title"] = "Risk vs. reach",
            ["xAxis"] = new JsonObject { ["title"] = "Reach", ["min"] = 0, ["max"] = 40, ["majorUnit"] = 10 },
            ["yAxis"] = new JsonObject { ["title"] = "Risk", ["min"] = 0, ["max"] = 20, ["majorUnit"] = 5 },
            ["style"] = new JsonObject
            {
                ["legend"] = "none",
                ["showGridlines"] = true,
                ["bubbleScale"] = 145,
                ["bubbleSizeMode"] = "width",
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray(),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "risk-reach",
                        ["name"] = "Sites",
                        ["xValues"] = new JsonArray(10, 20, 34),
                        ["values"] = new JsonArray(5, 12, 8),
                        ["bubbleSizes"] = new JsonArray(4, 9, 16),
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#0B8F8F" },
                        ["stroke"] = new JsonObject { ["color"] = "#F2C14E", ["width"] = 1 },
                    },
                },
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "transform-group-main",
            ["type"] = "group",
            ["role"] = "frame transform contract",
            ["frame"] = new JsonObject
            {
                ["x"] = 780,
                ["y"] = 360,
                ["width"] = 96,
                ["height"] = 48,
                ["rotation"] = 12,
                ["flipH"] = true,
            },
            ["elements"] = new JsonArray
            {
                new JsonObject
                {
                    ["id"] = "transform-group-child",
                    ["type"] = "shape",
                    ["frame"] = new JsonObject { ["x"] = 780, ["y"] = 360, ["width"] = 96, ["height"] = 48 },
                    ["accessibility"] = new JsonObject { ["decorative"] = true },
                    ["geometry"] = new JsonObject
                    {
                        ["kind"] = "preset",
                        ["preset"] = "round2SameRect",
                        ["adjustments"] = new JsonArray(18000, 8000),
                    },
                    ["style"] = new JsonObject
                    {
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#DCEFEA" },
                    },
                },
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-line-main",
            ["type"] = "chart",
            ["role"] = "bounded line chart behavior",
            ["frame"] = new JsonObject { ["x"] = 620, ["y"] = 120, ["width"] = 280, ["height"] = 150 },
            ["chartType"] = "line",
            ["title"] = "Measured trend",
            ["style"] = new JsonObject
            {
                ["legend"] = "none",
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 17,
                    ["fontFamily"] = "Georgia",
                    ["fontFamilyEastAsia"] = "Noto Serif CJK SC",
                    ["bold"] = true,
                    ["italic"] = true,
                    ["color"] = "#0B8F8FCC",
                },
                ["smooth"] = false,
                ["varyColors"] = true,
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Baseline", "Pilot", "Review"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "measured-trend",
                        ["name"] = "Index",
                        ["values"] = new JsonArray(JsonValue.Create(92), null, JsonValue.Create(121)),
                    },
                },
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "risk-radar-main",
            ["type"] = "chart",
            ["role"] = "bounded native radar chart",
            ["frame"] = new JsonObject { ["x"] = 610, ["y"] = 300, ["width"] = 300, ["height"] = 180 },
            ["chartType"] = "radar",
            ["title"] = "Risk profile",
            ["style"] = new JsonObject { ["legend"] = "right" },
            ["spokeAxis"] = new JsonObject
            {
                ["show"] = true,
                ["min"] = 0,
                ["max"] = 100,
                ["majorUnit"] = 20,
                ["label"] = false,
                ["axisLine"] = new JsonObject { ["color"] = "#CBD5E1", ["width"] = 0.75 },
                ["gridLine"] = new JsonObject { ["color"] = "#E2E8F0", ["width"] = 0.5, ["dash"] = "dot" },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Liquidity", "Growth", "Margin", "Resilience"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "risk-current",
                        ["name"] = "Current",
                        ["values"] = new JsonArray(72, 81, 64, 77),
                        ["stroke"] = new JsonObject { ["color"] = "#0A84FF", ["width"] = 2 },
                        ["marker"] = new JsonObject { ["symbol"] = "circle", ["size"] = 5 },
                    },
                },
            },
        });
        var invalidRadarAxisProgram = authoredProgram.DeepClone().AsObject();
        invalidRadarAxisProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "risk-radar-main")["xAxis"] = new JsonObject();
        var invalidRadarAxis = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidRadarAxisProgram.ToJsonString()));
        Assert.False(invalidRadarAxis.IsValid);
        Assert.Contains(invalidRadarAxis.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.spokeAxisConflict");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "allocation-doughnut-main",
            ["type"] = "chart",
            ["role"] = "bounded circular allocation evidence",
            ["frame"] = new JsonObject { ["x"] = 620, ["y"] = 300, ["width"] = 280, ["height"] = 160 },
            ["chartType"] = "doughnut",
            ["title"] = "Allocation mix",
            ["style"] = new JsonObject
            {
                ["legend"] = "right",
                ["startAngle"] = 135,
                ["holeSize"] = 68,
                ["dataLabels"] = new JsonObject { ["showPercent"] = true, ["position"] = "center" },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Core", "Growth", "Reserve"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "allocation-mix",
                        ["name"] = "Allocation",
                        ["values"] = new JsonArray(52, 31, 17),
                    },
                },
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "waterfall-bridge-main",
            ["type"] = "chart",
            ["role"] = "cumulative operating bridge",
            ["frame"] = new JsonObject { ["x"] = 620, ["y"] = 300, ["width"] = 280, ["height"] = 160 },
            ["chartType"] = "waterfall",
            ["title"] = "Operating bridge",
            ["yAxis"] = new JsonObject { ["title"] = "Run-rate", ["min"] = 0, ["max"] = 180, ["majorUnit"] = 30 },
            ["style"] = new JsonObject
            {
                ["legend"] = "none",
                ["gapWidth"] = 55,
                ["chartAreaFill"] = new JsonObject { ["type"] = "none" },
                ["plotAreaFill"] = new JsonObject { ["type"] = "none" },
                ["waterfall"] = new JsonObject
                {
                    ["increase"] = new JsonObject
                    {
                        ["label"] = "Increase",
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#0B8F8F" },
                        ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 0.5 },
                    },
                    ["decrease"] = new JsonObject
                    {
                        ["label"] = "Decrease",
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#C8644A" },
                        ["stroke"] = new JsonObject { ["color"] = "#C8644A", ["width"] = 0.5 },
                    },
                    ["total"] = new JsonObject
                    {
                        ["label"] = "Total",
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#16324F" },
                        ["stroke"] = new JsonObject { ["color"] = "#16324F", ["width"] = 0.5 },
                    },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Opening", "Growth", "Churn", "Cost", "Closing"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "run-rate-bridge",
                        ["name"] = "Run-rate",
                        ["values"] = new JsonArray(120, 40, -25, -10, 125),
                        ["pointRoles"] = new JsonArray("total", "delta", "delta", "delta", "total"),
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Run-rate opens at 120, rises 40, falls 25 and 10, and closes at 125.",
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-heatmap-main",
            ["type"] = "chart",
            ["role"] = "correlation intensity matrix",
            ["frame"] = new JsonObject { ["x"] = 610, ["y"] = 275, ["width"] = 300, ["height"] = 195 },
            ["chartType"] = "heatmap",
            ["title"] = "Observed relationship strength",
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 13,
                    ["fontFamily"] = "Aptos Display",
                    ["bold"] = true,
                    ["color"] = "#16324F",
                },
                ["heatmap"] = new JsonObject
                {
                    ["scale"] = "diverging",
                    ["colors"] = new JsonArray("#C8644A", "#F8F6EF", "#0B8F8F"),
                    ["domain"] = new JsonArray(-10, 10),
                    ["midpoint"] = 0,
                    ["showValues"] = true,
                    ["showColorBar"] = true,
                    ["cellGap"] = 2,
                    ["missingFill"] = "#E5E7EB",
                    ["cellStroke"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 0.5 },
                    ["axisTextStyle"] = new JsonObject { ["fontSize"] = 7.5, ["color"] = "#52606D" },
                    ["valueTextStyle"] = new JsonObject { ["fontSize"] = 8, ["bold"] = true },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Acquisition", "Retention", "Margin", "Reliability"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "segment-enterprise",
                        ["name"] = "Enterprise",
                        ["values"] = new JsonArray(8, 5, JsonValue.Create(2), null),
                    },
                    new JsonObject
                    {
                        ["id"] = "segment-midmarket",
                        ["name"] = "Mid-market",
                        ["values"] = new JsonArray(4, -2, 6, 7),
                    },
                    new JsonObject
                    {
                        ["id"] = "segment-smb",
                        ["name"] = "SMB",
                        ["values"] = new JsonArray(-6, -4, 1, 3),
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Three customer segments by four operating measures, colored from negative to positive relationship strength.",
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-candlestick-main",
            ["type"] = "chart",
            ["role"] = "daily price range",
            ["frame"] = new JsonObject { ["x"] = 40, ["y"] = 275, ["width"] = 540, ["height"] = 195 },
            ["chartType"] = "candlestick",
            ["title"] = "Daily OHLC",
            ["xAxis"] = new JsonObject
            {
                ["visible"] = true,
                ["title"] = "Session",
                ["tickLabelInterval"] = 1,
                ["textStyle"] = new JsonObject { ["fontSize"] = 7.5, ["color"] = "#52606D" },
            },
            ["yAxis"] = new JsonObject
            {
                ["visible"] = true,
                ["title"] = "USD",
                ["numberFormat"] = "0.0",
                ["min"] = 88,
                ["max"] = 120,
                ["majorUnit"] = 8,
            },
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 13,
                    ["fontFamily"] = "Aptos Display",
                    ["bold"] = true,
                    ["color"] = "#16324F",
                },
                ["candlestick"] = new JsonObject
                {
                    ["up"] = new JsonObject
                    {
                        ["fill"] = new JsonObject
                        {
                            ["type"] = "gradient",
                            ["kind"] = "linear",
                            ["angle"] = 90,
                            ["stops"] = new JsonArray
                            {
                                new JsonObject { ["offset"] = 0, ["color"] = "#DCEFEA" },
                                new JsonObject { ["offset"] = 1, ["color"] = "#0B8F8F" },
                            },
                        },
                        ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 0.6 },
                    },
                    ["down"] = new JsonObject
                    {
                        ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#C8644A" },
                        ["stroke"] = new JsonObject { ["color"] = "#8B3E2F", ["width"] = 0.6 },
                    },
                    ["wick"] = new JsonObject { ["color"] = "#16324F", ["width"] = 0.8, ["cap"] = "round" },
                    ["bodyWidthRatio"] = 0.5,
                    ["showCloseValues"] = true,
                    ["gridlineStroke"] = new JsonObject { ["color"] = "#CBD5E1", ["width"] = 0.5, ["opacity"] = 0.7 },
                    ["axisTextStyle"] = new JsonObject { ["fontSize"] = 8, ["color"] = "#52606D" },
                    ["valueTextStyle"] = new JsonObject { ["fontSize"] = 7, ["bold"] = true, ["color"] = "#16324F" },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "daily-ohlc",
                        ["name"] = "Price",
                        ["openValues"] = new JsonArray(92, 96, 94, 101, 99, 108, 111, 109),
                        ["highValues"] = new JsonArray(98, 99, 103, 104, 110, 114, 115, 117),
                        ["lowValues"] = new JsonArray(90, 91, 92, 96, 97, 104, 106, 107),
                        ["values"] = new JsonArray(96, 94, 101, 99, 108, 111, 109, 116),
                    },
                    new JsonObject
                    {
                        ["id"] = "daily-average",
                        ["name"] = "Moving average",
                        ["chartType"] = "line",
                        ["values"] = new JsonArray(94, 95, 97, 99, 102, 106, 109, 112),
                        ["stroke"] = new JsonObject { ["color"] = "#F2C14E", ["width"] = 1.4, ["cap"] = "round" },
                        ["marker"] = new JsonObject { ["symbol"] = "circle", ["size"] = 4, ["fill"] = "#F2C14E" },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Eight ordered daily open, high, low, and close observations.",
            },
        });
        var invalidCandlestickProgram = authoredProgram.DeepClone().AsObject();
        invalidCandlestickProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-candlestick-main")
            ["data"]!["series"]![0]!["lowValues"]![2] = 110;
        var invalidCandlestick = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidCandlestickProgram.ToJsonString()));
        Assert.False(invalidCandlestick.IsValid);
        Assert.Contains(invalidCandlestick.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.candlestickRange");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "adoption-numeric-overlay",
            ["type"] = "chart",
            ["role"] = "numeric adoption evidence and fitted trajectory",
            ["frame"] = new JsonObject { ["x"] = 610, ["y"] = 60, ["width"] = 300, ["height"] = 195 },
            ["chartType"] = "combo",
            ["title"] = "Adoption response",
            ["xAxis"] = new JsonObject
            {
                ["visible"] = true,
                ["title"] = "Exposure",
                ["numberFormat"] = "0.0",
                ["gridLine"] = false,
            },
            ["yAxis"] = new JsonObject
            {
                ["visible"] = true,
                ["title"] = "Adoption",
                ["numberFormat"] = "0",
                ["gridLine"] = new JsonObject { ["color"] = "#CBD5E1", ["width"] = 0.5 },
            },
            ["style"] = new JsonObject
            {
                ["legend"] = "right",
                ["bubbleScale"] = 90,
                ["bubbleSizeMode"] = "area",
                ["titleTextStyle"] = new JsonObject { ["fontSize"] = 13, ["bold"] = true, ["color"] = "#16324F" },
                ["legendTextStyle"] = new JsonObject { ["fontSize"] = 7.5, ["color"] = "#52606D" },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray(),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "observed-sites",
                        ["name"] = "Observed",
                        ["chartType"] = "bubble",
                        ["xValues"] = new JsonArray(1, 2, 3, 4),
                        ["values"] = new JsonArray(18, 31, 47, 66),
                        ["bubbleSizes"] = new JsonArray(8, 14, 20, 12),
                        ["color"] = "#0B8F8FCC",
                    },
                    new JsonObject
                    {
                        ["id"] = "fitted-response",
                        ["name"] = "Fitted",
                        ["chartType"] = "line",
                        ["xValues"] = new JsonArray(1, 2, 3, 4),
                        ["values"] = new JsonArray(20, 32, 46, 64),
                        ["stroke"] = new JsonObject { ["color"] = "#16324F", ["width"] = 1.5, ["cap"] = "round" },
                    },
                    new JsonObject
                    {
                        ["id"] = "minimum-threshold",
                        ["name"] = "Threshold",
                        ["chartType"] = "column",
                        ["xValues"] = new JsonArray(1, 2, 3, 4),
                        ["values"] = new JsonArray(10, 12, 14, 16),
                        ["color"] = "#C8644A99",
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Observed bubble evidence, fitted line, and minimum threshold columns share one numeric coordinate system.",
            },
        });
        var invalidNumericComboProgram = authoredProgram.DeepClone().AsObject();
        invalidNumericComboProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "adoption-numeric-overlay")
            ["data"]!["series"]![1]!["xValues"]![2] = 2;
        var invalidNumericCombo = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidNumericComboProgram.ToJsonString()));
        Assert.False(invalidNumericCombo.IsValid);
        Assert.Contains(invalidNumericCombo.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.numericComboXOrder");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-treemap-main",
            ["type"] = "chart",
            ["role"] = "hierarchical budget allocation",
            ["frame"] = new JsonObject { ["x"] = 40, ["y"] = 275, ["width"] = 540, ["height"] = 195 },
            ["chartType"] = "treemap",
            ["title"] = "Budget allocation",
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 13,
                    ["fontFamily"] = "Aptos Display",
                    ["bold"] = true,
                    ["color"] = "#16324F",
                },
                ["treemap"] = new JsonObject
                {
                    ["rootColors"] = new JsonArray("#0B8F8F", "#C8644A", "#F2C14E"),
                    ["border"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 0.75, ["opacity"] = 0.9 },
                    ["gap"] = 2,
                    ["headerHeight"] = 17,
                    ["depthLighten"] = 0.1,
                    ["showValues"] = true,
                    ["labelTextStyle"] = new JsonObject { ["fontSize"] = 8, ["bold"] = true },
                    ["valueTextStyle"] = new JsonObject { ["fontSize"] = 7 },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray(
                    "Engineering", "Frontend", "Backend",
                    "Sales", "Enterprise", "SMB",
                    "Design", "Research", "Product"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "budget-hierarchy",
                        ["name"] = "Budget",
                        ["levels"] = 1,
                        ["values"] = new JsonArray(1000, 400, 600, 800, 500, 300, 400, 150, 250),
                        ["parents"] = new JsonArray(
                            null, "Engineering", "Engineering",
                            null, "Sales", "Sales",
                            null, "Design", "Design"),
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Three department budgets partitioned into six direct child allocations.",
            },
        });
        var invalidTreemapProgram = authoredProgram.DeepClone().AsObject();
        invalidTreemapProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-treemap-main")
            ["data"]!["series"]![0]!["parents"]![0] = "Backend";
        var invalidTreemap = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidTreemapProgram.ToJsonString()));
        Assert.False(invalidTreemap.IsValid);
        Assert.Contains(invalidTreemap.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.treemapCycle");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-sunburst-main",
            ["type"] = "chart",
            ["role"] = "portfolio contribution hierarchy",
            ["frame"] = new JsonObject { ["x"] = 520, ["y"] = 240, ["width"] = 390, ["height"] = 270 },
            ["chartType"] = "sunburst",
            ["title"] = "Contribution by portfolio",
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 13,
                    ["fontFamily"] = "Aptos Display",
                    ["bold"] = true,
                    ["color"] = "#16324F",
                },
                ["sunburst"] = new JsonObject
                {
                    ["rootColors"] = new JsonArray("#0B8F8F", "#C8644A"),
                    ["border"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 0.6, ["opacity"] = 0.9 },
                    ["innerRadiusRatio"] = 0.18,
                    ["ringGap"] = 1.5,
                    ["segmentGapDegrees"] = 1,
                    ["startAngle"] = -90,
                    ["clockwise"] = true,
                    ["depthLighten"] = 0.1,
                    ["showValues"] = true,
                    ["labelTextStyle"] = new JsonObject { ["fontSize"] = 8, ["bold"] = true },
                    ["valueTextStyle"] = new JsonObject { ["fontSize"] = 7 },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray(
                    "Company", "Product", "Operations",
                    "Platform", "Applications", "Delivery", "Support"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "portfolio-hierarchy",
                        ["name"] = "Contribution",
                        ["levels"] = 2,
                        ["values"] = new JsonArray(100, 55, 45, 30, 25, 20, 25),
                        ["parents"] = new JsonArray(
                            null, "Company", "Company",
                            "Product", "Product", "Operations", "Operations"),
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Company contribution partitioned into product and operations, each with two child portfolios.",
            },
        });
        var invalidSunburstProgram = authoredProgram.DeepClone().AsObject();
        invalidSunburstProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-sunburst-main")
            ["data"]!["series"]![0]!["values"]![0] = 99;
        var invalidSunburst = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidSunburstProgram.ToJsonString()));
        Assert.False(invalidSunburst.IsValid);
        Assert.Contains(invalidSunburst.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.sunburstTotal");
        var invalidSunburstLevelsProgram = authoredProgram.DeepClone().AsObject();
        invalidSunburstLevelsProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-sunburst-main")
            ["data"]!["series"]![0]!["levels"] = 7;
        var invalidSunburstLevels = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidSunburstLevelsProgram.ToJsonString()));
        Assert.False(invalidSunburstLevels.IsValid);
        Assert.Contains(invalidSunburstLevels.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.sunburstLevels");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-sankey-main",
            ["type"] = "chart",
            ["role"] = "customer conversion flow",
            ["frame"] = new JsonObject { ["x"] = 40, ["y"] = 275, ["width"] = 870, ["height"] = 205 },
            ["chartType"] = "sankey",
            ["title"] = "Lead conversion flow",
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject
                {
                    ["fontSize"] = 13,
                    ["fontFamily"] = "Aptos Display",
                    ["bold"] = true,
                    ["color"] = "#16324F",
                },
                ["sankey"] = new JsonObject
                {
                    ["nodeColors"] = new JsonArray("#16324F", "#0B8F8F", "#F2C14E", "#C8644A"),
                    ["nodeStroke"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 0.5, ["opacity"] = 0.85 },
                    ["nodeWidth"] = 14,
                    ["nodeGap"] = 10,
                    ["nodeAlign"] = "right",
                    ["nodeColorMap"] = new JsonObject { ["Paid"] = "#C1121F" },
                    ["flowOpacity"] = 0.42,
                    ["flowCurvature"] = 0.72,
                    ["flowColorMode"] = "source",
                    ["showValues"] = true,
                    ["labelTextStyle"] = new JsonObject { ["fontSize"] = 8, ["bold"] = true },
                    ["valueTextStyle"] = new JsonObject { ["fontSize"] = 7, ["color"] = "#52606D" },
                },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Leads", "Qualified", "Trial", "Nurture", "Paid", "Churn", "Direct"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "conversion-flow",
                        ["name"] = "Accounts",
                        ["values"] = new JsonArray(100, 60, 40, 45, 15, 25, 15, 10),
                        ["sources"] = new JsonArray("Leads", "Qualified", "Qualified", "Trial", "Trial", "Nurture", "Nurture", "Direct"),
                        ["targets"] = new JsonArray("Qualified", "Trial", "Nurture", "Paid", "Churn", "Paid", "Churn", "Paid"),
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "One hundred leads split into trial and nurture paths while ten direct accounts join the paid outcome.",
            },
        });
        var invalidSankeyProgram = authoredProgram.DeepClone().AsObject();
        invalidSankeyProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-sankey-main")
            ["data"]!["series"]![0]!["targets"]![6] = "Qualified";
        var invalidSankey = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidSankeyProgram.ToJsonString()));
        Assert.False(invalidSankey.IsValid);
        Assert.Contains(invalidSankey.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.sankeyCycle");
        var invalidSankeyColorProgram = authoredProgram.DeepClone().AsObject();
        invalidSankeyColorProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-sankey-main")
            ["style"]!["sankey"]!["nodeColorMap"]!["Undeclared"] = "#000000";
        var invalidSankeyColor = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidSankeyColorProgram.ToJsonString()));
        Assert.False(invalidSankeyColor.IsValid);
        Assert.Contains(invalidSankeyColor.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.sankeyNodeColor");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-stream-main",
            ["type"] = "chart",
            ["role"] = "audience composition stream",
            ["frame"] = new JsonObject { ["x"] = 48, ["y"] = 272, ["width"] = 560, ["height"] = 190 },
            ["chartType"] = "area",
            ["title"] = "Audience composition changed without losing reach",
            ["xAxis"] = new JsonObject
            {
                ["visible"] = true,
                ["textStyle"] = new JsonObject { ["fontSize"] = 7, ["color"] = "#52606D" },
            },
            ["style"] = new JsonObject
            {
                ["stacking"] = "stream",
                ["legend"] = "right",
                ["titleTextStyle"] = new JsonObject { ["fontSize"] = 13, ["bold"] = true, ["color"] = "#16324F" },
                ["legendTextStyle"] = new JsonObject { ["fontSize"] = 7.5, ["color"] = "#16324F" },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Jan", "Feb", "Mar", "Apr", "May", "Jun"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "new-users",
                        ["name"] = "New",
                        ["values"] = new JsonArray(22, 28, 35, 31, 38, 44),
                        ["fill"] = new JsonObject
                        {
                            ["type"] = "gradient",
                            ["kind"] = "linear",
                            ["angle"] = 0,
                            ["stops"] = new JsonArray
                            {
                                new JsonObject { ["offset"] = 0, ["color"] = "#0B8F8F" },
                                new JsonObject { ["offset"] = 1, ["color"] = "#74C7C7", ["opacity"] = 0.82 },
                            },
                        },
                    },
                    new JsonObject
                    {
                        ["id"] = "returning-users",
                        ["name"] = "Returning",
                        ["values"] = new JsonArray(31, 34, 30, 39, 42, 46),
                        ["color"] = "#F2C14ECC",
                    },
                    new JsonObject
                    {
                        ["id"] = "enterprise-users",
                        ["name"] = "Enterprise",
                        ["values"] = new JsonArray(12, 14, 18, 22, 29, 36),
                        ["color"] = "#C8644AE6",
                        ["stroke"] = new JsonObject { ["color"] = "#8F3D2E", ["width"] = 0.5, ["opacity"] = 0.7 },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Three centered flowing bands compare audience composition from January through June.",
            },
        });
        var invalidStreamProgram = authoredProgram.DeepClone().AsObject();
        invalidStreamProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-stream-main")["chartType"] = "line";
        var invalidStream = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidStreamProgram.ToJsonString()));
        Assert.False(invalidStream.IsValid);
        Assert.Contains(invalidStream.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.streamType");
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "participants-pictograph-main",
            ["type"] = "chart",
            ["role"] = "participant pictograph bar",
            ["frame"] = new JsonObject { ["x"] = 48, ["y"] = 272, ["width"] = 560, ["height"] = 190 },
            ["chartType"] = "bar",
            ["title"] = "Verified participants by cohort",
            ["style"] = new JsonObject
            {
                ["titleTextStyle"] = new JsonObject { ["fontSize"] = 13, ["bold"] = true, ["color"] = "#16324F" },
            },
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Control", "Pilot", "Follow-up"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "verified-participants",
                        ["name"] = "Participants",
                        ["values"] = new JsonArray(30, 50, 20),
                        ["fill"] = new JsonObject
                        {
                            ["type"] = "gradient",
                            ["kind"] = "linear",
                            ["angle"] = 0,
                            ["stops"] = new JsonArray
                            {
                                new JsonObject { ["offset"] = 0, ["color"] = "#0B8F8F" },
                                new JsonObject { ["offset"] = 1, ["color"] = "#74C7C7" },
                            },
                        },
                        ["symbol"] = new JsonObject
                        {
                            ["kind"] = "icon",
                            ["iconName"] = "fas:user",
                            ["unit"] = 10,
                            ["gap"] = 2,
                            ["showValue"] = true,
                            ["unitLabel"] = "participants",
                        },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Each person symbol represents ten verified participants across three cohorts.",
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "milestones-pictograph-main",
            ["type"] = "chart",
            ["role"] = "milestone pictograph column",
            ["frame"] = new JsonObject { ["x"] = 620, ["y"] = 272, ["width"] = 290, ["height"] = 190 },
            ["chartType"] = "column",
            ["title"] = "Milestones cleared",
            ["data"] = new JsonObject
            {
                ["categories"] = new JsonArray("Q1", "Q2", "Q3"),
                ["series"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "cleared-milestones",
                        ["name"] = "Milestones",
                        ["values"] = new JsonArray(2, 4, 3),
                        ["color"] = "#F2C14E",
                        ["stroke"] = new JsonObject { ["color"] = "#9B6A00", ["width"] = 0.5 },
                        ["symbol"] = new JsonObject
                        {
                            ["kind"] = "preset",
                            ["preset"] = "star5",
                            ["unit"] = 1,
                            ["gap"] = 2,
                            ["showValue"] = true,
                            ["unitLabel"] = "gates",
                        },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "One star represents one cleared milestone in each quarter.",
            },
        });
        var invalidPictographProgram = authoredProgram.DeepClone().AsObject();
        invalidPictographProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "participants-pictograph-main")
            ["data"]!["series"]![0]!["values"]![1] = 55;
        var invalidPictograph = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidPictographProgram.ToJsonString()));
        Assert.False(invalidPictograph.IsValid);
        Assert.Contains(invalidPictograph.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.pictographUnit");
        var invalidHeatmapProgram = authoredProgram.DeepClone().AsObject();
        invalidHeatmapProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "evidence-heatmap-main")
            ["style"]!["heatmap"]!["domain"] = new JsonArray(10, 20);
        var invalidHeatmap = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidHeatmapProgram.ToJsonString()));
        Assert.False(invalidHeatmap.IsValid);
        Assert.Contains(invalidHeatmap.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.heatmapMidpoint");
        var invalidWaterfallProgram = authoredProgram.DeepClone().AsObject();
        invalidWaterfallProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "waterfall-bridge-main")
            ["data"]!["series"]![0]!["values"]![4] = 124;
        var invalidWaterfall = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidWaterfallProgram.ToJsonString()));
        Assert.False(invalidWaterfall.IsValid);
        Assert.Contains(invalidWaterfall.Diagnostics, diagnostic => diagnostic.Code == "ppj.chart.waterfallTotalMismatch");
        var invalidTextPaintProgram = authoredProgram.DeepClone().AsObject();
        invalidTextPaintProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-title")
            ["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["color"] = "#16324F";
        var invalidTextPaint = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidTextPaintProgram.ToJsonString()));
        Assert.False(invalidTextPaint.IsValid);
        Assert.Contains(invalidTextPaint.Diagnostics, diagnostic => diagnostic.Path.Contains("claim-title", StringComparison.Ordinal) ||
            diagnostic.Path.Contains("runs[0].style", StringComparison.Ordinal));
        var invalidAdjustmentProgram = authoredProgram.DeepClone().AsObject();
        invalidAdjustmentProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "transform-group-main")["elements"]![0]!["geometry"]!["adjustments"]!.AsArray()
            .Add(32000);
        var invalidAdjustments = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidAdjustmentProgram.ToJsonString()));
        Assert.False(invalidAdjustments.IsValid);
        Assert.Contains(invalidAdjustments.Diagnostics, diagnostic => diagnostic.Code == "ppj.geometry.adjustmentCount");
        var invalidArcProgram = authoredProgram.DeepClone().AsObject();
        var invalidArcCommands = invalidArcProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "claim-rule")
            ["geometry"]!["paths"]![0]!["commands"]!.AsArray();
        invalidArcCommands.Insert(0, new JsonObject
        {
            ["op"] = "arcTo",
            ["radiusX"] = 20,
            ["radiusY"] = 20,
            ["startAngle"] = 0,
            ["sweepAngle"] = 90,
        });
        var invalidArc = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidArcProgram.ToJsonString()));
        Assert.False(invalidArc.IsValid);
        Assert.Contains(invalidArc.Diagnostics, diagnostic => diagnostic.Code == "ppj.geometry.arcCurrentPoint");

        static JsonObject DiagramNode(string id, string text, string? parent = null, bool picture = false)
        {
            var node = new JsonObject { ["id"] = id, ["text"] = text };
            if (parent is not null) node["parent"] = parent;
            if (picture) node["asset"] = "evidence-mark";
            return node;
        }

        static JsonObject Diagram(
            string layout,
            double x,
            double y,
            JsonArray nodes,
            bool connected = false,
            string geometry = "roundRect")
        {
            var connections = new JsonArray();
            var nodeObjects = nodes.Select(node => node!.AsObject()).ToArray();
            var parentIds = nodeObjects.Select(node => node["parent"]?.GetValue<string>()).ToArray();
            for (var index = 0; index < nodeObjects.Length; index++) nodeObjects[index].Remove("parent");
            if (layout == "hierarchy")
            {
                for (var index = 0; index < nodeObjects.Length; index++)
                {
                    if (parentIds[index] is not { } parentId) continue;
                    var childId = nodeObjects[index]["id"]!.GetValue<string>();
                    connections.Add(new JsonObject
                    {
                        ["id"] = $"parent-{parentId}-{childId}",
                        ["from"] = parentId,
                        ["to"] = childId,
                        ["role"] = "parent",
                        ["order"] = index,
                    });
                }
            }
            else if (layout is "process" or "cycle")
            {
                var count = layout == "cycle" ? nodeObjects.Length : nodeObjects.Length - 1;
                for (var index = 0; index < count; index++)
                {
                    var fromId = nodeObjects[index]["id"]!.GetValue<string>();
                    var toId = nodeObjects[(index + 1) % nodeObjects.Length]["id"]!.GetValue<string>();
                    connections.Add(new JsonObject
                    {
                        ["id"] = $"sequence-{fromId}-{toId}",
                        ["from"] = fromId,
                        ["to"] = toId,
                        ["role"] = "sequence",
                        ["order"] = index,
                    });
                }
            }
            else if (layout == "relationship")
            {
                var rootId = nodeObjects[0]["id"]!.GetValue<string>();
                for (var index = 1; index < nodeObjects.Length; index++)
                {
                    var toId = nodeObjects[index]["id"]!.GetValue<string>();
                    connections.Add(new JsonObject
                    {
                        ["id"] = $"association-{rootId}-{toId}",
                        ["from"] = rootId,
                        ["to"] = toId,
                        ["role"] = "association",
                        ["order"] = index - 1,
                    });
                }
            }
            var diagram = new JsonObject
            {
                ["id"] = $"authored-{layout}-diagram",
                ["type"] = "smartArt",
                ["role"] = $"authored {layout} diagram",
                ["frame"] = new JsonObject { ["x"] = x, ["y"] = y, ["width"] = 210, ["height"] = 216 },
                ["mode"] = "authored",
                ["layout"] = layout,
                ["shapeStyleRef"] = "decision-band",
                ["textStyleRef"] = "body",
                ["nodeGeometry"] = new JsonObject { ["kind"] = "preset", ["preset"] = geometry },
                ["nodes"] = nodes,
                ["accessibility"] = new JsonObject
                {
                    ["decorative"] = false,
                    ["description"] = $"Editable native {layout} diagram.",
                },
            };
            if (connections.Count > 0) diagram["connections"] = connections;
            if (connected)
                diagram["connector"] = new JsonObject
                {
                    ["stroke"] = new JsonObject { ["color"] = "#0B8F8F", ["width"] = 1.5 },
                    ["endArrow"] = "triangle",
                };
            return diagram;
        }

        authoredProgram["pages"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "page-authored-diagrams",
            ["name"] = "Authored diagrams",
            ["role"] = "Exercise bounded authored diagram layouts",
            ["layout"] = "layout-evidence",
            ["background"] = new JsonObject { ["type"] = "solid", ["color"] = new JsonObject { ["token"] = "paper" } },
            ["elements"] = new JsonArray
            {
                Diagram("list", 24, 24, new JsonArray(
                    DiagramNode("list-a", "Observe"), DiagramNode("list-b", "Measure"), DiagramNode("list-c", "Decide"))),
                Diagram("process", 258, 24, new JsonArray(
                    DiagramNode("process-a", "Input"), DiagramNode("process-b", "Evaluate"), DiagramNode("process-c", "Act")), connected: true),
                Diagram("cycle", 492, 24, new JsonArray(
                    DiagramNode("cycle-a", "Plan"), DiagramNode("cycle-b", "Run"), DiagramNode("cycle-c", "Learn")), connected: true, geometry: "ellipse"),
                Diagram("hierarchy", 726, 24, new JsonArray(
                    DiagramNode("hierarchy-root", "Program"),
                    DiagramNode("hierarchy-left", "Evidence", "hierarchy-root"),
                    DiagramNode("hierarchy-right", "Delivery", "hierarchy-root")), connected: true),
                Diagram("relationship", 24, 282, new JsonArray(
                    DiagramNode("relationship-core", "Decision"), DiagramNode("relationship-a", "Cost"),
                    DiagramNode("relationship-b", "Risk"), DiagramNode("relationship-c", "Benefit")), connected: true, geometry: "ellipse"),
                Diagram("matrix", 258, 282, new JsonArray(
                    DiagramNode("matrix-a", "High / now"), DiagramNode("matrix-b", "High / later"),
                    DiagramNode("matrix-c", "Low / now"), DiagramNode("matrix-d", "Low / later")), geometry: "rect"),
                Diagram("pyramid", 492, 282, new JsonArray(
                    DiagramNode("pyramid-a", "Signal"), DiagramNode("pyramid-b", "Evidence"), DiagramNode("pyramid-c", "Decision")), geometry: "rect"),
                Diagram("picture", 726, 282, new JsonArray(
                    DiagramNode("picture-a", "Baseline", picture: true), DiagramNode("picture-b", "Pilot", picture: true),
                    DiagramNode("picture-c", "Review", picture: true), DiagramNode("picture-d", "Scale", picture: true)), geometry: "rect"),
            },
        });
        authoredProgram["pages"]![2]!["transition"] = new JsonObject
        {
            ["type"] = "split",
            ["orientation"] = "horizontal",
            ["direction"] = "in",
            ["speed"] = "fast",
            ["durationMs"] = 750,
            ["advanceOnClick"] = false,
            ["advanceAfterMs"] = 1250,
        };
        authoredProgram["sections"]![0]!["pages"] = new JsonArray("page-claim");
        authoredProgram["sections"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "section-evidence",
            ["name"] = "Evidence and appendix",
            ["pages"] = new JsonArray("page-evidence", "page-authored-diagrams"),
        });
        var mediaBytes = Convert.FromHexString("000000186674797069736F6D0000020069736F6D6D703431");
        var mediaSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(mediaBytes)).ToLowerInvariant();
        authoredProgram["assets"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-video",
            ["uri"] = "ppj-assets/evidence-video.mp4",
            ["mimeType"] = "video/mp4",
            ["sha256"] = mediaSha256,
            ["rights"] = new JsonObject
            {
                ["status"] = "internal",
                ["author"] = "OfficeKit",
                ["licenseName"] = "AGPL-3.0-or-later",
                ["creditLine"] = "Synthetic authored-media contract fixture.",
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Synthetic embedded video used to verify the PPJ media contract.",
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "evidence-video",
            ["type"] = "media",
            ["name"] = "authored evidence video",
            ["role"] = "playback evidence",
            ["frame"] = new JsonObject { ["x"] = 700, ["y"] = 300, ["width"] = 180, ["height"] = 100 },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "Embedded evidence video with an explicit poster.",
            },
            ["mediaType"] = "video",
            ["asset"] = "evidence-video",
            ["posterAsset"] = "evidence-mark",
            ["startAtMs"] = 1200,
            ["endAtMs"] = 400,
            ["loop"] = true,
            ["mute"] = true,
        });
        authoredProgram["pages"]![1]!["notes"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject
                {
                    ["id"] = "evidence-notes-paragraph",
                    ["runs"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["id"] = "evidence-notes-run-1",
                            ["text"] = "All data are ",
                            ["style"] = new JsonObject { ["fontFamily"] = "Aptos", ["size"] = 16 },
                        },
                        new JsonObject
                        {
                            ["id"] = "evidence-notes-run-2",
                            ["text"] = "illustrative",
                            ["style"] = new JsonObject
                            {
                                ["fontFamily"] = "Aptos",
                                ["size"] = 16,
                                ["bold"] = true,
                                ["color"] = "#A83232",
                            },
                        },
                        new JsonObject
                        {
                            ["id"] = "evidence-notes-run-3",
                            ["text"] = " fixture values.",
                            ["style"] = new JsonObject { ["fontFamily"] = "Aptos", ["size"] = 16 },
                        },
                    },
                },
            },
        };
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "named-icon-main",
            ["type"] = "icon",
            ["role"] = "catalog lightbulb",
            ["frame"] = new JsonObject { ["x"] = 820, ["y"] = 36, ["width"] = 80, ["height"] = 80 },
            ["iconName"] = "fas:lightbulb",
            ["style"] = new JsonObject
            {
                ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#F2C14E" },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "A lightbulb marks the central experimental insight.",
            },
        });
        authoredProgram["pages"]![0]!["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "formula-main",
            ["type"] = "text",
            ["name"] = "native formula proof",
            ["role"] = "editable quantitative model",
            ["frame"] = new JsonObject { ["x"] = 48, ["y"] = 445, ["width"] = 520, ["height"] = 52 },
            ["text"] = new JsonObject
            {
                ["paragraphs"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["id"] = "formula-paragraph",
                        ["runs"] = new JsonArray
                        {
                            new JsonObject { ["id"] = "formula-label", ["text"] = "Model: " },
                            new JsonObject
                            {
                                ["id"] = "formula-expression",
                                ["formula"] = new JsonObject
                                {
                                    ["syntax"] = "latex",
                                    ["source"] = "\\int_0^1 x^2 \\,\\mathrm{d}x = \\frac{1}{3} + \\sqrt{\\alpha+\\beta}",
                                },
                                ["style"] = new JsonObject { ["size"] = 22, ["color"] = "#14324A" },
                            },
                        },
                    },
                },
            },
            ["accessibility"] = new JsonObject
            {
                ["decorative"] = false,
                ["description"] = "A native editable integral, fraction, radical, and scripted expression.",
            },
        });
        var invalidIconProgram = authoredProgram.DeepClone().AsObject();
        invalidIconProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "named-icon-main")["iconName"] = "fas:not-an-officekit-icon";
        var invalidIcon = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidIconProgram.ToJsonString()));
        Assert.False(invalidIcon.IsValid);
        Assert.Contains(invalidIcon.Diagnostics, diagnostic => diagnostic.Code == "ppj.icon.unknown");
        var invalidFormulaProgram = authoredProgram.DeepClone().AsObject();
        invalidFormulaProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == "formula-main")
            ["text"]!["paragraphs"]![0]!["runs"]![1]!["formula"]!["source"] = "\\begin{matrix}1\\end{matrix}";
        var invalidFormula = PpjProgramValidator.Validate(Encoding.UTF8.GetBytes(invalidFormulaProgram.ToJsonString()));
        Assert.False(invalidFormula.IsValid);
        Assert.Contains(invalidFormula.Diagnostics, diagnostic => diagnostic.Code == "ppj.formula.unsupportedCommand");
        var programBytes = Encoding.UTF8.GetBytes(authoredProgram.ToJsonString());
        var assetBytes = File.ReadAllBytes(Path.Combine(fixtureDirectory, "ppj-assets", "evidence-mark.svg"));
        var request = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFrom(programBytes),
                IncludeNodeMap = true,
                Assets =
                {
                    new Asset
                    {
                        Id = "evidence-mark",
                        FileName = "evidence-mark.svg",
                        ContentType = "image/svg+xml",
                        Data = ByteString.CopyFrom(assetBytes),
                        Sha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(assetBytes)).ToLowerInvariant(),
                    },
                    new Asset
                    {
                        Id = "evidence-video",
                        FileName = "evidence-video.mp4",
                        ContentType = "video/mp4",
                        Data = ByteString.CopyFrom(mediaBytes),
                        Sha256 = mediaSha256,
                    },
                },
            },
        };

        var translucentBackgroundProgram = authoredProgram.DeepClone().AsObject();
        translucentBackgroundProgram["pages"]![0]!["background"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = "#0A84FF80",
        };
        var translucentBackgroundRequest = request.Clone();
        translucentBackgroundRequest.PresentationProgram.ProgramJson =
            ByteString.CopyFromUtf8(translucentBackgroundProgram.ToJsonString());
        var translucentBackground = Invoke(translucentBackgroundRequest);
        Assert.True(translucentBackground.Ok, Diagnostics(translucentBackground));
        using (var stream = new MemoryStream(translucentBackground.File.ToByteArray()))
        using (var package = PresentationDocument.Open(stream, false))
            Assert.Equal(50_196, OrderedSlides(package)[0].Slide!.CommonSlideData!.Background!
                .Descendants<A.Alpha>().Single().Val!.Value);

        var invalidTransitionProgram = authoredProgram.DeepClone().AsObject();
        invalidTransitionProgram["pages"]![2]!["transition"] = new JsonObject
        {
            ["type"] = "circle",
            ["direction"] = "left",
        };
        var invalidTransitionRequest = request.Clone();
        invalidTransitionRequest.PresentationProgram.ProgramJson =
            ByteString.CopyFromUtf8(invalidTransitionProgram.ToJsonString());
        var invalidTransition = Invoke(invalidTransitionRequest);
        Assert.False(invalidTransition.Ok);
        Assert.Contains(invalidTransition.Diagnostics, diagnostic => diagnostic.Code == "ppj.transition.profile");

        var first = Invoke(request);
        Assert.True(first.Ok, Diagnostics(first));
        Assert.Equal(44U, first.PresentationProgram.ExpandedElementCount);
        Assert.NotEmpty(first.PresentationProgram.NodeMapJson);
        var authoredParts = ZipPartPaths(first.File.ToByteArray());
        Assert.Contains("officeKit/program.ppj", authoredParts);
        Assert.Contains("officeKit/program-map.json", authoredParts);
        Assert.Equal(programBytes, ZipBytes(first.File.ToByteArray(), "officeKit/program.ppj"));
        using (var embeddedMap = JsonDocument.Parse(ZipBytes(first.File.ToByteArray(), "officeKit/program-map.json")))
        {
            Assert.Equal("office-kit/ppj-map/v1", embeddedMap.RootElement.GetProperty("schema").GetString());
            Assert.Equal(first.PresentationProgram.ProgramSha256, embeddedMap.RootElement.GetProperty("programSha256").GetString());
            Assert.Contains(embeddedMap.RootElement.GetProperty("nativeBindings").EnumerateArray(), binding =>
                binding.GetProperty("id").GetString() == "claim-title" && binding.GetProperty("nativeId").GetUInt32() >= 2);
            Assert.Contains(embeddedMap.RootElement.GetProperty("assets").EnumerateArray(), asset =>
                asset.GetProperty("id").GetString() == "evidence-mark");
            Assert.Contains(embeddedMap.RootElement.GetProperty("assets").EnumerateArray(), asset =>
                asset.GetProperty("id").GetString() == "evidence-video");
        }
        var validationOnlyRequest = request.Clone();
        validationOnlyRequest.PresentationProgram.ValidationOnly = true;
        var validationOnly = Invoke(validationOnlyRequest);
        Assert.True(validationOnly.Ok, Diagnostics(validationOnly));
        Assert.Empty(validationOnly.File);
        Assert.Equal(first.PresentationProgram.ProgramSha256, validationOnly.PresentationProgram.ProgramSha256);
        Assert.Empty(validationOnly.PresentationProgram.OutputSha256);
        using (var stream = new MemoryStream(first.File.ToByteArray(), writable: false))
        using (var package = PresentationDocument.Open(stream, false))
        {
            var sdkValidationErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(package).ToArray();
            var sdkMathFalsePositive = Assert.Single(sdkValidationErrors);
            Assert.Contains("/a14:m[", sdkMathFalsePositive.Path!.XPath, StringComparison.Ordinal);
            Assert.Contains("leaf element and cannot contain children", sdkMathFalsePositive.Description!, StringComparison.Ordinal);
            var nativeMaster = Assert.Single(package.PresentationPart!.SlideMasterParts);
            var nativeLayout = Assert.Single(nativeMaster.SlideLayoutParts);
            Assert.Equal("titleOnly", nativeLayout.SlideLayout!.GetAttribute("type", string.Empty).Value);
            Assert.NotNull(nativeMaster.SlideMaster!.CommonSlideData!.Background);
            Assert.Single(nativeMaster.SlideMaster.CommonSlideData.ShapeTree!.Elements<P.Shape>());
            Assert.Single(nativeLayout.SlideLayout.CommonSlideData!.ShapeTree!.Elements<P.Shape>());
            Assert.All(package.PresentationPart.SlideParts, slide => Assert.Equal(nativeLayout.Uri, slide.SlideLayoutPart!.Uri));
            var nativeTransition = package.PresentationPart.SlideParts.ElementAt(2).Slide!.GetFirstChild<P.Transition>()!;
            Assert.Equal(P.TransitionSpeedValues.Fast, nativeTransition.Speed!.Value);
            Assert.False(nativeTransition.AdvanceOnClick!.Value);
            Assert.Equal("1250", nativeTransition.AdvanceAfterTime!.Value);
            Assert.Equal("750", nativeTransition.Duration!.Value);
            var nativeSplit = nativeTransition.GetFirstChild<P.SplitTransition>()!;
            Assert.Equal("horz", nativeSplit.GetAttribute("orient", string.Empty).Value);
            Assert.Equal("in", nativeSplit.GetAttribute("dir", string.Empty).Value);
            Assert.NotNull(nativeMaster.SlideMaster.TextStyles!.TitleStyle!
                .GetFirstChild<A.Level1ParagraphProperties>()!
                .GetFirstChild<A.DefaultRunProperties>());
            var mediaSlide = package.PresentationPart.SlideParts.First();
            var nativeMedia = mediaSlide.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                .Single(picture => picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value == "authored evidence video");
            Assert.Equal("ppaction://media", nativeMedia.NonVisualPictureProperties!.NonVisualDrawingProperties!
                .GetFirstChild<A.HyperlinkOnClick>()!.Action!.Value);
            var nativeMediaProperties = nativeMedia.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties!;
            Assert.NotNull(nativeMediaProperties.GetFirstChild<A.VideoFromFile>());
            var nativeMediaExtension = nativeMediaProperties.GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtensionList>()!
                .GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtension>()!
                .GetFirstChild<P14.Media>()!;
            Assert.Equal("1200", nativeMediaExtension.GetFirstChild<P14.MediaTrim>()!.Start!.Value);
            Assert.Equal("400", nativeMediaExtension.GetFirstChild<P14.MediaTrim>()!.End!.Value);
            var videoRelationship = Assert.Single(mediaSlide.DataPartReferenceRelationships.OfType<VideoReferenceRelationship>());
            var mediaRelationship = Assert.Single(mediaSlide.DataPartReferenceRelationships.OfType<MediaReferenceRelationship>());
            Assert.Same(videoRelationship.DataPart, mediaRelationship.DataPart);
            Assert.Equal(mediaBytes, ReadDataPart(mediaRelationship.DataPart));
            var mediaTiming = mediaSlide.Slide.Timing!.OuterXml;
            Assert.Contains("<p:video", mediaTiming, StringComparison.Ordinal);
            Assert.Contains("mute=\"1\"", mediaTiming, StringComparison.Ordinal);
            Assert.Contains("repeatCount=\"indefinite\"", mediaTiming, StringComparison.Ordinal);
            Assert.Contains("<p:animEffect", mediaTiming, StringComparison.Ordinal);
            var nativeLockedClaim = mediaSlide.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>()
                .Single(shape => shape.Descendants<A.Text>().Any(text => text.Text == "Reduce incident hours "));
            Assert.True(nativeLockedClaim.NonVisualShapeProperties!.NonVisualDrawingProperties!.Hidden!.Value);
            var nativeClaimLocks = nativeLockedClaim.NonVisualShapeProperties.NonVisualShapeDrawingProperties!
                .GetFirstChild<A.ShapeLocks>()!;
            Assert.True(nativeClaimLocks.NoSelection!.Value);
            Assert.True(nativeClaimLocks.NoMove!.Value);
            Assert.True(nativeClaimLocks.NoResize!.Value);
            Assert.True(nativeClaimLocks.NoTextEdit!.Value);
            var nativeCatalogIcon = mediaSlide.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>()
                .Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "catalog lightbulb");
            var nativeIconGeometry = nativeCatalogIcon.ShapeProperties!.GetFirstChild<A.CustomGeometry>()!;
            Assert.NotEmpty(nativeIconGeometry.Descendants<A.CubicBezierCurveTo>());
            Assert.Equal("F2C14E", nativeCatalogIcon.ShapeProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("A lightbulb marks the central experimental insight.",
                nativeCatalogIcon.NonVisualShapeProperties.NonVisualDrawingProperties!.Description!.Value);
            var nativeFormula = mediaSlide.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>()
                .Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "native formula proof");
            var nativeMath = Assert.Single(nativeFormula.TextBody!.Descendants(), PptxMathCodec.IsMathElement);
            XNamespace mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var nativeMathXml = XElement.Parse(nativeMath.OuterXml);
            Assert.True(nativeMathXml.HasElements, nativeMath.OuterXml);
            Assert.True(PptxMathCodec.IsCanonical(nativeMathXml), nativeMath.OuterXml);
            Assert.Contains(nativeMathXml.Descendants(), element => element.Name == mathNamespace + "f");
            Assert.Contains(nativeMathXml.Descendants(), element => element.Name == mathNamespace + "rad");
            Assert.Contains(nativeMathXml.Descendants(), element => element.Name == mathNamespace + "sSubSup");
            Assert.Contains(nativeFormula.TextBody.Descendants<A.Text>(), text => text.Text == "Model: ");
            var nativeImageBackground = package.PresentationPart!.SlideParts.ElementAt(1).Slide!
                .CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>();
            Assert.NotNull(nativeImageBackground);
            Assert.NotNull(nativeImageBackground!.GetFirstChild<A.Stretch>());
            Assert.Null(nativeImageBackground.GetFirstChild<A.Tile>());
            var nativeImageFillShape = package.PresentationPart.SlideParts.ElementAt(1).Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "decision-flow-start");
            var nativeCompoundImageFill = nativeImageFillShape.ShapeProperties!.GetFirstChild<A.BlipFill>()!;
            Assert.Equal(33_000, nativeCompoundImageFill.GetFirstChild<A.Blip>()!
                .GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
            Assert.Equal(40_000, nativeImageFillShape.ShapeProperties.GetFirstChild<A.Outline>()!
                .GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!
                .GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(30_000, nativeImageFillShape.ShapeProperties.GetFirstChild<A.EffectList>()!
                .GetFirstChild<A.OuterShadow>()!.GetFirstChild<A.RgbColorModelHex>()!
                .GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(50_000, nativeImageFillShape.TextBody!.Descendants<A.RunProperties>().Single()
                .GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!
                .GetFirstChild<A.Alpha>()!.Val!.Value);
            var nativeTiledPicture = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                .Single(picture => picture.BlipFill!.GetFirstChild<A.Tile>() is not null);
            Assert.NotNull(nativeTiledPicture.BlipFill!.GetFirstChild<A.Tile>());
            Assert.Null(nativeTiledPicture.BlipFill.GetFirstChild<A.Stretch>());
            var nativeCustomMaskPicture = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.Picture>()
                .Single(picture => picture.ShapeProperties!.GetFirstChild<A.CustomGeometry>() is not null);
            Assert.NotNull(nativeCustomMaskPicture.ShapeProperties!.GetFirstChild<A.CustomGeometry>());
            Assert.Null(nativeCustomMaskPicture.ShapeProperties.GetFirstChild<A.PresetGeometry>());
            var nativeClaim = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single(shape => shape.TextBody?.Descendants<A.Text>().Any(text => text.Text == "Reduce incident hours ") == true);
            var nativeClaimParagraph = Assert.Single(nativeClaim.TextBody!.Elements<A.Paragraph>());
            var nativeClaimRunProperties = nativeClaimParagraph.Elements<A.Run>().First().RunProperties!;
            Assert.Equal(18 * 60_000, nativeClaimRunProperties.GetFirstChild<A.GradientFill>()!
                .GetFirstChild<A.LinearGradientFill>()!.Angle!.Value);
            Assert.Equal(80_000, nativeClaimRunProperties.GetFirstChild<A.GradientFill>()!
                .GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>().Last()
                .GetFirstChild<A.RgbColorModelHex>()!.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(3 * 12_700, nativeClaimRunProperties.GetFirstChild<A.EffectList>()!
                .GetFirstChild<A.OuterShadow>()!.BlurRadius!.Value);
            var nativeDefaultText = nativeClaimParagraph.ParagraphProperties!.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.Equal(A.PathShadeValues.Circle, nativeDefaultText.GetFirstChild<A.GradientFill>()!
                .GetFirstChild<A.PathGradientFill>()!.Path!.Value);
            Assert.Equal(2 * 12_700, nativeDefaultText.GetFirstChild<A.EffectList>()!
                .GetFirstChild<A.OuterShadow>()!.BlurRadius!.Value);
            var comboChartPath = package.PresentationPart.SlideParts.ElementAt(1).ChartParts.Single()
                .Uri.OriginalString.TrimStart('/');
            var comboChartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(first.File.ToByteArray(), comboChartPath)));
            XNamespace richChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace richDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var richTitleRuns = comboChartXml.Root!.Element(richChartNamespace + "chart")!
                .Element(richChartNamespace + "title")!
                .Descendants(richDrawingNamespace + "r").ToArray();
            Assert.Equal(2, richTitleRuns.Length);
            Assert.Equal(["Measured profile: ", "−38% incidents"],
                richTitleRuns.Select(run => run.Element(richDrawingNamespace + "t")!.Value));
            Assert.Equal("16324F", richTitleRuns[0].Element(richDrawingNamespace + "rPr")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!
                .Attribute("val")!.Value);
            Assert.Equal("C1121F", richTitleRuns[1].Element(richDrawingNamespace + "rPr")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!
                .Attribute("val")!.Value);
            var primaryCategoryAxis = comboChartXml.Descendants(richChartNamespace + "catAx")
                .Single(axis => axis.Element(richChartNamespace + "axPos")?.Attribute("val")?.Value == "b");
            Assert.Equal("maxMin", primaryCategoryAxis.Element(richChartNamespace + "scaling")!
                .Element(richChartNamespace + "orientation")!.Attribute("val")!.Value);
            Assert.Equal("16324F", primaryCategoryAxis.Element(richChartNamespace + "spPr")!
                .Element(richDrawingNamespace + "ln")!.Element(richDrawingNamespace + "solidFill")!
                .Element(richDrawingNamespace + "srgbClr")!.Attribute("val")!.Value);
            Assert.Equal("arrow", primaryCategoryAxis.Element(richChartNamespace + "spPr")!
                .Element(richDrawingNamespace + "ln")!.Element(richDrawingNamespace + "headEnd")!
                .Attribute("type")!.Value);
            Assert.Equal("triangle", primaryCategoryAxis.Element(richChartNamespace + "spPr")!
                .Element(richDrawingNamespace + "ln")!.Element(richDrawingNamespace + "tailEnd")!
                .Attribute("type")!.Value);
            Assert.NotNull(primaryCategoryAxis.Element(richChartNamespace + "majorGridlines")!
                .Element(richChartNamespace + "spPr")!.Element(richDrawingNamespace + "ln")!
                .Element(richDrawingNamespace + "noFill"));
            var primaryValueAxis = comboChartXml.Descendants(richChartNamespace + "valAx")
                .Single(axis => axis.Element(richChartNamespace + "axPos")?.Attribute("val")?.Value == "l");
            Assert.Equal("DCEFEA", primaryValueAxis.Element(richChartNamespace + "majorGridlines")!
                .Element(richChartNamespace + "spPr")!.Element(richDrawingNamespace + "ln")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!
                .Attribute("val")!.Value);
            Assert.Contains(comboChartXml.Descendants(richChartNamespace + "dLbls"), labels =>
                labels.Element(richChartNamespace + "numFmt")?.Attribute("formatCode")?.Value == "#,##0");
            var analyticalLabelContainer = comboChartXml.Descendants(richChartNamespace + "ser")
                .Single(series => series.Element(richChartNamespace + "tx")?.Element(richChartNamespace + "v")?.Value == authoredChartSeries["name"]!.GetValue<string>())
                .Element(richChartNamespace + "dLbls")!;
            Assert.Equal("0.0", analyticalLabelContainer.Element(richChartNamespace + "numFmt")!.Attribute("formatCode")!.Value);
            Assert.Equal(["2", "7"], analyticalLabelContainer.Elements(richChartNamespace + "dLbl")
                .Select(label => label.Element(richChartNamespace + "idx")!.Attribute("val")!.Value));
            Assert.Equal("0.0x", analyticalLabelContainer.Elements(richChartNamespace + "dLbl").Last()
                .Element(richChartNamespace + "numFmt")!.Attribute("formatCode")!.Value);
            var highlightedIncidentPoint = comboChartXml.Descendants(richChartNamespace + "ser")
                .Single(series => series.Element(richChartNamespace + "tx")?.Element(richChartNamespace + "v")?.Value == "P1 incident hours")
                .Elements(richChartNamespace + "dPt").Single();
            Assert.Equal("7", highlightedIncidentPoint.Element(richChartNamespace + "idx")!.Attribute("val")!.Value);
            Assert.Equal("F2C14E", highlightedIncidentPoint.Element(richChartNamespace + "spPr")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!.Attribute("val")!.Value);
            var lineChartPath = package.PresentationPart.SlideParts.First().ChartParts
                .Single(part => part.ChartSpace!.Descendants<C.LineChart>().Any())
                .Uri.OriginalString.TrimStart('/');
            var lineChartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(first.File.ToByteArray(), lineChartPath)));
            XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var lineValues = lineChartXml.Descendants(chartNamespace + "lineChart").Single()
                .Descendants(chartNamespace + "val").Single()
                .Element(chartNamespace + "numLit")!;
            Assert.Equal("3", lineValues.Element(chartNamespace + "ptCount")!.Attribute("val")!.Value);
            Assert.Equal(["0", "2"], lineValues.Elements(chartNamespace + "pt").Select(point => point.Attribute("idx")!.Value));
            var waterfallChartPart = package.PresentationPart.SlideParts.First().ChartParts
                .Single(part => part.ChartSpace!.Descendants<C.BarChart>().Any(chart =>
                    chart.Descendants<C.SeriesText>().Any(text => text.InnerText == "__offset__")));
            var waterfallChart = waterfallChartPart.ChartSpace!.Descendants<C.BarChart>().Single();
            Assert.Equal(C.BarGroupingValues.Stacked, waterfallChart.BarGrouping!.Val!.Value);
            var waterfallSeries = waterfallChart.Elements<C.BarChartSeries>().ToArray();
            Assert.Equal(4, waterfallSeries.Length);
            Assert.Equal(["__offset__", "Increase", "Decrease", "Total"],
                waterfallSeries.Select(series => series.SeriesText!.InnerText));
            Assert.NotNull(waterfallSeries[0].ChartShapeProperties!.GetFirstChild<A.NoFill>());
            static Dictionary<uint, double> LiteralValues(C.BarChartSeries series) =>
                series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberLiteral>()!.Elements<C.NumericPoint>()
                    .ToDictionary(point => point.Index!.Value, point => double.Parse(point.NumericValue!.Text, CultureInfo.InvariantCulture));
            Assert.Equal(new Dictionary<uint, double> { [0] = 0, [1] = 120, [2] = 135, [3] = 125, [4] = 0 }, LiteralValues(waterfallSeries[0]));
            Assert.Equal(new Dictionary<uint, double> { [1] = 40 }, LiteralValues(waterfallSeries[1]));
            Assert.Equal(new Dictionary<uint, double> { [2] = 25, [3] = 10 }, LiteralValues(waterfallSeries[2]));
            Assert.Equal(new Dictionary<uint, double> { [0] = 120, [4] = 125 }, LiteralValues(waterfallSeries[3]));
            var circularChartPath = package.PresentationPart.SlideParts.First().ChartParts
                .Single(part => part.ChartSpace!.Descendants<C.DoughnutChart>().Any())
                .Uri.OriginalString.TrimStart('/');
            var circularChartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(first.File.ToByteArray(), circularChartPath)));
            Assert.Equal("135", circularChartXml.Descendants(chartNamespace + "firstSliceAng").Single().Attribute("val")!.Value);
            Assert.Equal("68", circularChartXml.Descendants(chartNamespace + "holeSize").Single().Attribute("val")!.Value);
            var bubbleChartPath = package.PresentationPart.SlideParts.First().ChartParts
                .Single(part => part.ChartSpace!.Descendants<C.BubbleChart>().Any())
                .Uri.OriginalString.TrimStart('/');
            var bubbleChartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(first.File.ToByteArray(), bubbleChartPath)));
            Assert.Equal("145", bubbleChartXml.Descendants(chartNamespace + "bubbleScale").Single().Attribute("val")!.Value);
            Assert.Equal("w", bubbleChartXml.Descendants(chartNamespace + "sizeRepresents").Single().Attribute("val")!.Value);
            var radarChartPath = package.PresentationPart.SlideParts.First().ChartParts
                .Single(part => part.ChartSpace!.Descendants<C.RadarChart>().Any())
                .Uri.OriginalString.TrimStart('/');
            var radarChartXml = XDocument.Parse(Encoding.UTF8.GetString(ZipBytes(first.File.ToByteArray(), radarChartPath)));
            var radarCategoryAxis = radarChartXml.Descendants(chartNamespace + "catAx").Single();
            var radarValueAxis = radarChartXml.Descendants(chartNamespace + "valAx").Single();
            Assert.Equal("0", radarCategoryAxis.Element(chartNamespace + "delete")!.Attribute("val")!.Value);
            Assert.Equal("CBD5E1", radarCategoryAxis.Element(chartNamespace + "majorGridlines")!
                .Element(chartNamespace + "spPr")!.Element(richDrawingNamespace + "ln")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!
                .Attribute("val")!.Value);
            Assert.Equal("0", radarValueAxis.Element(chartNamespace + "scaling")!
                .Element(chartNamespace + "min")!.Attribute("val")!.Value);
            Assert.Equal("100", radarValueAxis.Element(chartNamespace + "scaling")!
                .Element(chartNamespace + "max")!.Attribute("val")!.Value);
            Assert.Equal("20", radarValueAxis.Element(chartNamespace + "majorUnit")!.Attribute("val")!.Value);
            Assert.Equal("none", radarValueAxis.Element(chartNamespace + "tickLblPos")!.Attribute("val")!.Value);
            Assert.Equal("E2E8F0", radarValueAxis.Element(chartNamespace + "majorGridlines")!
                .Element(chartNamespace + "spPr")!.Element(richDrawingNamespace + "ln")!
                .Element(richDrawingNamespace + "solidFill")!.Element(richDrawingNamespace + "srgbClr")!
                .Attribute("val")!.Value);
            Assert.Equal(6, package.PresentationPart.SlideParts.SelectMany(slide => slide.ChartParts).Count());
            var nativeHeatmap = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "correlation intensity matrix");
            Assert.Equal(12, nativeHeatmap.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("heatmap cell ", StringComparison.Ordinal) == true));
            Assert.Contains(nativeHeatmap.Descendants<A.Text>(), text => text.Text == "Observed relationship strength");
            Assert.Contains(nativeHeatmap.Descendants<A.Text>(), text => text.Text == "Enterprise");
            Assert.Contains(nativeHeatmap.Descendants<A.Text>(), text => text.Text == "-6");
            Assert.NotNull(nativeHeatmap.Descendants<A.GradientFill>().Single());
            var nativeCandlestick = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "daily price range");
            Assert.Equal(8, nativeCandlestick.Elements<P.ConnectionShape>().Count(connector =>
                connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("candlestick wick ", StringComparison.Ordinal) == true));
            Assert.Equal(8, nativeCandlestick.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.Contains(" body ", StringComparison.Ordinal) == true));
            Assert.Contains(nativeCandlestick.Descendants<A.Text>(), text => text.Text == "Daily OHLC");
            Assert.Contains(nativeCandlestick.Descendants<A.Text>(), text => text.Text == "116.0");
            Assert.NotEmpty(nativeCandlestick.Descendants<A.GradientFill>());
            Assert.Equal(7, nativeCandlestick.Elements<P.ConnectionShape>().Count(connector =>
                connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("candlestick line Moving average ", StringComparison.Ordinal) == true));
            var candlestickChildren = nativeCandlestick.ChildElements.ToArray();
            var lastBodyIndex = Array.FindLastIndex(candlestickChildren, child =>
                child.Descendants<P.NonVisualDrawingProperties>().Any(properties => properties.Name?.Value?.Contains(" body ", StringComparison.Ordinal) == true));
            var firstAverageIndex = Array.FindIndex(candlestickChildren, child =>
                child.Descendants<P.NonVisualDrawingProperties>().Any(properties => properties.Name?.Value?.StartsWith("candlestick line Moving average ", StringComparison.Ordinal) == true));
            Assert.True(firstAverageIndex > lastBodyIndex);
            var nativeNumericCombo = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "numeric adoption evidence and fitted trajectory");
            Assert.Equal(4, nativeNumericCombo.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("numeric bubble Observed ", StringComparison.Ordinal) == true));
            Assert.Equal(4, nativeNumericCombo.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("numeric column Threshold ", StringComparison.Ordinal) == true));
            Assert.Equal(3, nativeNumericCombo.Elements<P.ConnectionShape>().Count(connector =>
                connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("numeric line Fitted ", StringComparison.Ordinal) == true));
            var numericChildren = nativeNumericCombo.ChildElements.ToArray();
            var lastColumnIndex = Array.FindLastIndex(numericChildren, child =>
                child.Descendants<P.NonVisualDrawingProperties>().Any(properties => properties.Name?.Value?.StartsWith("numeric column Threshold ", StringComparison.Ordinal) == true));
            var firstLineIndex = Array.FindIndex(numericChildren, child =>
                child.Descendants<P.NonVisualDrawingProperties>().Any(properties => properties.Name?.Value?.StartsWith("numeric line Fitted ", StringComparison.Ordinal) == true));
            var firstBubbleIndex = Array.FindIndex(numericChildren, child =>
                child.Descendants<P.NonVisualDrawingProperties>().Any(properties => properties.Name?.Value?.StartsWith("numeric bubble Observed ", StringComparison.Ordinal) == true));
            Assert.True(lastColumnIndex < firstLineIndex && firstLineIndex < firstBubbleIndex);
            var nativeTreemap = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "hierarchical budget allocation");
            Assert.Equal(3, nativeTreemap.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("treemap node ", StringComparison.Ordinal) == true));
            Assert.Contains(nativeTreemap.Descendants<A.Text>(), text => text.Text == "Engineering");
            Assert.DoesNotContain(nativeTreemap.Descendants<A.Text>(), text => text.Text == "Frontend");
            Assert.Contains(nativeTreemap.Descendants<A.Text>(), text => text.Text == "1000");
            var nativeSunburst = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "portfolio contribution hierarchy");
            Assert.Equal(3, nativeSunburst.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("sunburst sector ", StringComparison.Ordinal) == true));
            Assert.All(nativeSunburst.Elements<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("sunburst sector ", StringComparison.Ordinal) == true),
                shape => Assert.NotNull(shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()));
            Assert.NotEmpty(nativeSunburst.Descendants<A.CubicBezierCurveTo>());
            Assert.Contains(nativeSunburst.Descendants<A.Text>(), text => text.Text == "Product");
            Assert.DoesNotContain(nativeSunburst.Descendants<A.Text>(), text => text.Text == "Platform");
            var nativeSankey = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "customer conversion flow");
            Assert.Equal(8, nativeSankey.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("sankey flow ", StringComparison.Ordinal) == true));
            Assert.Equal(7, nativeSankey.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("sankey node ", StringComparison.Ordinal) == true));
            Assert.All(nativeSankey.Elements<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("sankey flow ", StringComparison.Ordinal) == true),
                shape => Assert.NotNull(shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()));
            Assert.Contains(nativeSankey.Descendants<A.Text>(), text => text.Text == "Qualified");
            Assert.Contains(nativeSankey.Descendants<A.Text>(), text => text.Text == "100");
            var paidNode = nativeSankey.Elements<P.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "sankey node Paid");
            Assert.Equal("C1121F", paidNode.ShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            var directNode = nativeSankey.Elements<P.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "sankey node Direct");
            var trialNode = nativeSankey.Elements<P.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "sankey node Trial");
            Assert.Equal(
                trialNode.ShapeProperties!.Transform2D!.Offset!.X!.Value,
                directNode.ShapeProperties!.Transform2D!.Offset!.X!.Value);
            var nativeStream = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "audience composition stream");
            Assert.Equal(3, nativeStream.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("stream band ", StringComparison.Ordinal) == true));
            Assert.All(nativeStream.Elements<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("stream band ", StringComparison.Ordinal) == true),
                shape =>
                {
                    Assert.NotNull(shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>());
                    Assert.NotEmpty(shape.Descendants<A.CubicBezierCurveTo>());
                });
            Assert.Contains(nativeStream.Descendants<A.Text>(), text => text.Text == "Audience composition changed without losing reach");
            Assert.Contains(nativeStream.Descendants<A.Text>(), text => text.Text == "Enterprise");
            Assert.NotEmpty(nativeStream.Descendants<A.GradientFill>());
            var nativeParticipantPictograph = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "participant pictograph bar");
            Assert.Equal(10, nativeParticipantPictograph.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("pictographic symbol ", StringComparison.Ordinal) == true));
            Assert.All(nativeParticipantPictograph.Elements<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("pictographic symbol ", StringComparison.Ordinal) == true),
                shape => Assert.NotNull(shape.ShapeProperties!.GetFirstChild<A.CustomGeometry>()));
            Assert.Contains(nativeParticipantPictograph.Descendants<A.Text>(), text => text.Text == "1 symbol = 10 participants");
            Assert.Contains(nativeParticipantPictograph.Descendants<A.Text>(), text => text.Text == "50 participants");
            Assert.NotEmpty(nativeParticipantPictograph.Descendants<A.GradientFill>());
            var nativeMilestonePictograph = package.PresentationPart.SlideParts.First().Slide!
                .CommonSlideData!.ShapeTree!.Elements<P.GroupShape>()
                .Single(group => group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "milestone pictograph column");
            Assert.Equal(9, nativeMilestonePictograph.Elements<P.Shape>().Count(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("pictographic symbol ", StringComparison.Ordinal) == true));
            Assert.All(nativeMilestonePictograph.Elements<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("pictographic symbol ", StringComparison.Ordinal) == true),
                shape => Assert.Equal(A.ShapeTypeValues.Star5, shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.Value));
            Assert.Contains(nativeMilestonePictograph.Descendants<A.Text>(), text => text.Text == "4 gates");
            var diagramSlide = package.PresentationPart.SlideParts.ElementAt(2);
            var diagramFrames = diagramSlide.Slide!.CommonSlideData!.ShapeTree!.Elements<P.GraphicFrame>()
                .Where(frame => frame.Descendants<Dgm.RelationshipIds>().Any()).ToArray();
            Assert.Equal(8, diagramFrames.Length);
            Assert.Equal(8, diagramSlide.DiagramDataParts.Count());
            Assert.Equal(8, diagramSlide.DiagramLayoutDefinitionParts.Count());
            Assert.Equal(8, diagramSlide.DiagramStyleParts.Count());
            Assert.Equal(8, diagramSlide.DiagramColorsParts.Count());
            Assert.Equal(8, diagramSlide.Parts.Count(pair => pair.OpenXmlPart is DiagramPersistLayoutPart));
            var nativeProcessDiagram = diagramFrames.Single(frame =>
                frame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name!.Value == "authored process diagram");
            var processRelationships = nativeProcessDiagram.Descendants<Dgm.RelationshipIds>().Single();
            var processDataPart = Assert.IsType<DiagramDataPart>(diagramSlide.GetPartById(processRelationships.DataPart!));
            Assert.Contains(processDataPart.DataModelRoot!.Descendants<A.Text>(), text => text.Text == "Evaluate");
            Assert.Equal(2, processDataPart.DataModelRoot.Descendants<Dgm.Connection>()
                .Count(connection => connection.Type?.Value == Dgm.ConnectionValues.UnknownRelationship));
            var pictureDrawing = diagramSlide.Parts.Select(pair => pair.OpenXmlPart).OfType<DiagramPersistLayoutPart>()
                .Single(part => part.Drawing!.Descendants<A.Text>().Any(text => text.Text == "Review"));
            Assert.Equal(4, pictureDrawing.Drawing!.Descendants<OD.Shape>().Count());
            Assert.Single(pictureDrawing.ImageParts);
            var nativeTable = package.PresentationPart!.SlideParts.ElementAt(1).Slide!.Descendants<A.Table>().Single();
            Assert.True(nativeTable.TableProperties!.FirstRow!.Value);
            var firstCell = nativeTable.Descendants<A.TableCell>().First();
            Assert.Equal(A.TextAnchoringTypeValues.Center, firstCell.TextBody!.BodyProperties!.Anchor!.Value);
            Assert.Equal("DCEFEA", firstCell.TableCellProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(80_000, firstCell.TableCellProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.GetFirstChild<A.Alpha>()!.Val!.Value);
            var firstHeaderRun = firstCell.TextBody!.Descendants<A.RunProperties>().Single();
            Assert.True(firstHeaderRun.Bold!.Value);
            Assert.Equal("C1121F", firstHeaderRun.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            var explicitImageCell = nativeTable.Descendants<A.TableCell>().ElementAt(1);
            var explicitImageFill = explicitImageCell.TableCellProperties!.GetFirstChild<A.BlipFill>()!;
            Assert.NotNull(explicitImageFill.GetFirstChild<A.Tile>());
            Assert.Equal(55_000, explicitImageFill.GetFirstChild<A.Blip>()!
                .GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
            var secondHeaderCell = nativeTable.Descendants<A.TableCell>().ElementAt(2);
            var secondHeaderFill = secondHeaderCell.TableCellProperties!.GetFirstChild<A.SolidFill>()!;
            Assert.Equal("B7DEE8", secondHeaderFill.RgbColorModelHex!.Val!.Value);
            Assert.Equal(64_000, secondHeaderFill.RgbColorModelHex.GetFirstChild<A.Alpha>()!.Val!.Value);
            var secondHeaderRun = secondHeaderCell.TextBody!.Descendants<A.RunProperties>().Single();
            Assert.True(secondHeaderRun.Bold!.Value);
            Assert.Equal(1_100, secondHeaderRun.FontSize!.Value);
            Assert.Equal("0B8F8F", secondHeaderRun.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(88_000, secondHeaderRun.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.True(secondHeaderRun.Italic!.Value);
            var inheritedBodyRightBorder = secondHeaderCell.TableCellProperties.GetFirstChild<A.RightBorderLineProperties>()!;
            Assert.Equal("9525", inheritedBodyRightBorder.GetAttribute("w", string.Empty).Value);
            Assert.Equal("C1121F", inheritedBodyRightBorder.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            var inheritedImageCell = nativeTable.Descendants<A.TableCell>().ElementAt(4);
            var inheritedImageFill = inheritedImageCell.TableCellProperties!.GetFirstChild<A.BlipFill>()!;
            Assert.NotNull(inheritedImageFill.GetFirstChild<A.Stretch>());
            Assert.NotNull(inheritedImageFill.GetFirstChild<A.SourceRectangle>());
            Assert.Equal(22_000, inheritedImageFill.GetFirstChild<A.Blip>()!
                .GetFirstChild<A.AlphaModulationFixed>()!.Amount!.Value);
            Assert.True(inheritedImageCell.TextBody!.Descendants<A.RunProperties>().Single().Italic!.Value);
            var inheritedBaseLeftBorder = inheritedImageCell.TableCellProperties.GetFirstChild<A.LeftBorderLineProperties>()!;
            Assert.Equal("6350", inheritedBaseLeftBorder.GetAttribute("w", string.Empty).Value);
            var inheritedLastRowBorder = inheritedImageCell.TableCellProperties.GetFirstChild<A.BottomBorderLineProperties>()!;
            Assert.Equal("25400", inheritedLastRowBorder.GetAttribute("w", string.Empty).Value);
            var bottomBorder = firstCell.TableCellProperties.GetFirstChild<A.BottomBorderLineProperties>()!;
            Assert.Equal("19050", bottomBorder.GetAttribute("w", string.Empty).Value);
            Assert.Equal("0B8F8F", bottomBorder.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(65_000, bottomBorder.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.GetFirstChild<A.Alpha>()!.Val!.Value);
        }

        var imported = Import(first.File.ToByteArray());
        Assert.True(imported.Ok, Diagnostics(imported));
        Assert.Equal(3, imported.Artifact.Presentation.Slides.Count);
        var importedTransition = imported.Artifact.Presentation.Slides[2].Transition;
        Assert.Equal("split", importedTransition.Effect);
        Assert.Equal("horizontal", importedTransition.Orientation);
        Assert.Equal("in", importedTransition.Direction);
        Assert.Equal("fast", importedTransition.Speed);
        Assert.True(importedTransition.HasDurationMs);
        Assert.Equal(750U, importedTransition.DurationMs);
        Assert.False(importedTransition.AdvanceOnClick);
        Assert.True(importedTransition.HasAdvanceAfterMs);
        Assert.Equal(1250U, importedTransition.AdvanceAfterMs);
        var importedDiagrams = imported.Artifact.Presentation.Slides[2].Elements
            .Where(element => element.ContentCase == PresentationElement.ContentOneofCase.Diagram)
            .ToArray();
        Assert.Equal(8, importedDiagrams.Length);
        var importedProcessDiagram = importedDiagrams.Single(element => element.Name == "authored process diagram").Diagram;
        Assert.Equal(3, importedProcessDiagram.Nodes.Count);
        Assert.Equal(2, importedProcessDiagram.Connections.Count);
        Assert.Equal("Evaluate", PptxTextCodec.Flatten(importedProcessDiagram.Nodes[1].TextBody));
        var importedPictureDiagram = importedDiagrams.Single(element => element.Name == "authored picture diagram").Diagram;
        Assert.Equal(4, importedPictureDiagram.Nodes.Count);
        Assert.All(importedPictureDiagram.Nodes, node => Assert.False(string.IsNullOrWhiteSpace(node.AssetId)));
        var importedBackground = imported.Artifact.Presentation.Slides[0].Background.GradientFill;
        Assert.Equal(PresentationGradientFill.Types.Kind.Linear, importedBackground.Kind);
        Assert.Equal(3, importedBackground.Stops.Count);
        Assert.True(importedBackground.Stops[1].HasOpacityThousandthPercent);
        Assert.Equal(86_000U, importedBackground.Stops[1].OpacityThousandthPercent);
        var importedImageBackground = imported.Artifact.Presentation.Slides[1].Background.ImagePaint;
        Assert.Equal(PresentationImagePaint.Types.Mode.Stretch, importedImageBackground.Mode);
        Assert.Equal(72_000U, importedImageBackground.OpacityThousandthPercent);
        Assert.True(importedImageBackground.Crop.TopThousandthPercent > 0);
        Assert.True(importedImageBackground.Crop.BottomThousandthPercent > 0);
        var importedImage = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Image && element.Name == "evidence identity").Image;
        Assert.True(importedImage.Tiled);
        Assert.Equal(92_000U, importedImage.OpacityThousandthPercent);
        Assert.Equal(4_000, importedImage.Crop.LeftThousandthPercent);
        Assert.Equal(3_000, importedImage.Crop.TopThousandthPercent);
        Assert.Equal(2_000, importedImage.Crop.RightThousandthPercent);
        Assert.Equal(1_000, importedImage.Crop.BottomThousandthPercent);
        Assert.Equal("roundRect", importedImage.MaskPreset);
        Assert.Equal([24000], importedImage.MaskPresetAdjustments);
        Assert.Equal("0B8F8F", importedImage.Border.ColorRgb);
        Assert.Equal(24_000U, importedImage.Shadow.OpacityThousandthPercent);
        Assert.Equal("Rising evidence line", importedImage.AltText);
        var importedCustomMaskImage = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Image && element.Name == "irregular editorial crop").Image;
        Assert.Empty(importedCustomMaskImage.MaskPreset);
        Assert.Empty(importedCustomMaskImage.MaskPresetAdjustments);
        Assert.Equal(5, Assert.Single(importedCustomMaskImage.CustomMaskPaths).Commands.Count);
        var importedBubble = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Chart &&
            element.Chart.Type == SpreadsheetChartType.Bubble).Chart;
        Assert.Equal(SpreadsheetChartType.Bubble, importedBubble.Type);
        Assert.Equal([10D, 20D, 34D], Assert.Single(importedBubble.Series).XValues);
        Assert.Equal([4D, 9D, 16D], importedBubble.Series[0].BubbleSizes);
        Assert.Equal(40, importedBubble.XAxis.Maximum);
        var importedLine = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Chart &&
            element.Chart.Type == SpreadsheetChartType.Line).Chart;
        Assert.Equal(17, importedLine.TitleTextStyle.FontSizePoints);
        Assert.Equal("Georgia", importedLine.TitleTextStyle.FontFamily);
        Assert.Equal("Noto Serif CJK SC", importedLine.TitleTextStyle.FontFamilyEastAsia);
        Assert.True(importedLine.TitleTextStyle.Bold);
        Assert.True(importedLine.TitleTextStyle.Italic);
        Assert.Equal("0B8F8F", importedLine.TitleTextStyle.ColorRgb);
        Assert.Equal(80_000U, importedLine.TitleTextStyle.OpacityThousandthPercent);
        Assert.True(importedLine.LineOptions.HasSmooth);
        Assert.False(importedLine.LineOptions.Smooth);
        Assert.True(importedLine.LineOptions.VaryColors);
        Assert.Equal([92D, 0D, 121D], Assert.Single(importedLine.Series).Values);
        Assert.Equal([1U], importedLine.Series[0].MissingValueIndexes);
        var importedCircular = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Chart &&
            element.Chart.Type == SpreadsheetChartType.Doughnut).Chart;
        Assert.True(importedCircular.HasFirstSliceAngle);
        Assert.Equal(135U, importedCircular.FirstSliceAngle);
        Assert.True(importedCircular.HasDoughnutHoleSize);
        Assert.Equal(68U, importedCircular.DoughnutHoleSize);
        var importedClaimElement = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Shape &&
            element.Shape.Text.Contains("Reduce incident hours", StringComparison.Ordinal));
        Assert.True(importedClaimElement.HasHidden && importedClaimElement.Hidden);
        Assert.True(importedClaimElement.HasLocked && importedClaimElement.Locked);
        var importedClaim = importedClaimElement.Shape;
        Assert.Equal(2, Assert.Single(importedClaim.TextBody.Paragraphs).Runs.Count);
        Assert.Equal("Reduce incident hours ", importedClaim.TextBody.Paragraphs[0].Runs[0].Text);
        Assert.Equal("without weakening workload", importedClaim.TextBody.Paragraphs[0].Runs[1].Text);
        Assert.Equal("Main decision claim", importedClaim.Accessibility.Description);
        var importedFormula = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Shape && element.Name == "native formula proof").Shape;
        var importedFormulaRuns = Assert.Single(importedFormula.TextBody.Paragraphs).Runs;
        Assert.Equal(PresentationTextRun.ContentOneofCase.Text, importedFormulaRuns[0].ContentCase);
        Assert.Equal(PresentationTextRun.ContentOneofCase.Formula, importedFormulaRuns[1].ContentCase);
        Assert.Empty(importedFormulaRuns[1].Formula.SourceLatex);
        Assert.Contains("(1)/(3)", importedFormulaRuns[1].Formula.PlainText, StringComparison.Ordinal);
        Assert.Equal("•", importedClaim.TextBody.Paragraphs[0].BulletCharacter);
        Assert.Equal("0B8F8F", importedClaim.TextBody.Paragraphs[0].BulletColorRgb);
        Assert.Equal(50_000U, importedClaim.TextBody.Paragraphs[0].BulletColorOpacityThousandthPercent);
        Assert.Equal(PresentationGradientFill.Types.Kind.Linear, importedClaim.TextBody.Paragraphs[0].Runs[0].GradientFill.Kind);
        Assert.Equal(2, importedClaim.TextBody.Paragraphs[0].Runs[0].GradientFill.Stops.Count);
        Assert.Equal(80_000U, importedClaim.TextBody.Paragraphs[0].Runs[0].GradientFill.Stops[1].OpacityThousandthPercent);
        Assert.Equal(3 * 12_700, importedClaim.TextBody.Paragraphs[0].Runs[0].Shadow.BlurRadiusEmu);
        Assert.Equal(PresentationGradientFill.Types.Kind.Radial, importedClaim.TextBody.Paragraphs[0].DefaultRunProperties.GradientFill.Kind);
        Assert.Equal(2 * 12_700, importedClaim.TextBody.Paragraphs[0].DefaultRunProperties.Shadow.BlurRadiusEmu);
        var importedCompoundShape = Assert.Single(imported.Artifact.Presentation.Slides[1].Elements, element =>
            element.Name == "decision-flow-start" && element.ContentCase == PresentationElement.ContentOneofCase.Shape).Shape;
        Assert.Equal(33_000U, importedCompoundShape.ImageFill.OpacityThousandthPercent);
        Assert.Equal(40_000U, importedCompoundShape.LineOpacityThousandthPercent);
        Assert.Equal(30_000U, importedCompoundShape.Shadow.OpacityThousandthPercent);
        Assert.Equal(50_000U, Assert.Single(Assert.Single(importedCompoundShape.TextBody.Paragraphs).Runs).ColorOpacityThousandthPercent);
        var importedCustomShape = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.Name == "claim-rule" && element.ContentCase == PresentationElement.ContentOneofCase.Shape).Shape;
        Assert.Equal("custom", importedCustomShape.Geometry);
        var importedCustomPath = Assert.Single(importedCustomShape.CustomPaths);
        Assert.Equal(6, importedCustomPath.Commands.Count);
        var importedArc = importedCustomPath.Commands[1].ArcTo;
        Assert.Equal(50_000, importedArc.WidthRadius);
        Assert.Equal(50_000, importedArc.HeightRadius);
        Assert.Equal(180 * 60_000, importedArc.StartAngle);
        Assert.Equal(180 * 60_000, importedArc.SweepAngle);
        Assert.Equal(PresentationGradientFill.Types.Kind.Radial, importedCustomShape.GradientFill.Kind);
        Assert.Equal(2, importedCustomShape.GradientFill.Stops.Count);
        Assert.Equal("0B8F8F", importedCustomShape.GradientFill.Stops[1].ColorRgb);
        Assert.Equal(50_000U, importedCustomShape.GradientFill.Stops[0].OpacityThousandthPercent);
        Assert.Equal(17_500U, importedCustomShape.GradientFill.Stops[1].OpacityThousandthPercent);
        Assert.True(importedCustomShape.HasLineOpacityThousandthPercent);
        Assert.Equal(21_000U, importedCustomShape.LineOpacityThousandthPercent);
        var importedChart = Assert.Single(imported.Artifact.Presentation.Slides[1].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Chart).Chart;
        Assert.Equal("Measured profile: −38% incidents", importedChart.Title);
        Assert.NotNull(importedChart.TitleBody);
        var importedTitleRuns = Assert.Single(importedChart.TitleBody.Paragraphs).Runs;
        Assert.Equal(2, importedTitleRuns.Count);
        Assert.Equal("Measured profile: ", importedTitleRuns[0].Text);
        Assert.Equal("16324F", importedTitleRuns[0].ColorRgb);
        Assert.Equal("−38% incidents", importedTitleRuns[1].Text);
        Assert.True(importedTitleRuns[1].Bold);
        Assert.Equal("C1121F", importedTitleRuns[1].ColorRgb);
        Assert.Equal("Noto Serif CJK SC", importedTitleRuns[1].FontFamilyEastAsia);
        Assert.Equal(360_000, importedChart.FrameTransform.RotationAngle60000);
        Assert.True(importedChart.FrameTransform.FlipHorizontal);
        Assert.Equal("bottom", importedChart.LegendPosition);
        Assert.Equal(10, importedChart.LegendTextStyle.FontSizePoints);
        Assert.Equal("Aptos", importedChart.LegendTextStyle.FontFamily);
        Assert.Equal("none", importedChart.Grouping);
        Assert.True(importedChart.ChartAreaFill.NoFill);
        Assert.Equal(PresentationGradientFill.Types.Kind.Radial, importedChart.PlotAreaFill.GradientFill.Kind);
        Assert.Equal(2, importedChart.PlotAreaFill.GradientFill.Stops.Count);
        Assert.Equal(PresentationGradientFill.Types.Kind.Linear, importedChart.ComboSeries[0].Series.SeriesFill.GradientFill.Kind);
        Assert.Equal(90 * 60_000, importedChart.ComboSeries[0].Series.SeriesFill.GradientFill.Angle60000);
        Assert.Equal(70_000U, importedChart.ComboSeries[0].Series.SeriesFill.GradientFill.Stops[1].OpacityThousandthPercent);
        Assert.Equal(90U, importedChart.GapWidth);
        Assert.Equal("round", importedChart.ComboSeries[1].Series.Line.Cap);
        Assert.Equal("round", importedChart.ComboSeries[1].Series.Line.Join);
        Assert.Equal(SpreadsheetChartType.Area, importedChart.ComboSeries[2].Type);
        Assert.Equal(PresentationChartAxisGroup.Primary, importedChart.ComboSeries[2].AxisGroup);
        Assert.Equal("Expected operating band", importedChart.ComboSeries[2].Series.Name);
        Assert.Equal(40, importedChart.ComboSeries[2].Series.Values[^1]);
        Assert.Equal("Half-year", importedChart.XAxis.Title);
        Assert.Equal(1U, importedChart.XAxis.TickLabelInterval);
        Assert.Equal(9, importedChart.XAxis.TextStyle.FontSizePoints);
        Assert.Equal("Aptos", importedChart.XAxis.TextStyle.FontFamily);
        Assert.Equal("Noto Sans CJK SC", importedChart.XAxis.TextStyle.FontFamilyEastAsia);
        Assert.True(importedChart.XAxis.TextStyle.Bold);
        Assert.True(importedChart.XAxis.TextStyle.HasItalic);
        Assert.False(importedChart.XAxis.TextStyle.Italic);
        Assert.Equal("16324F", importedChart.XAxis.TextStyle.ColorRgb);
        Assert.Equal(80_000U, importedChart.XAxis.TextStyle.OpacityThousandthPercent);
        Assert.Equal(11, importedChart.XAxis.TitleTextStyle.FontSizePoints);
        Assert.Equal("Georgia", importedChart.XAxis.TitleTextStyle.FontFamily);
        Assert.Equal("0B8F8F", importedChart.XAxis.TitleTextStyle.ColorRgb);
        Assert.Equal("Incident hours", importedChart.YAxis.Title);
        Assert.Equal(0, importedChart.YAxis.Minimum);
        Assert.Equal(80, importedChart.YAxis.Maximum);
        Assert.Equal("Workload index", importedChart.SecondaryYAxis.Title);
        Assert.Equal(130, importedChart.SecondaryYAxis.Maximum);
        var importedAnalyticalSeries = importedChart.ComboSeries[1].Series;
        Assert.Equal(SpreadsheetChartMarkerSymbol.Circle, importedAnalyticalSeries.Marker.Symbol);
        Assert.Equal(8U, importedAnalyticalSeries.Marker.Size);
        Assert.Equal("F2C14E", importedAnalyticalSeries.Marker.Fill.Rgb);
        Assert.Equal(50_196U, importedAnalyticalSeries.Marker.FillOpacityThousandthPercent);
        Assert.Equal(SpreadsheetChartTrendlineType.Linear, Assert.Single(importedAnalyticalSeries.Trendlines).Type);
        Assert.Equal(SpreadsheetChartErrorBarValueType.StandardError, importedAnalyticalSeries.ErrorBars.ValueType);
        Assert.True(importedAnalyticalSeries.ErrorBars.NoEndCap);
        Assert.True(importedAnalyticalSeries.DataLabels.Defaults.ShowValue);
        Assert.Equal("0.0", importedAnalyticalSeries.DataLabels.Defaults.NumberFormatCode);
        Assert.Equal(2, importedAnalyticalSeries.DataLabels.Points.Count);
        Assert.Equal(7U, importedAnalyticalSeries.DataLabels.Points[1].Index);
        Assert.Equal("0.0x", importedAnalyticalSeries.DataLabels.Points[1].Override.NumberFormatCode);
        Assert.True(importedChart.DataLabels.HasShowSeriesName);
        Assert.False(importedChart.DataLabels.ShowSeriesName);
        Assert.Equal(8.5, importedChart.DataLabels.TextStyle.FontSizePoints);
        Assert.True(importedChart.DataLabels.TextStyle.Bold);
        Assert.Equal("16324F", importedChart.DataLabels.TextStyle.ColorRgb);
        Assert.Equal(80_000U, importedChart.DataLabels.TextStyle.OpacityThousandthPercent);
        Assert.Equal("Incident hours decline from 69 to 43 while protected workload index rises from 100 to 127.", importedChart.Accessibility.Description);
        var importedTable = Assert.Single(imported.Artifact.Presentation.Slides[1].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Table).Table;
        Assert.Equal(-240_000, importedTable.FrameTransform.RotationAngle60000);
        Assert.True(importedTable.FrameTransform.FlipVertical);
        Assert.Equal("Pilot method table", importedTable.Accessibility.Description);
        Assert.Equal(3, importedTable.Rows.Count);
        var importedGroup = Assert.Single(imported.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Group && element.Name == "frame transform contract").Group;
        Assert.Equal(720_000, importedGroup.FrameTransform.RotationAngle60000);
        Assert.True(importedGroup.FrameTransform.FlipHorizontal);
        Assert.Equal("round2SameRect", Assert.Single(importedGroup.Children).Shape.Geometry);
        Assert.Equal([18000, 8000], importedGroup.Children[0].Shape.PresetAdjustments);
        var importedConnector = Assert.Single(imported.Artifact.Presentation.Slides[1].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Connector).Connector;
        Assert.True(importedConnector.HasLineOpacityThousandthPercent);
        Assert.Equal(58_000U, importedConnector.LineOpacityThousandthPercent);

        var recovered = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = first.File,
            PresentationProgram = new PresentationProgramRequest
            {
                IncludeNodeMap = true,
                SourceUri = "deck.assets/source/source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(recovered.Ok, Diagnostics(recovered));
        Assert.True(recovered.PresentationProgram.RestoredEmbeddedProgram);
        Assert.False(recovered.PresentationProgram.SourceBound);
        Assert.Empty(recovered.PresentationProgram.SourceSha256);
        Assert.Equal(first.PresentationProgram.ProgramJson, recovered.PresentationProgram.ProgramJson);
        Assert.Equal(programBytes, recovered.PresentationProgram.OriginalProgramJson.ToByteArray());
        Assert.Equal(first.PresentationProgram.NodeMapJson, recovered.PresentationProgram.NodeMapJson);
        Assert.Equal(2, recovered.PresentationProgram.Assets.Count);
        var recoveredAsset = Assert.Single(recovered.PresentationProgram.Assets, asset => asset.Id == "evidence-mark");
        Assert.Equal("evidence-mark", recoveredAsset.Id);
        Assert.Equal("ppj-assets/evidence-mark.svg", recoveredAsset.FileName);
        Assert.Equal(assetBytes, recoveredAsset.Data.ToByteArray());
        var recoveredMedia = Assert.Single(recovered.PresentationProgram.Assets, asset => asset.Id == "evidence-video");
        Assert.Equal("ppj-assets/evidence-video.mp4", recoveredMedia.FileName);
        Assert.Equal("video/mp4", recoveredMedia.ContentType);
        Assert.Equal(mediaBytes, recoveredMedia.Data.ToByteArray());

        var nativeDrift = ReplaceZipText(first.File.ToByteArray(), "ppt/slides/slide1.xml", xml =>
            xml.Replace("</p:sld>", "<!-- external native drift --></p:sld>", StringComparison.Ordinal));
        var driftRecovery = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeDrift),
            PresentationProgram = new PresentationProgramRequest { IncludeNodeMap = true },
        });
        Assert.True(driftRecovery.Ok, Diagnostics(driftRecovery));
        Assert.True(driftRecovery.PresentationProgram.RestoredEmbeddedProgram);
        Assert.Equal(first.PresentationProgram.ProgramJson, driftRecovery.PresentationProgram.ProgramJson);
        Assert.Equal(programBytes, driftRecovery.PresentationProgram.OriginalProgramJson.ToByteArray());
        Assert.Contains(driftRecovery.Diagnostics, diagnostic => diagnostic.Code == "ppj.embedded.nativeDriftIgnored");

        var corruptSnapshot = ReplaceZipText(first.File.ToByteArray(), "officeKit/program-map.json", _ => "{}");
        var corruptFallback = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(corruptSnapshot),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/corrupt-snapshot.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(corruptFallback.Ok, Diagnostics(corruptFallback));
        Assert.False(corruptFallback.PresentationProgram.RestoredEmbeddedProgram);
        Assert.True(corruptFallback.PresentationProgram.SourceBound);

        var thirdPartySource = RemoveEmbeddedPpj(first.File.ToByteArray());
        var thirdPartySha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(thirdPartySource)).ToLowerInvariant();
        var projected = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                IncludeNodeMap = true,
                SourceUri = "deck.assets/source/source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(projected.Ok, Diagnostics(projected));
        Assert.True(projected.PresentationProgram.SourceBound);
        Assert.False(projected.PresentationProgram.RestoredEmbeddedProgram);
        Assert.Equal(thirdPartySha256, projected.PresentationProgram.SourceSha256);
        Assert.NotEmpty(projected.PresentationProgram.Assets);
        using (var projectedJson = JsonDocument.Parse(projected.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var projectedRoot = projectedJson.RootElement;
            Assert.Equal("office-kit/ppj/v1", projectedRoot.GetProperty("schema").GetString());
            Assert.Equal(projected.PresentationProgram.SourceSha256, projectedRoot.GetProperty("source").GetProperty("sha256").GetString());
            var projectedCanvas = projectedRoot.GetProperty("design").GetProperty("canvas");
            Assert.Contains(projectedCanvas.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setCanvas" &&
                capability.GetProperty("fields").EnumerateArray().Select(field => field.GetString()).SequenceEqual(["canvas.width", "canvas.height"]));
            Assert.Equal(3, projectedRoot.GetProperty("pages").GetArrayLength());
            Assert.All(projectedRoot.GetProperty("pages").EnumerateArray(), page =>
                Assert.StartsWith("layout-", page.GetProperty("layout").GetString(), StringComparison.Ordinal));
            Assert.Equal("linear", projectedRoot.GetProperty("pages")[0].GetProperty("background").GetProperty("kind").GetString());
            var projectedTransition = projectedRoot.GetProperty("pages")[2].GetProperty("transition");
            Assert.Equal("split", projectedTransition.GetProperty("type").GetString());
            Assert.Equal("horizontal", projectedTransition.GetProperty("orientation").GetString());
            Assert.Equal("in", projectedTransition.GetProperty("direction").GetString());
            Assert.Equal("fast", projectedTransition.GetProperty("speed").GetString());
            Assert.Equal(750, projectedTransition.GetProperty("durationMs").GetInt32());
            Assert.False(projectedTransition.GetProperty("advanceOnClick").GetBoolean());
            Assert.Equal(1250, projectedTransition.GetProperty("advanceAfterMs").GetInt32());
            Assert.Contains(projectedRoot.GetProperty("pages")[2].GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setTransition");
            var projectedNotes = projectedRoot.GetProperty("pages")[1].GetProperty("notes")
                .GetProperty("paragraphs")[0].GetProperty("runs");
            Assert.Equal(3, projectedNotes.GetArrayLength());
            Assert.Equal("illustrative", projectedNotes[1].GetProperty("text").GetString());
            Assert.True(projectedNotes[1].GetProperty("style").GetProperty("bold").GetBoolean());
            Assert.Equal("#A83232", projectedNotes[1].GetProperty("style").GetProperty("color").GetString());
            Assert.Contains(projectedRoot.GetProperty("pages")[1].GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setNotes" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "notes"));
            Assert.False(projectedRoot.GetProperty("pages")[2].TryGetProperty("notes", out _));
            Assert.Contains(projectedRoot.GetProperty("pages")[2].GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setNotes");
            var projectedComment = Assert.Single(projectedRoot.GetProperty("comments").EnumerateArray());
            Assert.Contains(projectedComment.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "replaceText" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "text"));
            var projectedSections = projectedRoot.GetProperty("sections");
            Assert.Equal(2, projectedSections.GetArrayLength());
            Assert.All(projectedSections.EnumerateArray(), section =>
            {
                var capabilities = section.GetProperty("nativeRef").GetProperty("capabilities");
                Assert.Contains(capabilities.EnumerateArray(), capability => capability.GetProperty("operation").GetString() == "setName");
                Assert.Contains(capabilities.EnumerateArray(), capability => capability.GetProperty("operation").GetString() == "setPages");
            });
            var projectedShow = Assert.Single(projectedRoot.GetProperty("customShows").EnumerateArray());
            Assert.Contains(projectedShow.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setPages" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "pages"));
            Assert.Equal("image", projectedRoot.GetProperty("pages")[1].GetProperty("background").GetProperty("type").GetString());
            Assert.Equal("stretch", projectedRoot.GetProperty("pages")[1].GetProperty("background").GetProperty("fit").GetString());
            Assert.Contains(projectedRoot.GetProperty("pages")[1].GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setBackground");
            Assert.Contains(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("type").GetString() == "image" &&
                item.GetProperty("fit").GetString() == "tile" &&
                item.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray().Any(capability =>
                    capability.GetProperty("operation").GetString() == "setImageFit"));
            Assert.Contains(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "irregular editorial crop" &&
                item.GetProperty("mask").GetProperty("kind").GetString() == "custom" &&
                item.GetProperty("mask").GetProperty("paths")[0].GetProperty("commands").GetArrayLength() == 5);
            Assert.Contains(projectedRoot.GetProperty("pages")[1].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "decision-flow-start" &&
                item.GetProperty("style").GetProperty("fill").GetProperty("type").GetString() == "image" &&
                item.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray().Any(capability =>
                    capability.GetProperty("operation").GetString() == "setFill"));
            Assert.Contains(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "claim-rule" &&
                item.GetProperty("geometry").GetProperty("kind").GetString() == "custom" &&
                item.GetProperty("style").GetProperty("fill").GetProperty("kind").GetString() == "radial" &&
                item.GetProperty("style").GetProperty("stroke").GetProperty("opacity").GetDouble() == 0.21);
            var projectedCustomShape = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("name").GetString() == "claim-rule");
            var projectedArc = projectedCustomShape.GetProperty("geometry").GetProperty("paths")[0]
                .GetProperty("commands").EnumerateArray()
                .Single(command => command.GetProperty("op").GetString() == "arcTo");
            Assert.Equal(50, projectedArc.GetProperty("radiusX").GetDouble());
            Assert.Equal(50, projectedArc.GetProperty("radiusY").GetDouble());
            Assert.Equal(180, projectedArc.GetProperty("startAngle").GetDouble());
            Assert.Equal(180, projectedArc.GetProperty("sweepAngle").GetDouble());
            Assert.Contains(
                projectedRoot.GetProperty("pages").EnumerateArray()
                    .SelectMany(page => page.GetProperty("elements").EnumerateArray()),
                item => item.GetProperty("name").GetString() == "decision-flow-link" &&
                    item.GetProperty("stroke").GetProperty("opacity").GetDouble() == 0.58);
            Assert.Contains(
                projectedRoot.GetProperty("pages").EnumerateArray()
                    .SelectMany(page => page.GetProperty("elements").EnumerateArray()),
                item => item.GetProperty("type").GetString() == "table" &&
                    item.GetProperty("nativeRef").GetProperty("leaves").EnumerateArray().Any(leaf =>
                        leaf.GetProperty("kind").GetString() == "tableCellText"));
            var projectedClaim = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "text" &&
                    item.GetProperty("text").ValueKind == JsonValueKind.Object &&
                    item.GetProperty("text").GetProperty("paragraphs")[0].GetProperty("runs")[0]
                        .GetProperty("text").GetString() == "Reduce incident hours ");
            Assert.True(projectedClaim.GetProperty("hidden").GetBoolean());
            Assert.True(projectedClaim.GetProperty("locked").GetBoolean());
            Assert.Contains(projectedClaim.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setHidden");
            Assert.Contains(projectedClaim.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setLocked");
            Assert.Equal("linear", projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("runs")[0].GetProperty("style").GetProperty("gradient").GetProperty("kind").GetString());
            Assert.Equal(0.8, projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("runs")[0].GetProperty("style").GetProperty("gradient").GetProperty("stops")[1]
                .GetProperty("opacity").GetDouble(), 5);
            Assert.Equal(3, projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("runs")[0].GetProperty("style").GetProperty("shadow").GetProperty("blur").GetDouble());
            Assert.Equal("radial", projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("style").GetProperty("defaultText").GetProperty("gradient").GetProperty("kind").GetString());
            Assert.Equal("#FFF2CC", projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("runs")[0].GetProperty("style").GetProperty("highlight").GetString());
            Assert.Equal("zh-CN", projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("runs")[0].GetProperty("style").GetProperty("language").GetString());
            Assert.Equal("#0B8F8F80", projectedClaim.GetProperty("text").GetProperty("paragraphs")[0]
                .GetProperty("style").GetProperty("bullet").GetProperty("color").GetString());
            var projectedChart = projectedRoot.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "chart");
            var projectedTitleRuns = projectedChart.GetProperty("title").GetProperty("paragraphs")[0]
                .GetProperty("runs");
            Assert.Equal(2, projectedTitleRuns.GetArrayLength());
            Assert.Equal("Measured profile: ", projectedTitleRuns[0].GetProperty("text").GetString());
            Assert.Equal("−38% incidents", projectedTitleRuns[1].GetProperty("text").GetString());
            Assert.Contains(projectedChart.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartTitle");
            Assert.Equal(6, projectedChart.GetProperty("frame").GetProperty("rotation").GetDouble());
            Assert.True(projectedChart.GetProperty("frame").GetProperty("flipH").GetBoolean());
            Assert.Contains(projectedChart.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setFrame" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "frame.rotation"));
            Assert.Equal("Half-year", projectedChart.GetProperty("xAxis").GetProperty("title").GetString());
            Assert.Equal(80, projectedChart.GetProperty("yAxis").GetProperty("max").GetDouble());
            var projectedSeries = projectedChart.GetProperty("data").GetProperty("series")[1];
            Assert.Equal(8, projectedSeries.GetProperty("marker").GetProperty("size").GetInt32());
            Assert.Equal("#F2C14E80", projectedSeries.GetProperty("marker").GetProperty("fill").GetString());
            Assert.Equal("solid", projectedSeries.GetProperty("fill").GetProperty("type").GetString());
            Assert.Equal("#F2C14E", projectedSeries.GetProperty("fill").GetProperty("color").GetString());
            Assert.Equal(128 / 255d, projectedSeries.GetProperty("fill").GetProperty("opacity").GetDouble(), 5);
            Assert.Equal("linear", projectedSeries.GetProperty("trendlines")[0].GetProperty("type").GetString());
            Assert.Equal("standard-error", projectedSeries.GetProperty("errorBars").GetProperty("valueType").GetString());
            Assert.Equal("0.0", projectedSeries.GetProperty("dataLabels").GetProperty("numberFormat").GetString());
            Assert.False(projectedSeries.GetProperty("dataLabels").GetProperty("points")[0].GetProperty("showValue").GetBoolean());
            Assert.Equal(7, projectedSeries.GetProperty("dataLabels").GetProperty("points")[1].GetProperty("index").GetInt32());
            Assert.Equal("top", projectedSeries.GetProperty("dataLabels").GetProperty("points")[1].GetProperty("position").GetString());
            Assert.True(projectedChart.GetProperty("style").GetProperty("dataLabels").GetProperty("showValue").GetBoolean());
            Assert.Equal("#,##0", projectedChart.GetProperty("style").GetProperty("dataLabels").GetProperty("numberFormat").GetString());
            Assert.True(projectedChart.GetProperty("xAxis").GetProperty("reverse").GetBoolean());
            Assert.Equal("#16324F", projectedChart.GetProperty("xAxis").GetProperty("axisLine").GetProperty("color").GetString());
            Assert.Equal("open", projectedChart.GetProperty("xAxis").GetProperty("axisLineArrow").GetProperty("start").GetString());
            Assert.Equal("triangle", projectedChart.GetProperty("xAxis").GetProperty("axisLineArrow").GetProperty("end").GetString());
            Assert.False(projectedChart.GetProperty("xAxis").GetProperty("gridLine").GetBoolean());
            Assert.Equal("#DCEFEA", projectedChart.GetProperty("yAxis").GetProperty("gridLine").GetProperty("color").GetString());
            Assert.Contains(projectedChart.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartLabels");
            Assert.Contains(projectedChart.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartAxis");
            Assert.Equal(10, projectedChart.GetProperty("style").GetProperty("legendTextStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal(8.5, projectedChart.GetProperty("style").GetProperty("dataLabels").GetProperty("textStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal(11, projectedChart.GetProperty("xAxis").GetProperty("titleTextStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal("none", projectedChart.GetProperty("style").GetProperty("chartAreaFill").GetProperty("type").GetString());
            Assert.Equal("radial", projectedChart.GetProperty("style").GetProperty("plotAreaFill").GetProperty("kind").GetString());
            Assert.Equal("linear", projectedChart.GetProperty("data").GetProperty("series")[0].GetProperty("fill").GetProperty("kind").GetString());
            var projectedPointStyle = projectedChart.GetProperty("data").GetProperty("series")[0].GetProperty("pointStyles")[0];
            Assert.Equal(7, projectedPointStyle.GetProperty("index").GetInt32());
            Assert.Equal("#F2C14E", projectedPointStyle.GetProperty("fill").GetProperty("color").GetString());
            Assert.Equal(1.25, projectedPointStyle.GetProperty("stroke").GetProperty("width").GetDouble());
            Assert.Equal("area", projectedChart.GetProperty("data").GetProperty("series")[2].GetProperty("chartType").GetString());
            Assert.Equal("primary", projectedChart.GetProperty("data").GetProperty("series")[2].GetProperty("axis").GetString());
            Assert.Equal(40, projectedChart.GetProperty("data").GetProperty("series")[2].GetProperty("values")[7].GetDouble());
            var projectedBubble = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "chart" &&
                    item.GetProperty("chartType").GetString() == "bubble");
            Assert.Equal(20, projectedBubble.GetProperty("data").GetProperty("series")[0].GetProperty("xValues")[1].GetDouble());
            Assert.Equal(16, projectedBubble.GetProperty("data").GetProperty("series")[0].GetProperty("bubbleSizes")[2].GetDouble());
            Assert.Equal(145, projectedBubble.GetProperty("style").GetProperty("bubbleScale").GetInt32());
            Assert.Equal("width", projectedBubble.GetProperty("style").GetProperty("bubbleSizeMode").GetString());
            Assert.Contains(projectedBubble.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartPlot");
            var projectedLine = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "chart" &&
                    item.GetProperty("chartType").GetString() == "line");
            var projectedLineTitleStyle = projectedLine.GetProperty("style").GetProperty("titleTextStyle");
            Assert.Equal(17, projectedLineTitleStyle.GetProperty("fontSize").GetDouble());
            Assert.Equal("Georgia", projectedLineTitleStyle.GetProperty("fontFamily").GetString());
            Assert.Equal("Noto Serif CJK SC", projectedLineTitleStyle.GetProperty("fontFamilyEastAsia").GetString());
            Assert.True(projectedLineTitleStyle.GetProperty("bold").GetBoolean());
            Assert.True(projectedLineTitleStyle.GetProperty("italic").GetBoolean());
            Assert.Equal("#0B8F8FCC", projectedLineTitleStyle.GetProperty("color").GetString());
            Assert.Contains(projectedLine.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartTextStyle");
            Assert.False(projectedLine.GetProperty("style").GetProperty("smooth").GetBoolean());
            Assert.True(projectedLine.GetProperty("style").GetProperty("varyColors").GetBoolean());
            Assert.Equal(JsonValueKind.Null, projectedLine.GetProperty("data").GetProperty("series")[0]
                .GetProperty("values")[1].ValueKind);
            var projectedRadar = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "chart" &&
                    item.GetProperty("chartType").GetString() == "radar");
            Assert.Equal("Risk profile", projectedRadar.GetProperty("title").GetString());
            Assert.Equal(77, projectedRadar.GetProperty("data").GetProperty("series")[0]
                .GetProperty("values")[3].GetDouble());
            Assert.Equal("circle", projectedRadar.GetProperty("data").GetProperty("series")[0]
                .GetProperty("marker").GetProperty("symbol").GetString());
            var projectedSpokeAxis = projectedRadar.GetProperty("spokeAxis");
            Assert.True(projectedSpokeAxis.GetProperty("show").GetBoolean());
            Assert.Equal(0, projectedSpokeAxis.GetProperty("min").GetDouble());
            Assert.Equal(100, projectedSpokeAxis.GetProperty("max").GetDouble());
            Assert.Equal(20, projectedSpokeAxis.GetProperty("majorUnit").GetDouble());
            Assert.False(projectedSpokeAxis.GetProperty("label").GetBoolean());
            Assert.Equal("#CBD5E1", projectedSpokeAxis.GetProperty("axisLine").GetProperty("color").GetString());
            Assert.Equal("#E2E8F0", projectedSpokeAxis.GetProperty("gridLine").GetProperty("color").GetString());
            Assert.Contains(projectedRadar.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartAxis");
            var projectedCircular = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "chart" &&
                    item.GetProperty("chartType").GetString() == "doughnut");
            Assert.Equal(135, projectedCircular.GetProperty("style").GetProperty("startAngle").GetInt32());
            Assert.Equal(68, projectedCircular.GetProperty("style").GetProperty("holeSize").GetInt32());
            Assert.Contains(projectedCircular.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setChartPlot" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "chart.plot"));
            var projectedHeatmap = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "correlation intensity matrix");
            Assert.Contains(projectedHeatmap.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "heatmap cell 1,1");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "correlation intensity matrix" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedCandlestick = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "daily price range");
            Assert.Contains(projectedCandlestick.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "candlestick wick 1");
            Assert.Contains(projectedCandlestick.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "candlestick line Moving average 1");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "daily price range" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedNumericCombo = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "numeric adoption evidence and fitted trajectory");
            Assert.Contains(projectedNumericCombo.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "numeric bubble Observed 1");
            Assert.Contains(projectedNumericCombo.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "numeric line Fitted 1");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "numeric adoption evidence and fitted trajectory" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedTreemap = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "hierarchical budget allocation");
            Assert.DoesNotContain(projectedTreemap.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "treemap node Frontend");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "hierarchical budget allocation" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedSunburst = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "portfolio contribution hierarchy");
            Assert.DoesNotContain(projectedSunburst.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "sunburst sector Platform");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "portfolio contribution hierarchy" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedSankey = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "customer conversion flow");
            Assert.Contains(projectedSankey.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "sankey flow Qualified to Trial" &&
                    item.GetProperty("geometry").GetProperty("kind").GetString() == "custom");
            Assert.Contains(projectedSankey.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "sankey node Paid");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "customer conversion flow" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedStream = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "audience composition stream");
            Assert.Equal(3, projectedStream.GetProperty("elements").EnumerateArray().Count(item =>
                item.GetProperty("name").GetString()!.StartsWith("stream band ", StringComparison.Ordinal)));
            Assert.Contains(projectedStream.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "stream band Enterprise" &&
                    item.GetProperty("geometry").GetProperty("kind").GetString() == "custom");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "audience composition stream" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedParticipantPictograph = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "participant pictograph bar");
            Assert.Equal(10, projectedParticipantPictograph.GetProperty("elements").EnumerateArray().Count(item =>
                item.GetProperty("name").GetString()!.StartsWith("pictographic symbol ", StringComparison.Ordinal)));
            Assert.Contains(projectedParticipantPictograph.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "pictographic unit" &&
                    item.GetProperty("text").GetProperty("paragraphs")[0].GetProperty("runs")[0]
                        .GetProperty("text").GetString() == "1 symbol = 10 participants");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "participant pictograph bar" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedMilestonePictograph = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "milestone pictograph column");
            Assert.Equal(9, projectedMilestonePictograph.GetProperty("elements").EnumerateArray().Count(item =>
                item.GetProperty("name").GetString()!.StartsWith("pictographic symbol ", StringComparison.Ordinal)));
            Assert.Contains(projectedMilestonePictograph.GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "pictographic symbol 2.1" &&
                    item.GetProperty("geometry").GetProperty("preset").GetString() == "star5");
            Assert.DoesNotContain(projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray(), item =>
                item.GetProperty("name").GetString() == "milestone pictograph column" &&
                    item.GetProperty("type").GetString() is "chart" or "image");
            var projectedAdjustedShape = projectedRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(item => item.GetProperty("type").GetString() == "group" &&
                    item.GetProperty("name").GetString() == "frame transform contract")
                .GetProperty("elements")[0];
            Assert.Equal("round2SameRect", projectedAdjustedShape.GetProperty("geometry").GetProperty("preset").GetString());
            Assert.Equal(18000, projectedAdjustedShape.GetProperty("geometry").GetProperty("adjustments")[0].GetInt32());
            Assert.Contains(projectedAdjustedShape.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "setGeometry" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "geometry.adjustments"));
            Assert.DoesNotContain("part_path", projected.PresentationProgram.ProgramJson.ToStringUtf8(), StringComparison.Ordinal);
            Assert.DoesNotContain("relationship_id", projected.PresentationProgram.ProgramJson.ToStringUtf8(), StringComparison.Ordinal);
            Assert.DoesNotContain("raw_xml", projected.PresentationProgram.ProgramJson.ToStringUtf8(), StringComparison.Ordinal);
        }
        var circularProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var circularChart = circularProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "doughnut");
        circularChart["style"]!["startAngle"] = 210;
        circularChart["style"]!["holeSize"] = 74;
        var circularChartId = circularChart["id"]!.GetValue<string>();
        var circularEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(circularProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(circularEdit.Ok, Diagnostics(circularEdit));
        Assert.Single(circularEdit.PresentationProgram.ChangedParts);
        Assert.Contains("/charts/chart", circularEdit.PresentationProgram.ChangedParts[0], StringComparison.Ordinal);
        Assert.EndsWith(".xml", circularEdit.PresentationProgram.ChangedParts[0], StringComparison.Ordinal);
        Assert.Contains(circularChartId, circularEdit.PresentationProgram.ChangedNodeIds);
        var circularReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = circularEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/circular-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(circularReprojection.Ok, Diagnostics(circularReprojection));
        using (var circularJson = JsonDocument.Parse(circularReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedCircular = circularJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == circularChartId);
            Assert.Equal(210, reprojectedCircular.GetProperty("style").GetProperty("startAngle").GetInt32());
            Assert.Equal(74, reprojectedCircular.GetProperty("style").GetProperty("holeSize").GetInt32());
        }
        var formattingProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var formattingChart = formattingProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart");
        formattingChart["style"]!["dataLabels"]!["numberFormat"] = "0.0";
        formattingChart["xAxis"]!["reverse"] = false;
        formattingChart["xAxis"]!["axisLine"]!["color"] = "#0B8F8F";
        formattingChart["xAxis"]!["axisLineArrow"]!["start"] = "none";
        formattingChart["xAxis"]!["axisLineArrow"]!["end"] = "diamond";
        formattingChart["yAxis"]!["gridLine"] = false;
        formattingChart["data"]!["series"]![1]!["dataLabels"]!["numberFormat"] = "0.00";
        formattingChart["data"]!["series"]![1]!["dataLabels"]!["points"]![1]!["numberFormat"] = "$0.0";
        formattingChart["data"]!["series"]![0]!["pointStyles"]![0]!["fill"]!["color"] = "#C1121F";
        formattingChart["data"]!["series"]![0]!["pointStyles"]![0]!["stroke"]!["width"] = 2;
        var formattingChartId = formattingChart["id"]!.GetValue<string>();
        var formattingBubble = formattingProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "bubble");
        formattingBubble["style"]!["bubbleScale"] = 180;
        formattingBubble["style"]!["bubbleSizeMode"] = "area";
        var formattingBubbleId = formattingBubble["id"]!.GetValue<string>();
        var formattingEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(formattingProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(formattingEdit.Ok, Diagnostics(formattingEdit));
        Assert.Equal(2, formattingEdit.PresentationProgram.ChangedParts.Count);
        Assert.All(formattingEdit.PresentationProgram.ChangedParts, part => Assert.Contains("/charts/chart", part, StringComparison.Ordinal));
        Assert.Contains(formattingChartId, formattingEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(formattingBubbleId, formattingEdit.PresentationProgram.ChangedNodeIds);
        var formattingReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = formattingEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/chart-formatting-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(formattingReprojection.Ok, Diagnostics(formattingReprojection));
        using (var formattingJson = JsonDocument.Parse(formattingReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedFormattingChart = formattingJson.RootElement.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == formattingChartId);
            Assert.Equal("0.0", reprojectedFormattingChart.GetProperty("style").GetProperty("dataLabels").GetProperty("numberFormat").GetString());
            Assert.False(reprojectedFormattingChart.GetProperty("xAxis").GetProperty("reverse").GetBoolean());
            Assert.Equal("#0B8F8F", reprojectedFormattingChart.GetProperty("xAxis").GetProperty("axisLine").GetProperty("color").GetString());
            Assert.Equal("none", reprojectedFormattingChart.GetProperty("xAxis").GetProperty("axisLineArrow").GetProperty("start").GetString());
            Assert.Equal("diamond", reprojectedFormattingChart.GetProperty("xAxis").GetProperty("axisLineArrow").GetProperty("end").GetString());
            Assert.False(reprojectedFormattingChart.GetProperty("yAxis").GetProperty("gridLine").GetBoolean());
            var reprojectedLabels = reprojectedFormattingChart.GetProperty("data").GetProperty("series")[1].GetProperty("dataLabels");
            Assert.Equal("0.00", reprojectedLabels.GetProperty("numberFormat").GetString());
            Assert.Equal("$0.0", reprojectedLabels.GetProperty("points")[1].GetProperty("numberFormat").GetString());
            var reprojectedPointStyle = reprojectedFormattingChart.GetProperty("data").GetProperty("series")[0].GetProperty("pointStyles")[0];
            Assert.Equal("#C1121F", reprojectedPointStyle.GetProperty("fill").GetProperty("color").GetString());
            Assert.Equal(2, reprojectedPointStyle.GetProperty("stroke").GetProperty("width").GetDouble());
            var reprojectedFormattingBubble = formattingJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == formattingBubbleId);
            Assert.Equal(180, reprojectedFormattingBubble.GetProperty("style").GetProperty("bubbleScale").GetInt32());
            Assert.Equal("area", reprojectedFormattingBubble.GetProperty("style").GetProperty("bubbleSizeMode").GetString());
        }
        var stateProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var stateClaim = stateProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "text" &&
                element["text"] is JsonObject text &&
                text["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>() == "Reduce incident hours ");
        stateClaim["hidden"] = false;
        stateClaim["locked"] = false;
        var stateClaimId = stateClaim["id"]!.GetValue<string>();
        stateProgram["pages"]![0]!["name"] = "Decision gate";
        stateProgram["pages"]![0]!["hidden"] = true;
        var statePageId = stateProgram["pages"]![0]!["id"]!.GetValue<string>();
        var stateEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(stateProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(stateEdit.Ok, Diagnostics(stateEdit));
        Assert.Equal(["ppt/slides/slide1.xml"], stateEdit.PresentationProgram.ChangedParts);
        Assert.Contains(stateClaimId, stateEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(statePageId, stateEdit.PresentationProgram.ChangedNodeIds);
        using (var stateStream = new MemoryStream(stateEdit.File.ToByteArray(), writable: false))
        using (var statePackage = PresentationDocument.Open(stateStream, false))
        {
            var unlockedClaim = statePackage.PresentationPart!.SlideParts.First().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().Single(shape => shape.Descendants<A.Text>().Any(text => text.Text == "Reduce incident hours "));
            Assert.Null(unlockedClaim.NonVisualShapeProperties!.NonVisualDrawingProperties!.Hidden);
            Assert.Null(unlockedClaim.NonVisualShapeProperties.NonVisualShapeDrawingProperties!.GetFirstChild<A.ShapeLocks>());
        }
        var stateReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = stateEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/state-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(stateReprojection.Ok, Diagnostics(stateReprojection));
        using (var stateJson = JsonDocument.Parse(stateReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            Assert.Equal("Decision gate", stateJson.RootElement.GetProperty("pages")[0].GetProperty("name").GetString());
            Assert.True(stateJson.RootElement.GetProperty("pages")[0].GetProperty("hidden").GetBoolean());
            var unlockedClaim = stateJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == stateClaimId);
            Assert.False(unlockedClaim.TryGetProperty("hidden", out _));
            Assert.False(unlockedClaim.TryGetProperty("locked", out _));
        }
        var unnamedThirdPartySource = ReplaceZipText(thirdPartySource, "ppt/slides/slide1.xml", xml =>
        {
            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            document.Descendants().First(element => element.Name.LocalName == "cSld").Attribute("name")?.Remove();
            return document.ToString(SaveOptions.DisableFormatting);
        });
        var unnamedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(unnamedThirdPartySource),
            PresentationProgram = new PresentationProgramRequest(),
        });
        Assert.True(unnamedProjection.Ok, Diagnostics(unnamedProjection));
        using (var unnamedJson = JsonDocument.Parse(unnamedProjection.PresentationProgram.ProgramJson.ToByteArray()))
            Assert.False(unnamedJson.RootElement.GetProperty("pages")[0].TryGetProperty("name", out _));
        var repeatedProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                IncludeNodeMap = true,
                SourceUri = "deck.assets/source/source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(repeatedProjection.Ok, Diagnostics(repeatedProjection));
        Assert.Equal(projected.PresentationProgram.ProgramJson, repeatedProjection.PresentationProgram.ProgramJson);
        Assert.Equal(projected.PresentationProgram.NodeMapJson, repeatedProjection.PresentationProgram.NodeMapJson);

        var sourceNoOp = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = projected.PresentationProgram.ProgramJson,
                IncludeNodeMap = true,
            },
        });
        Assert.True(sourceNoOp.Ok, Diagnostics(sourceNoOp));
        Assert.Equal(ByteString.CopyFrom(thirdPartySource), sourceNoOp.File);
        Assert.Empty(sourceNoOp.PresentationProgram.ChangedParts);
        Assert.Empty(sourceNoOp.PresentationProgram.ChangedNodeIds);

        var pairedSvgSource = AddPairedSvgFallback(thirdPartySource);
        var pairedSvgProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(pairedSvgSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/paired-svg-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(pairedSvgProjection.Ok, Diagnostics(pairedSvgProjection));
        string projectedSvgImageId;
        string projectedSvgFallbackAssetId;
        string projectedSvgAssetId;
        string projectedSvgFallbackSha256;
        using (var pairedSvgJson = JsonDocument.Parse(pairedSvgProjection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var pairedSvgImage = pairedSvgJson.RootElement.GetProperty("pages").EnumerateArray()
                .SelectMany(page => page.GetProperty("elements").EnumerateArray())
                .Single(item => item.GetProperty("type").GetString() == "image" && item.TryGetProperty("svgAsset", out _));
            projectedSvgImageId = pairedSvgImage.GetProperty("id").GetString()!;
            projectedSvgFallbackAssetId = pairedSvgImage.GetProperty("asset").GetString()!;
            projectedSvgAssetId = pairedSvgImage.GetProperty("svgAsset").GetString()!;
            Assert.Contains(pairedSvgImage.GetProperty("nativeRef").GetProperty("capabilities").EnumerateArray(), capability =>
                capability.GetProperty("operation").GetString() == "replaceSvg" &&
                capability.GetProperty("fields").EnumerateArray().Any(field => field.GetString() == "image.svgAsset"));
            var pairedSvgAssets = pairedSvgJson.RootElement.GetProperty("assets").EnumerateArray()
                .ToDictionary(asset => asset.GetProperty("id").GetString()!, StringComparer.Ordinal);
            Assert.Equal("image/png", pairedSvgAssets[projectedSvgFallbackAssetId].GetProperty("mimeType").GetString());
            Assert.Equal("image/svg+xml", pairedSvgAssets[projectedSvgAssetId].GetProperty("mimeType").GetString());
            projectedSvgFallbackSha256 = pairedSvgAssets[projectedSvgFallbackAssetId].GetProperty("sha256").GetString()!;
        }
        var pairedSvgNoOp = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(pairedSvgSource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = pairedSvgProjection.PresentationProgram.ProgramJson,
            },
        });
        Assert.True(pairedSvgNoOp.Ok, Diagnostics(pairedSvgNoOp));
        Assert.Equal(ByteString.CopyFrom(pairedSvgSource), pairedSvgNoOp.File);

        var replacementSvgBytes = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"32\" height=\"32\" viewBox=\"0 0 32 32\"><rect width=\"32\" height=\"32\" fill=\"#00A6A6\"/><path d=\"M6 16h20\" stroke=\"#FFFFFF\" stroke-width=\"3\"/></svg>");
        var replacementSvgSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(replacementSvgBytes)).ToLowerInvariant();
        const string replacementSvgAssetId = "replacement-paired-svg";
        const string replacementSvgUri = "deck.assets/media/replacement-paired-svg.svg";
        var svgProgram = JsonNode.Parse(pairedSvgProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var svgImage = svgProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element.ContainsKey("svgAsset"));
        Assert.Equal(projectedSvgImageId, svgImage["id"]!.GetValue<string>());
        var originalSvgDeclaration = svgProgram["assets"]!.AsArray()
            .Select(asset => asset!.AsObject())
            .Single(asset => asset["id"]!.GetValue<string>() == projectedSvgAssetId);
        var replacementSvgDeclaration = originalSvgDeclaration.DeepClone().AsObject();
        replacementSvgDeclaration["id"] = replacementSvgAssetId;
        replacementSvgDeclaration["uri"] = replacementSvgUri;
        replacementSvgDeclaration["sha256"] = replacementSvgSha256;
        svgProgram["assets"]!.AsArray().Add(replacementSvgDeclaration);
        svgImage["svgAsset"] = replacementSvgAssetId;
        var svgEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(pairedSvgSource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(svgProgram.ToJsonString()),
                IncludeNodeMap = true,
                Assets =
                {
                    new Asset
                    {
                        Id = replacementSvgAssetId,
                        FileName = replacementSvgUri,
                        ContentType = "image/svg+xml",
                        Data = ByteString.CopyFrom(replacementSvgBytes),
                        Sha256 = replacementSvgSha256,
                    },
                },
            },
        });
        Assert.True(svgEdit.Ok, Diagnostics(svgEdit));
        Assert.Contains(projectedSvgImageId, svgEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(svgEdit.PresentationProgram.ChangedParts, part =>
            part.Equals("ppt/slides/slide1.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(svgEdit.PresentationProgram.ChangedParts, part =>
            part.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase) && part.EndsWith(".svg", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(svgEdit.PresentationProgram.ChangedParts, part =>
            part.Equals("ppt/slides/slide2.xml", StringComparison.OrdinalIgnoreCase) ||
            part.Equals("ppt/slides/slide3.xml", StringComparison.OrdinalIgnoreCase));
        var svgReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = svgEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/svg-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(svgReprojection.Ok, Diagnostics(svgReprojection));
        using (var svgJson = JsonDocument.Parse(svgReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedSvgImage = svgJson.RootElement.GetProperty("pages").EnumerateArray()
                .SelectMany(page => page.GetProperty("elements").EnumerateArray())
                .Single(element => element.GetProperty("id").GetString() == projectedSvgImageId);
            Assert.Equal(projectedSvgFallbackAssetId, reprojectedSvgImage.GetProperty("asset").GetString());
            var reprojectedAssets = svgJson.RootElement.GetProperty("assets").EnumerateArray()
                .ToDictionary(asset => asset.GetProperty("id").GetString()!, StringComparer.Ordinal);
            Assert.Equal(projectedSvgFallbackSha256,
                reprojectedAssets[reprojectedSvgImage.GetProperty("asset").GetString()!].GetProperty("sha256").GetString());
            Assert.Equal(replacementSvgSha256,
                reprojectedAssets[reprojectedSvgImage.GetProperty("svgAsset").GetString()!].GetProperty("sha256").GetString());
        }
        var reorderProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var sourcePageNodes = reorderProgram["pages"]!.AsArray()
            .Select(page => page!.DeepClone()).ToArray();
        var sourcePageIds = sourcePageNodes.Select(page => page!["id"]!.GetValue<string>()).ToArray();
        var sourceElementIds = sourcePageNodes.ToDictionary(
            page => page!["id"]!.GetValue<string>(),
            page => page!["elements"]!.AsArray().Select(element => element!["id"]!.GetValue<string>()).ToArray(),
            StringComparer.Ordinal);
        var sourceCommentPageId = reorderProgram["comments"]![0]!["page"]!.GetValue<string>();
        var sourceShowPages = reorderProgram["customShows"]![0]!["pages"]!.AsArray()
            .Select(page => page!.GetValue<string>()).ToArray();
        var sourceCanvasHeight = reorderProgram["design"]!["canvas"]!["height"]!.GetValue<double>();
        var requestedCanvasWidth = reorderProgram["design"]!["canvas"]!["width"]!.GetValue<double>() + 72;
        reorderProgram["design"]!["canvas"]!["width"] = requestedCanvasWidth;
        reorderProgram["pages"] = new JsonArray(sourcePageNodes[2], sourcePageNodes[0], sourcePageNodes[1]);
        reorderProgram["sections"]![0]!["pages"] = new JsonArray(sourcePageIds[2], sourcePageIds[0]);
        reorderProgram["sections"]![1]!["pages"] = new JsonArray(sourcePageIds[1]);
        var reorderedPageIds = new[] { sourcePageIds[2], sourcePageIds[0], sourcePageIds[1] };
        var reorderEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(reorderProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(reorderEdit.Ok, Diagnostics(reorderEdit));
        Assert.Equal(["ppt/presentation.xml"], reorderEdit.PresentationProgram.ChangedParts);
        Assert.All(reorderedPageIds, id => Assert.Contains(id, reorderEdit.PresentationProgram.ChangedNodeIds));
        var reorderReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = reorderEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/reordered-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(reorderReprojection.Ok, Diagnostics(reorderReprojection));
        using (var reorderJson = JsonDocument.Parse(reorderReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reorderedCanvas = reorderJson.RootElement.GetProperty("design").GetProperty("canvas");
            Assert.Equal(requestedCanvasWidth, reorderedCanvas.GetProperty("width").GetDouble());
            Assert.Equal(sourceCanvasHeight, reorderedCanvas.GetProperty("height").GetDouble());
            var reorderedPages = reorderJson.RootElement.GetProperty("pages");
            Assert.Equal(reorderedPageIds, reorderedPages.EnumerateArray().Select(page => page.GetProperty("id").GetString()));
            foreach (var page in reorderedPages.EnumerateArray())
            {
                var pageId = page.GetProperty("id").GetString()!;
                Assert.Equal(sourceElementIds[pageId], page.GetProperty("elements").EnumerateArray().Select(element => element.GetProperty("id").GetString()));
            }
            Assert.Equal(
                new[] { sourcePageIds[2], sourcePageIds[0] },
                reorderJson.RootElement.GetProperty("sections")[0].GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
            Assert.Equal(
                new[] { sourcePageIds[1] },
                reorderJson.RootElement.GetProperty("sections")[1].GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
            Assert.Equal(sourceCommentPageId, reorderJson.RootElement.GetProperty("comments")[0].GetProperty("page").GetString());
            Assert.Equal(
                sourceShowPages,
                reorderJson.RootElement.GetProperty("customShows")[0].GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
        }

        var routeProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var routePageIds = routeProgram["pages"]!.AsArray()
            .Select(page => page!["id"]!.GetValue<string>()).ToArray();
        var firstSection = routeProgram["sections"]![0]!.AsObject();
        firstSection["name"] = "Decision and evidence";
        firstSection["pages"] = new JsonArray(routePageIds[0], routePageIds[1]);
        var secondSection = routeProgram["sections"]![1]!.AsObject();
        secondSection["pages"] = new JsonArray(routePageIds[2]);
        var editedShow = routeProgram["customShows"]![0]!.AsObject();
        editedShow["name"] = "Executive evidence route";
        editedShow["pages"] = new JsonArray(routePageIds[1], routePageIds[0], routePageIds[1]);
        var firstSectionId = firstSection["id"]!.GetValue<string>();
        var secondSectionId = secondSection["id"]!.GetValue<string>();
        var editedShowId = editedShow["id"]!.GetValue<string>();
        var routeEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(routeProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(routeEdit.Ok, Diagnostics(routeEdit));
        Assert.Equal(["ppt/presentation.xml"], routeEdit.PresentationProgram.ChangedParts);
        Assert.Contains(firstSectionId, routeEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(secondSectionId, routeEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(editedShowId, routeEdit.PresentationProgram.ChangedNodeIds);
        var routeReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = routeEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/route-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(routeReprojection.Ok, Diagnostics(routeReprojection));
        using (var routeJson = JsonDocument.Parse(routeReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var sections = routeJson.RootElement.GetProperty("sections");
            Assert.Equal("Decision and evidence", sections[0].GetProperty("name").GetString());
            Assert.Equal(routePageIds[..2], sections[0].GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
            Assert.Equal([routePageIds[2]], sections[1].GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
            var show = Assert.Single(routeJson.RootElement.GetProperty("customShows").EnumerateArray());
            Assert.Equal("Executive evidence route", show.GetProperty("name").GetString());
            Assert.Equal(
                new[] { routePageIds[1], routePageIds[0], routePageIds[1] },
                show.GetProperty("pages").EnumerateArray().Select(page => page.GetString()));
        }

        var notesProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        notesProgram["pages"]![1]!["notes"]!["paragraphs"]![0]!["runs"]![1]!["text"] = "independently verified";
        notesProgram["pages"]![2]!["notes"] = "Close with the accountable decision owner.";
        var editedNotesPageId = notesProgram["pages"]![1]!["id"]!.GetValue<string>();
        var addedNotesPageId = notesProgram["pages"]![2]!["id"]!.GetValue<string>();
        var notesEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(notesProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(notesEdit.Ok, Diagnostics(notesEdit));
        Assert.Contains(notesEdit.PresentationProgram.ChangedParts, part =>
            part.Equals("ppt/notesSlides/notesSlide2.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(notesEdit.PresentationProgram.ChangedParts, part =>
            part.StartsWith("ppt/notesSlides/notesSlide", StringComparison.OrdinalIgnoreCase) &&
            !part.Equals("ppt/notesSlides/notesSlide2.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(editedNotesPageId, notesEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(addedNotesPageId, notesEdit.PresentationProgram.ChangedNodeIds);
        var notesReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = notesEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/notes-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(notesReprojection.Ok, Diagnostics(notesReprojection));
        using (var notesJson = JsonDocument.Parse(notesReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var notesPages = notesJson.RootElement.GetProperty("pages");
            var editedRuns = notesPages[1].GetProperty("notes").GetProperty("paragraphs")[0].GetProperty("runs");
            Assert.Equal("independently verified", editedRuns[1].GetProperty("text").GetString());
            Assert.True(editedRuns[1].GetProperty("style").GetProperty("bold").GetBoolean());
            Assert.Equal("#A83232", editedRuns[1].GetProperty("style").GetProperty("color").GetString());
            Assert.Equal("Close with the accountable decision owner.", notesPages[2].GetProperty("notes").GetString());
        }

        var transitionProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        transitionProgram["pages"]![2]!["transition"] = new JsonObject
        {
            ["type"] = "wheel",
            ["spokes"] = 6,
            ["speed"] = "medium",
            ["advanceOnClick"] = true,
        };
        var transitionPageId = transitionProgram["pages"]![2]!["id"]!.GetValue<string>();
        var transitionEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(transitionProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(transitionEdit.Ok, Diagnostics(transitionEdit));
        Assert.Equal(["ppt/slides/slide3.xml"], transitionEdit.PresentationProgram.ChangedParts);
        Assert.Contains(transitionPageId, transitionEdit.PresentationProgram.ChangedNodeIds);
        var transitionReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = transitionEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/transition-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(transitionReprojection.Ok, Diagnostics(transitionReprojection));
        using (var transitionJson = JsonDocument.Parse(transitionReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedTransition = transitionJson.RootElement.GetProperty("pages")[2].GetProperty("transition");
            Assert.Equal("wheel", reprojectedTransition.GetProperty("type").GetString());
            Assert.Equal(6, reprojectedTransition.GetProperty("spokes").GetInt32());
            Assert.Equal("medium", reprojectedTransition.GetProperty("speed").GetString());
            Assert.True(reprojectedTransition.GetProperty("advanceOnClick").GetBoolean());
        }

        var richTitleProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var richTitleChart = richTitleProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart");
        richTitleChart["title"]!["paragraphs"]![0]!["runs"]![1]!["text"] = "−42% incidents";
        richTitleChart["title"]!["paragraphs"]![0]!["runs"]![1]!["style"]!["color"] = "#A83232";
        var richTitleChartId = richTitleChart["id"]!.GetValue<string>();
        var editedComment = richTitleProgram["comments"]![0]!.AsObject();
        editedComment["text"] = "Replace illustrative values with independently verified evidence.";
        var editedCommentId = editedComment["id"]!.GetValue<string>();
        var richTitleEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(richTitleProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(richTitleEdit.Ok, Diagnostics(richTitleEdit));
        Assert.Equal(2, richTitleEdit.PresentationProgram.ChangedParts.Count);
        var richTitlePart = Assert.Single(richTitleEdit.PresentationProgram.ChangedParts, part =>
            part.StartsWith("ppt/slides/charts/chart", StringComparison.OrdinalIgnoreCase));
        Assert.StartsWith("ppt/slides/charts/chart", richTitlePart, StringComparison.Ordinal);
        Assert.Single(richTitleEdit.PresentationProgram.ChangedParts, part =>
            part.StartsWith("ppt/comments/comment", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(richTitleChartId, richTitleEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(editedCommentId, richTitleEdit.PresentationProgram.ChangedNodeIds);
        var richTitleReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = richTitleEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/rich-title-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(richTitleReprojection.Ok, Diagnostics(richTitleReprojection));
        using (var richTitleJson = JsonDocument.Parse(richTitleReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedTitle = richTitleJson.RootElement.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == richTitleChartId)
                .GetProperty("title").GetProperty("paragraphs")[0].GetProperty("runs");
            Assert.Equal("−42% incidents", reprojectedTitle[1].GetProperty("text").GetString());
            Assert.Equal("#A83232", reprojectedTitle[1].GetProperty("style").GetProperty("color").GetString());
            Assert.Equal(
                "Replace illustrative values with independently verified evidence.",
                richTitleJson.RootElement.GetProperty("comments")[0].GetProperty("text").GetString());
        }

        var changedSourceLayoutProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        changedSourceLayoutProgram["pages"]![0]!["layout"] = "layout-invented";
        var changedSourceLayout = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(changedSourceLayoutProgram.ToJsonString()),
            },
        });
        Assert.False(changedSourceLayout.Ok);
        Assert.Equal("ppj.source.unsupportedMutation", Assert.Single(changedSourceLayout.Diagnostics).Code);

        var missingPointEditProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var missingPointChart = missingPointEditProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "line");
        missingPointChart["data"]!["series"]![0]!["values"]![0] = 95;
        var missingPointEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(missingPointEditProgram.ToJsonString()),
            },
        });
        Assert.True(missingPointEdit.Ok, Diagnostics(missingPointEdit));
        var missingPointReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = missingPointEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/missing-point-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(missingPointReprojection.Ok, Diagnostics(missingPointReprojection));
        using (var missingPointJson = JsonDocument.Parse(missingPointReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedValues = missingPointJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("type").GetString() == "chart" &&
                    element.GetProperty("chartType").GetString() == "line")
                .GetProperty("data").GetProperty("series")[0].GetProperty("values");
            Assert.Equal(95, reprojectedValues[0].GetDouble());
            Assert.Equal(JsonValueKind.Null, reprojectedValues[1].ValueKind);
        }

        var rejectedMissingTopologyProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var rejectedMissingTopologyChart = rejectedMissingTopologyProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "line");
        rejectedMissingTopologyChart["data"]!["series"]![0]!["values"]![1] = 106;
        var rejectedMissingTopology = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(rejectedMissingTopologyProgram.ToJsonString()),
            },
        });
        Assert.False(rejectedMissingTopology.Ok);
        Assert.Contains(rejectedMissingTopology.Diagnostics, diagnostic => diagnostic.Code == "ppj.source.unsupportedMutation");

        var chartTextStyleProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var chartTextStyleChart = chartTextStyleProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "line");
        chartTextStyleChart["style"]!["titleTextStyle"] = new JsonObject
        {
            ["fontSize"] = 19,
            ["fontFamily"] = "Aptos Display",
            ["fontFamilyEastAsia"] = "Noto Sans CJK SC",
            ["bold"] = false,
            ["italic"] = true,
            ["color"] = "#C0404080",
        };
        var chartHierarchyChart = chartTextStyleProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart");
        chartHierarchyChart["style"]!["legendTextStyle"] = new JsonObject
        {
            ["fontSize"] = 12,
            ["fontFamily"] = "Aptos Display",
            ["bold"] = true,
            ["color"] = "#0B8F8F",
        };
        chartHierarchyChart["style"]!["dataLabels"]!["textStyle"] = new JsonObject
        {
            ["fontSize"] = 9.5,
            ["fontFamily"] = "Aptos",
            ["italic"] = true,
            ["color"] = "#C0404080",
        };
        chartHierarchyChart["xAxis"]!["textStyle"]!["fontSize"] = 10;
        chartHierarchyChart["xAxis"]!["titleTextStyle"] = new JsonObject
        {
            ["fontSize"] = 12.5,
            ["fontFamily"] = "Georgia",
            ["color"] = "#16324F",
        };
        chartHierarchyChart["style"]!["chartAreaFill"] = new JsonObject
        {
            ["type"] = "gradient",
            ["kind"] = "linear",
            ["angle"] = 30,
            ["stops"] = new JsonArray
            {
                new JsonObject { ["offset"] = 0, ["color"] = "#FFFFFF" },
                new JsonObject { ["offset"] = 1, ["color"] = "#F2C14E", ["opacity"] = 0.5 },
            },
        };
        chartHierarchyChart["style"]!["plotAreaFill"] = new JsonObject { ["type"] = "none" };
        chartHierarchyChart["data"]!["series"]![0]!["fill"] = new JsonObject { ["type"] = "none" };
        var chartRadar = chartTextStyleProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "radar");
        chartRadar["spokeAxis"]!["max"] = 120;
        chartRadar["spokeAxis"]!["label"] = new JsonObject
        {
            ["numberFormat"] = "0.0",
            ["fontSize"] = 8.5,
            ["color"] = "#475569",
        };
        chartRadar["spokeAxis"]!["gridLine"] = new JsonObject
        {
            ["color"] = "#94A3B8",
            ["width"] = 0.75,
            ["dash"] = "dash",
        };
        var chartTextStyleId = chartTextStyleChart["id"]!.GetValue<string>();
        var chartHierarchyId = chartHierarchyChart["id"]!.GetValue<string>();
        var chartRadarId = chartRadar["id"]!.GetValue<string>();
        var chartTextStyleEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(chartTextStyleProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(chartTextStyleEdit.Ok, Diagnostics(chartTextStyleEdit));
        Assert.Equal(3, chartTextStyleEdit.PresentationProgram.ChangedParts.Count);
        Assert.Contains("ppt/slides/charts/chart2.xml", chartTextStyleEdit.PresentationProgram.ChangedParts);
        Assert.Contains("ppt/slides/charts/chart6.xml", chartTextStyleEdit.PresentationProgram.ChangedParts);
        Assert.Contains(chartTextStyleId, chartTextStyleEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(chartHierarchyId, chartTextStyleEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(chartRadarId, chartTextStyleEdit.PresentationProgram.ChangedNodeIds);
        var chartTextStyleReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = chartTextStyleEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/chart-text-style-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(chartTextStyleReprojection.Ok, Diagnostics(chartTextStyleReprojection));
        using (var chartTextStyleJson = JsonDocument.Parse(chartTextStyleReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedStyle = chartTextStyleJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == chartTextStyleId)
                .GetProperty("style").GetProperty("titleTextStyle");
            Assert.Equal(19, reprojectedStyle.GetProperty("fontSize").GetDouble());
            Assert.Equal("Aptos Display", reprojectedStyle.GetProperty("fontFamily").GetString());
            Assert.Equal("Noto Sans CJK SC", reprojectedStyle.GetProperty("fontFamilyEastAsia").GetString());
            Assert.True(reprojectedStyle.TryGetProperty("bold", out var reprojectedBold));
            Assert.False(reprojectedBold.GetBoolean());
            Assert.True(reprojectedStyle.GetProperty("italic").GetBoolean());
            Assert.Equal("#C0404080", reprojectedStyle.GetProperty("color").GetString());

            var hierarchy = chartTextStyleJson.RootElement.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == chartHierarchyId);
            Assert.Equal(12, hierarchy.GetProperty("style").GetProperty("legendTextStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal("#0B8F8F", hierarchy.GetProperty("style").GetProperty("legendTextStyle").GetProperty("color").GetString());
            Assert.Equal(9.5, hierarchy.GetProperty("style").GetProperty("dataLabels").GetProperty("textStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal("#C0404080", hierarchy.GetProperty("style").GetProperty("dataLabels").GetProperty("textStyle").GetProperty("color").GetString());
            Assert.Equal(10, hierarchy.GetProperty("xAxis").GetProperty("textStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal(12.5, hierarchy.GetProperty("xAxis").GetProperty("titleTextStyle").GetProperty("fontSize").GetDouble());
            Assert.Equal("linear", hierarchy.GetProperty("style").GetProperty("chartAreaFill").GetProperty("kind").GetString());
            Assert.Equal(30, hierarchy.GetProperty("style").GetProperty("chartAreaFill").GetProperty("angle").GetDouble());
            Assert.Equal("none", hierarchy.GetProperty("style").GetProperty("plotAreaFill").GetProperty("type").GetString());
            Assert.Equal("none", hierarchy.GetProperty("data").GetProperty("series")[0].GetProperty("fill").GetProperty("type").GetString());

            var radar = chartTextStyleJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == chartRadarId)
                .GetProperty("spokeAxis");
            Assert.Equal(120, radar.GetProperty("max").GetDouble());
            Assert.Equal("0.0", radar.GetProperty("label").GetProperty("numberFormat").GetString());
            Assert.Equal(8.5, radar.GetProperty("label").GetProperty("fontSize").GetDouble());
            Assert.Equal("#475569", radar.GetProperty("label").GetProperty("color").GetString());
            Assert.Equal("#94A3B8", radar.GetProperty("gridLine").GetProperty("color").GetString());
        }

        var imagePaintProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var imagePaintBackground = imagePaintProgram["pages"]![1]!["background"]!.AsObject();
        imagePaintBackground["opacity"] = 0.55;
        var imagePaintShape = imagePaintProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["name"]!.GetValue<string>() == "decision-flow-start");
        imagePaintShape["style"]!["fill"]!["opacity"] = 0.44;
        var tiledImage = imagePaintProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "image" &&
                element["mask"]?["kind"]?.GetValue<string>() == "preset");
        tiledImage["fit"] = "stretch";
        var imagePaintPageId = imagePaintProgram["pages"]![1]!["id"]!.GetValue<string>();
        var imagePaintShapeId = imagePaintShape["id"]!.GetValue<string>();
        var tiledImageId = tiledImage["id"]!.GetValue<string>();
        var imagePaintEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(imagePaintProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(imagePaintEdit.Ok, Diagnostics(imagePaintEdit));
        Assert.Equal(2, imagePaintEdit.PresentationProgram.ChangedParts.Count);
        Assert.Contains(imagePaintPageId, imagePaintEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(imagePaintShapeId, imagePaintEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(tiledImageId, imagePaintEdit.PresentationProgram.ChangedNodeIds);
        var imagePaintReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = imagePaintEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/image-paint-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(imagePaintReprojection.Ok, Diagnostics(imagePaintReprojection));
        using (var imagePaintJson = JsonDocument.Parse(imagePaintReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var imagePaintRoot = imagePaintJson.RootElement;
            Assert.Equal(0.55, imagePaintRoot.GetProperty("pages")[1].GetProperty("background").GetProperty("opacity").GetDouble(), 3);
            var reprojectedImageFill = imagePaintRoot.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("name").GetString() == "decision-flow-start")
                .GetProperty("style").GetProperty("fill");
            Assert.Equal(0.44, reprojectedImageFill.GetProperty("opacity").GetDouble(), 3);
            var reprojectedImage = imagePaintRoot.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("type").GetString() == "image" &&
                    element.GetProperty("mask").GetProperty("kind").GetString() == "preset");
            Assert.Equal("stretch", reprojectedImage.GetProperty("fit").GetString());
        }

        var translucentSolidBackgroundProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        translucentSolidBackgroundProgram["pages"]![1]!["background"] = new JsonObject
        {
            ["type"] = "solid",
            ["color"] = "#102030",
            ["opacity"] = 0.33,
        };
        var translucentSolidBackgroundEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(translucentSolidBackgroundProgram.ToJsonString()),
            },
        });
        Assert.True(translucentSolidBackgroundEdit.Ok, Diagnostics(translucentSolidBackgroundEdit));
        Assert.Contains("ppt/slides/slide2.xml", translucentSolidBackgroundEdit.PresentationProgram.ChangedParts);
        var translucentSolidBackgroundReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = translucentSolidBackgroundEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/translucent-solid-background-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(translucentSolidBackgroundReprojection.Ok, Diagnostics(translucentSolidBackgroundReprojection));
        using (var translucentSolidBackgroundJson = JsonDocument.Parse(translucentSolidBackgroundReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var background = translucentSolidBackgroundJson.RootElement.GetProperty("pages")[1].GetProperty("background");
            Assert.Equal("solid", background.GetProperty("type").GetString());
            Assert.Equal(0.33, background.GetProperty("opacity").GetDouble(), 3);
        }

        var rejectedTextOpacityProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var rejectedTextOpacity = rejectedTextOpacityProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "text" &&
                element["text"] is JsonObject text &&
                text["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>() == "Reduce incident hours ");
        rejectedTextOpacity["text"]!["paragraphs"]![0]!["runs"]![0]!["style"]!["gradient"]!["stops"]![0]!["opacity"] = 0.6;
        var rejectedTextOpacityEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(rejectedTextOpacityProgram.ToJsonString()),
            },
        });
        Assert.False(rejectedTextOpacityEdit.Ok);
        Assert.Contains(rejectedTextOpacityEdit.Diagnostics, diagnostic => diagnostic.Code == "ppj.source.unsupportedMutation");
        var rejectedChartStyleProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var rejectedChart = rejectedChartStyleProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "line");
        rejectedChart["style"]!["smooth"] = true;
        var rejectedChartStyleEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(rejectedChartStyleProgram.ToJsonString()),
            },
        });
        Assert.False(rejectedChartStyleEdit.Ok);
        Assert.Contains(rejectedChartStyleEdit.Diagnostics, diagnostic => diagnostic.Code == "ppj.source.unsupportedMutation");
        var rejectedBubbleProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var rejectedBubble = rejectedBubbleProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart" &&
                element["chartType"]!.GetValue<string>() == "bubble");
        rejectedBubble["data"]!["series"]![0]!["bubbleSizes"]![1] = 12;
        var rejectedBubbleEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(rejectedBubbleProgram.ToJsonString()),
            },
        });
        Assert.False(rejectedBubbleEdit.Ok);
        Assert.Contains(rejectedBubbleEdit.Diagnostics, diagnostic => diagnostic.Code == "ppj.source.unsupportedMutation");
        var sourceValidationRequest = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = projected.PresentationProgram.ProgramJson,
                IncludeNodeMap = true,
                ValidationOnly = true,
            },
        };
        var sourceValidation = Invoke(sourceValidationRequest);
        Assert.True(sourceValidation.Ok, Diagnostics(sourceValidation));
        Assert.Empty(sourceValidation.File);
        Assert.True(sourceValidation.PresentationProgram.SourceBound);
        Assert.Equal(projected.PresentationProgram.ProgramSha256, sourceValidation.PresentationProgram.ProgramSha256);

        var transformedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var transformedChart = transformedProgram["pages"]![1]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "chart");
        transformedChart["frame"]!["rotation"] = 9;
        transformedChart["frame"]!["flipH"] = false;
        var transformedChartId = transformedChart["id"]!.GetValue<string>();
        var sourceTransformEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(transformedProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(sourceTransformEdit.Ok, Diagnostics(sourceTransformEdit));
        Assert.Single(sourceTransformEdit.PresentationProgram.ChangedParts);
        Assert.Contains(transformedChartId, sourceTransformEdit.PresentationProgram.ChangedNodeIds);
        var sourceTransformReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = sourceTransformEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/transformed.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(sourceTransformReprojection.Ok, Diagnostics(sourceTransformReprojection));
        using (var transformedJson = JsonDocument.Parse(sourceTransformReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedChart = transformedJson.RootElement.GetProperty("pages")[1].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == transformedChartId);
            var reprojectedFrame = reprojectedChart.GetProperty("frame");
            Assert.Equal(9, reprojectedFrame.GetProperty("rotation").GetDouble());
            Assert.False(reprojectedFrame.TryGetProperty("flipH", out var flipHorizontal) && flipHorizontal.GetBoolean());
        }

        var adjustedGeometryProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var adjustedGeometryShape = adjustedGeometryProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "group" &&
                element["name"]!.GetValue<string>() == "frame transform contract")["elements"]![0]!.AsObject();
        adjustedGeometryShape["geometry"]!["adjustments"]![0] = 22000;
        var adjustedGeometryId = adjustedGeometryShape["id"]!.GetValue<string>();
        var sourceGeometryEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(adjustedGeometryProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(sourceGeometryEdit.Ok, Diagnostics(sourceGeometryEdit));
        Assert.Single(sourceGeometryEdit.PresentationProgram.ChangedParts);
        Assert.Contains(adjustedGeometryId, sourceGeometryEdit.PresentationProgram.ChangedNodeIds);
        var geometryReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = sourceGeometryEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/adjusted-geometry.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(geometryReprojection.Ok, Diagnostics(geometryReprojection));
        using (var geometryJson = JsonDocument.Parse(geometryReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedShape = geometryJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("type").GetString() == "group" &&
                    element.GetProperty("name").GetString() == "frame transform contract")
                .GetProperty("elements")[0];
            Assert.Equal(22000, reprojectedShape.GetProperty("geometry").GetProperty("adjustments")[0].GetInt32());
        }

        var changedCustomMaskProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var changedCustomMaskImage = changedCustomMaskProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["name"]?.GetValue<string>() == "irregular editorial crop");
        Assert.Contains(changedCustomMaskImage["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setImageMask");
        changedCustomMaskImage["mask"]!["paths"]![0]!["commands"]![1]!["x"] = 150;
        var customMaskEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(changedCustomMaskProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(customMaskEdit.Ok, Diagnostics(customMaskEdit));
        Assert.Single(customMaskEdit.PresentationProgram.ChangedParts);
        var customMaskReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = customMaskEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/custom-mask-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(customMaskReprojection.Ok, Diagnostics(customMaskReprojection));
        using (var customMaskJson = JsonDocument.Parse(customMaskReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedImage = customMaskJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == changedCustomMaskImage["id"]!.GetValue<string>());
            Assert.Equal(150, reprojectedImage.GetProperty("mask").GetProperty("paths")[0]
                .GetProperty("commands")[1].GetProperty("x").GetDouble());
        }

        var adjustedMaskProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var adjustedMaskImage = adjustedMaskProgram["pages"]![0]!["elements"]!.AsArray()
            .Select(element => element!.AsObject())
            .Single(element => element["type"]!.GetValue<string>() == "image" &&
                element["mask"]?["kind"]?.GetValue<string>() == "preset");
        Assert.Contains(adjustedMaskImage["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setImageMask");
        adjustedMaskImage["mask"]!["adjustments"]![0] = 32000;
        var adjustedMaskId = adjustedMaskImage["id"]!.GetValue<string>();
        var sourceMaskEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(adjustedMaskProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(sourceMaskEdit.Ok, Diagnostics(sourceMaskEdit));
        Assert.Single(sourceMaskEdit.PresentationProgram.ChangedParts);
        Assert.Contains(adjustedMaskId, sourceMaskEdit.PresentationProgram.ChangedNodeIds);
        var maskReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = sourceMaskEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/adjusted-mask.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(maskReprojection.Ok, Diagnostics(maskReprojection));
        using (var maskJson = JsonDocument.Parse(maskReprojection.PresentationProgram.ProgramJson.ToByteArray()))
        {
            var reprojectedImage = maskJson.RootElement.GetProperty("pages")[0].GetProperty("elements").EnumerateArray()
                .Single(element => element.GetProperty("id").GetString() == adjustedMaskId);
            Assert.Equal(32000, reprojectedImage.GetProperty("mask").GetProperty("adjustments")[0].GetInt32());
        }

        var editedProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var editableText = editedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .First(element => element["nativeRef"]!["capabilities"]!.AsArray()
                .Any(capability => capability!["operation"]!.GetValue<string>() == "replaceText"));
        const string sourceBoundReplacement = "PPJ source-bound evidence";
        var sourceBoundOriginal = editableText["text"] is JsonValue textValue
            ? textValue.GetValue<string>()
            : editableText["text"]!["paragraphs"]![0]!["runs"]![0]!["text"]!.GetValue<string>();
        if (editableText["text"] is JsonValue)
            editableText["text"] = sourceBoundReplacement;
        else
            editableText["text"]!["paragraphs"]![0]!["runs"]![0]!["text"] = sourceBoundReplacement;
        var editedTextId = editableText["id"]!.GetValue<string>();
        var sourceEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(editedProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(sourceEdit.Ok, Diagnostics(sourceEdit));
        Assert.NotEqual(ByteString.CopyFrom(thirdPartySource), sourceEdit.File);
        var changedTextPart = Assert.Single(sourceEdit.PresentationProgram.ChangedParts);
        Assert.Contains(editedTextId, sourceEdit.PresentationProgram.ChangedNodeIds);
        var sourceTextXml = Encoding.UTF8.GetString(ZipBytes(thirdPartySource, changedTextPart));
        var editedTextXml = Encoding.UTF8.GetString(ZipBytes(sourceEdit.File.ToByteArray(), changedTextPart));
        Assert.Equal(sourceTextXml, editedTextXml.Replace(
            $"<a:t>{sourceBoundReplacement}</a:t>",
            $"<a:t>{sourceBoundOriginal}</a:t>",
            StringComparison.Ordinal));
        var sourceEditRoundTrip = Import(sourceEdit.File.ToByteArray());
        Assert.True(sourceEditRoundTrip.Ok, Diagnostics(sourceEditRoundTrip));
        Assert.Contains(sourceEditRoundTrip.Artifact.Presentation.Slides.SelectMany(slide => slide.Elements), element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Shape && element.Shape.Text.Contains(sourceBoundReplacement, StringComparison.Ordinal));

        var gradientProgram = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var gradientShape = gradientProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["name"]!.GetValue<string>() == "claim-rule");
        Assert.Contains(gradientShape["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "setFill");
        gradientShape["style"]!["fill"]!["stops"]![1]!["color"] = "#2255AA";
        gradientShape["style"]!["fill"]!["stops"]![1]!["opacity"] = 0.64;
        var gradientEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(thirdPartySource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(gradientProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(gradientEdit.Ok, Diagnostics(gradientEdit));
        Assert.Single(gradientEdit.PresentationProgram.ChangedParts);
        Assert.Contains(gradientShape["id"]!.GetValue<string>(), gradientEdit.PresentationProgram.ChangedNodeIds);
        var gradientReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = gradientEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/gradient-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(gradientReprojection.Ok, Diagnostics(gradientReprojection));
        var gradientReprojectedProgram = JsonNode.Parse(gradientReprojection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var gradientReprojectedShape = gradientReprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["name"]!.GetValue<string>() == "claim-rule");
        Assert.Equal("#2255AA", gradientReprojectedShape["style"]!["fill"]!["stops"]![1]!["color"]!.GetValue<string>());
        Assert.Equal(0.64, gradientReprojectedShape["style"]!["fill"]!["stops"]![1]!["opacity"]!.GetValue<double>());

        // One integrated native-leaf contract is enough here: projection
        // issues the scalar, PPJ changes only value, the source-bound compiler
        // lowers it through the mature Edit Plan, and a fresh projection sees
        // the new value without exposing package paths or raw XML.
        var nativeLeafSource = ReplaceZipText(thirdPartySource, "ppt/slides/slide1.xml", xml =>
        {
            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace presentation = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var pictureGeometry = document.Descendants(presentation + "pic").First()
                .Descendants(drawing + "prstGeom").Single();
            pictureGeometry.Element(drawing + "avLst")?.RemoveNodes();
            var shadowShapeProperties = document.Descendants(presentation + "sp")
                .Select(shape => shape.Element(presentation + "spPr"))
                .OfType<XElement>()
                .First(properties => properties.Element(drawing + "solidFill")?.Element(drawing + "srgbClr") is not null &&
                    properties.Element(drawing + "effectLst") is null);
            shadowShapeProperties.Elements(drawing + "ln").Remove();
            shadowShapeProperties.Add(new XElement(drawing + "ln",
                new XAttribute("w", "9525"),
                new XAttribute("cap", "flat"),
                new XElement(drawing + "solidFill",
                    new XElement(drawing + "srgbClr", new XAttribute("val", "202020"))),
                new XElement(drawing + "prstDash", new XAttribute("val", "solid")),
                new XElement(drawing + "round")));
            shadowShapeProperties.Add(new XElement(drawing + "effectLst",
                new XElement(drawing + "outerShdw",
                    new XAttribute("blurRad", "142875"),
                    new XAttribute("dist", "95250"),
                    new XAttribute("dir", "2700000"),
                    new XAttribute("algn", "bl"),
                    new XAttribute("rotWithShape", "0"),
                    new XElement(drawing + "schemeClr",
                        new XAttribute("val", "dk1"),
                        new XElement(drawing + "alpha", new XAttribute("val", "43000"))))));
            return document.ToString(SaveOptions.DisableFormatting);
        });
        var nativeLeafProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeLeafSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/native-leaf-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(nativeLeafProjection.Ok, Diagnostics(nativeLeafProjection));
        var nativeLeafProgram = JsonNode.Parse(nativeLeafProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var nativeLeafOwner = nativeLeafProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .First(element => element["nativeRef"]?["leaves"]?.AsArray() is { } leaves &&
                leaves.Any(leaf => leaf!["kind"]!.GetValue<string>() == "fillRgb") &&
                leaves.Any(leaf => leaf!["kind"]!.GetValue<string>() == "widthEmu") &&
                leaves.Any(leaf => leaf!["kind"]!.GetValue<string>() == "shadowOpacityThousandthPercent"));
        var nativeLeaf = nativeLeafOwner["nativeRef"]!["leaves"]!.AsArray()
            .First(leaf => leaf!["kind"]!.GetValue<string>() == "fillRgb")!.AsObject();
        var nativeWidthLeaf = nativeLeafOwner["nativeRef"]!["leaves"]!.AsArray()
            .First(leaf => leaf!["kind"]!.GetValue<string>() == "widthEmu")!.AsObject();
        JsonObject ShadowLeaf(string kind) => nativeLeafOwner["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == kind)!.AsObject();
        var shadowBlurLeaf = ShadowLeaf("shadowBlurRadiusEmu");
        var shadowDistanceLeaf = ShadowLeaf("shadowDistanceEmu");
        var shadowDirectionLeaf = ShadowLeaf("shadowDirectionDegrees");
        var shadowAlignmentLeaf = ShadowLeaf("shadowAlignment");
        var shadowColorLeaf = ShadowLeaf("shadowColorScheme");
        var shadowOpacityLeaf = ShadowLeaf("shadowOpacityThousandthPercent");
        var lineStyleLeaf = ShadowLeaf("lineStyle");
        var lineCapLeaf = ShadowLeaf("lineCap");
        var lineJoinLeaf = ShadowLeaf("lineJoin");
        var highlightOwner = nativeLeafProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["nativeRef"]?["leaves"]?.AsArray().Any(leaf =>
                leaf!["kind"]!.GetValue<string>() == "fontHighlightRgb") == true);
        var highlightLeaf = highlightOwner["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "fontHighlightRgb")!.AsObject();
        var languageLeaf = highlightOwner["nativeRef"]!["leaves"]!.AsArray()
            .First(leaf => leaf!["kind"]!.GetValue<string>() == "fontLanguage" &&
                leaf["value"]!.GetValue<string>() == "zh-CN")!.AsObject();
        var imageLeafOwner = nativeLeafProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["nativeRef"]?["leaves"]?.AsArray() is { } leaves &&
                leaves.Any(leaf => leaf!["kind"]!.GetValue<string>() == "imageOpacityThousandthPercent") &&
                leaves.Any(leaf => leaf!["kind"]!.GetValue<string>() == "imageMaskPreset"));
        var imageOpacityLeaf = imageLeafOwner["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "imageOpacityThousandthPercent")!.AsObject();
        var imageMaskLeaf = imageLeafOwner["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "imageMaskPreset")!.AsObject();
        const string replacementFill = "#123456";
        const string replacementHighlight = "#F2C14E";
        const string replacementLanguage = "zh-TW";
        var replacementWidth = nativeWidthLeaf["value"]!.GetValue<long>() + 12_700;
        const int replacementImageOpacity = 61_000;
        const string replacementImageMask = "ellipse";
        const long replacementShadowBlur = 190_500;
        const long replacementShadowDistance = 127_000;
        const double replacementShadowDirection = 90;
        const string replacementShadowAlignment = "tr";
        const string replacementShadowColor = "accent1";
        const long replacementShadowOpacity = 33_000;
        const string replacementLineStyle = "dashed";
        const string replacementLineCap = "square";
        const string replacementLineJoin = "bevel";
        nativeLeaf["value"] = replacementFill;
        nativeWidthLeaf["value"] = replacementWidth;
        shadowBlurLeaf["value"] = replacementShadowBlur;
        shadowDistanceLeaf["value"] = replacementShadowDistance;
        shadowDirectionLeaf["value"] = replacementShadowDirection;
        shadowAlignmentLeaf["value"] = replacementShadowAlignment;
        shadowColorLeaf["value"] = replacementShadowColor;
        shadowOpacityLeaf["value"] = replacementShadowOpacity;
        lineStyleLeaf["value"] = replacementLineStyle;
        lineCapLeaf["value"] = replacementLineCap;
        lineJoinLeaf["value"] = replacementLineJoin;
        highlightLeaf["value"] = replacementHighlight;
        languageLeaf["value"] = replacementLanguage;
        imageOpacityLeaf["value"] = replacementImageOpacity;
        imageMaskLeaf["value"] = replacementImageMask;
        var nativeLeafEdit = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(nativeLeafSource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(nativeLeafProgram.ToJsonString()),
                IncludeNodeMap = true,
            },
        });
        Assert.True(nativeLeafEdit.Ok, Diagnostics(nativeLeafEdit));
        Assert.Single(nativeLeafEdit.PresentationProgram.ChangedParts);
        Assert.Contains(nativeLeafOwner["id"]!.GetValue<string>(), nativeLeafEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(highlightOwner["id"]!.GetValue<string>(), nativeLeafEdit.PresentationProgram.ChangedNodeIds);
        Assert.Contains(imageLeafOwner["id"]!.GetValue<string>(), nativeLeafEdit.PresentationProgram.ChangedNodeIds);
        var nativeLeafReprojection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = nativeLeafEdit.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/native-leaf-output.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(nativeLeafReprojection.Ok, Diagnostics(nativeLeafReprojection));
        var nativeLeafReprojectedProgram = JsonNode.Parse(nativeLeafReprojection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var reprojectedOwner = nativeLeafReprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == nativeLeafOwner["id"]!.GetValue<string>());
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "fillRgb" && leaf["value"]!.GetValue<string>() == replacementFill);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "widthEmu" && leaf["value"]!.GetValue<long>() == replacementWidth);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowBlurRadiusEmu" && leaf["value"]!.GetValue<long>() == replacementShadowBlur);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowDistanceEmu" && leaf["value"]!.GetValue<long>() == replacementShadowDistance);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowDirectionDegrees" && leaf["value"]!.GetValue<double>() == replacementShadowDirection);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowAlignment" && leaf["value"]!.GetValue<string>() == replacementShadowAlignment);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowColorScheme" && leaf["value"]!.GetValue<string>() == replacementShadowColor);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "shadowOpacityThousandthPercent" && leaf["value"]!.GetValue<long>() == replacementShadowOpacity);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "lineStyle" && leaf["value"]!.GetValue<string>() == replacementLineStyle);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "lineCap" && leaf["value"]!.GetValue<string>() == replacementLineCap);
        Assert.Contains(reprojectedOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "lineJoin" && leaf["value"]!.GetValue<string>() == replacementLineJoin);
        var reprojectedHighlightOwner = nativeLeafReprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == highlightOwner["id"]!.GetValue<string>());
        Assert.Contains(reprojectedHighlightOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "fontHighlightRgb" &&
            leaf["value"]!.GetValue<string>() == replacementHighlight.ToLowerInvariant());
        Assert.Contains(reprojectedHighlightOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "fontLanguage" &&
            leaf["value"]!.GetValue<string>() == replacementLanguage);
        var reprojectedImageOwner = nativeLeafReprojectedProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .Single(element => element["id"]!.GetValue<string>() == imageLeafOwner["id"]!.GetValue<string>());
        Assert.Contains(reprojectedImageOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "imageOpacityThousandthPercent" &&
            leaf["value"]!.GetValue<long>() == replacementImageOpacity);
        Assert.Contains(reprojectedImageOwner["nativeRef"]!["leaves"]!.AsArray(), leaf =>
            leaf!["kind"]!.GetValue<string>() == "imageMaskPreset" &&
            leaf["value"]!.GetValue<string>() == replacementImageMask);

        var tableSlidePath = ZipPartPaths(thirdPartySource).Single(path =>
            Regex.IsMatch(path, "^ppt/slides/slide[0-9]+\\.xml$", RegexOptions.CultureInvariant) &&
            Encoding.UTF8.GetString(ZipBytes(thirdPartySource, path)).Contains("<a:tbl", StringComparison.Ordinal));
        var opaqueTextSource = ReplaceZipText(thirdPartySource, tableSlidePath, xml =>
        {
            var tableStart = xml.IndexOf("<a:tbl", StringComparison.Ordinal);
            var tableEnd = xml.IndexOf("</a:tbl>", tableStart, StringComparison.Ordinal);
            var firstRunEnd = xml.IndexOf("</a:r>", tableStart, StringComparison.Ordinal);
            Assert.True(tableStart >= 0 && tableEnd > tableStart && firstRunEnd > tableStart && firstRunEnd < tableEnd);
            // A field run is intentionally outside the bounded plain-run
            // table profile.  The frame must remain opaque so a native text
            // replacement can preserve the source-owned companion field.
            // Keep all text leaves in the direct-run profile while adding an
            // unsupported table-level extension. The graphic frame therefore
            // falls back to opaque native text projection without flattening
            // any of the source-owned cell content.
            return xml.Insert(tableEnd, "<a:extLst/>");
        });
        var opaqueProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(opaqueTextSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/opaque-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(opaqueProjection.Ok, Diagnostics(opaqueProjection));
        var opaqueProgram = JsonNode.Parse(opaqueProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var opaqueText = opaqueProgram["pages"]!.AsArray()
            .SelectMany(page => page!["elements"]!.AsArray())
            .Select(element => element!.AsObject())
            .FirstOrDefault(element => element["type"]!.GetValue<string>() == "opaque" &&
                element["nativeRef"]!["capabilities"]!.AsArray().Any(capability =>
                    capability!["operation"]!.GetValue<string>() == "replaceText"));
        if (opaqueText is null)
        {
            Assert.DoesNotContain(opaqueProgram["pages"]!.AsArray()
                .SelectMany(page => page!["elements"]!.AsArray())
                .Select(element => element!.AsObject()),
                element => element["type"]!.GetValue<string>() == "table" &&
                    element["name"]!.GetValue<string>() == "method audit");
        }
        else
        {
            var opaqueTextId = opaqueText["id"]!.GetValue<string>();
            var oldNativeText = opaqueText["visibleText"]![0]!.GetValue<string>();
            const string newNativeText = "PPJ bounded native text";
            opaqueText["visibleText"]![0] = newNativeText;
            var opaqueTextEdit = Invoke(new CodecRequest
            {
                ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.CompilePpjToPptx,
                Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(opaqueTextSource),
                PresentationProgram = new PresentationProgramRequest
                {
                    ProgramJson = ByteString.CopyFromUtf8(opaqueProgram.ToJsonString()),
                },
            });
            Assert.True(opaqueTextEdit.Ok, Diagnostics(opaqueTextEdit));
            Assert.Equal([tableSlidePath], opaqueTextEdit.PresentationProgram.ChangedParts);
            Assert.Contains(opaqueTextId, opaqueTextEdit.PresentationProgram.ChangedNodeIds);
            var opaqueSourceXml = Encoding.UTF8.GetString(ZipBytes(opaqueTextSource, tableSlidePath));
            var opaqueOutputXml = Encoding.UTF8.GetString(ZipBytes(opaqueTextEdit.File.ToByteArray(), tableSlidePath));
            Assert.Equal(opaqueSourceXml, opaqueOutputXml.Replace(newNativeText, oldNativeText, StringComparison.Ordinal));
        }

        var deletionSource = ReplaceZipText(thirdPartySource, "ppt/slides/slide1.xml", xml => Regex.Replace(
            xml,
            "<p:timing\\b.*?</p:timing>",
            string.Empty,
            RegexOptions.Singleline | RegexOptions.CultureInvariant));
        var deletionProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(deletionSource),
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/deletion-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(deletionProjection.Ok, Diagnostics(deletionProjection));
        var deletionProgram = JsonNode.Parse(deletionProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var deletionPage = deletionProgram["pages"]!.AsArray().Select(page => page!.AsObject()).First(page =>
            page["elements"]!.AsArray().Any(element => element!["nativeRef"]!["capabilities"]!.AsArray().Any(capability =>
                capability!["operation"]!.GetValue<string>() == "delete")));
        var deletionElements = deletionPage["elements"]!.AsArray();
        var deletionIndex = deletionElements.Select((element, index) => (element, index)).First(item =>
            item.element!["nativeRef"]!["capabilities"]!.AsArray().Any(capability =>
                capability!["operation"]!.GetValue<string>() == "delete")).index;
        var deletedElementId = deletionElements[deletionIndex]!["id"]!.GetValue<string>();
        deletionElements.RemoveAt(deletionIndex);
        var deletion = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(deletionSource),
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(deletionProgram.ToJsonString()),
            },
        });
        Assert.True(deletion.Ok, Diagnostics(deletion));
        Assert.Contains(deletedElementId, deletion.PresentationProgram.ChangedNodeIds);
        Assert.Equal(deletionProjection.PresentationProgram.ExpandedElementCount - 1, deletion.PresentationProgram.ExpandedElementCount);

        var pageDeletionSourceRequest = ExportRequest();
        var secondPage = pageDeletionSourceRequest.Artifact.Presentation.Slides[0].Clone();
        secondPage.Id = "presentation/slide/2";
        secondPage.Name = "Companion";
        secondPage.Elements[0].Id = "presentation/slide/2/title";
        pageDeletionSourceRequest.Artifact.Presentation.Slides.Add(secondPage);
        var pageDeletionSource = Invoke(pageDeletionSourceRequest);
        Assert.True(pageDeletionSource.Ok, Diagnostics(pageDeletionSource));
        var pageDeletionProjection = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.ProjectPptxToPpj,
            Family = ArtifactFamily.Presentation,
            File = pageDeletionSource.File,
            PresentationProgram = new PresentationProgramRequest
            {
                SourceUri = "deck.assets/source/page-deletion-source.pptx",
                AssetRootUri = "deck.assets/media",
            },
        });
        Assert.True(pageDeletionProjection.Ok, Diagnostics(pageDeletionProjection));
        var pageDeletionProgram = JsonNode.Parse(pageDeletionProjection.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var pageDeletionPages = pageDeletionProgram["pages"]!.AsArray();
        var deletedPageId = pageDeletionPages[1]!["id"]!.GetValue<string>();
        Assert.Contains(pageDeletionPages[1]!["nativeRef"]!["capabilities"]!.AsArray(), capability =>
            capability!["operation"]!.GetValue<string>() == "delete");
        pageDeletionPages.RemoveAt(1);
        var pageDeletion = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            File = pageDeletionSource.File,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8(pageDeletionProgram.ToJsonString()),
            },
        });
        Assert.True(pageDeletion.Ok, Diagnostics(pageDeletion));
        Assert.Contains(deletedPageId, pageDeletion.PresentationProgram.ChangedNodeIds);
        var pageDeletionRoundTrip = Import(pageDeletion.File.ToByteArray());
        Assert.True(pageDeletionRoundTrip.Ok, Diagnostics(pageDeletionRoundTrip));
        Assert.Single(pageDeletionRoundTrip.Artifact.Presentation.Slides);

        var repeated = Invoke(request);
        Assert.True(repeated.Ok, Diagnostics(repeated));
        Assert.Equal(first.PresentationProgram.ProgramSha256, repeated.PresentationProgram.ProgramSha256);
        Assert.Equal(ZipPartPaths(first.File.ToByteArray()), ZipPartPaths(repeated.File.ToByteArray()));
        var differingParts = ZipPartPaths(first.File.ToByteArray())
            .Where(path => !ZipBytes(first.File.ToByteArray(), path).SequenceEqual(ZipBytes(repeated.File.ToByteArray(), path)))
            .ToArray();
        Assert.True(differingParts.Length == 0, $"Non-deterministic OPC parts: {string.Join(", ", differingParts)}");
        Assert.Equal(first.PresentationProgram.OutputSha256, repeated.PresentationProgram.OutputSha256);
        Assert.Equal(first.File, repeated.File);

        var editedArtifact = imported.Artifact.Clone();
        var editedImage = Assert.Single(editedArtifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Image && element.Name == "evidence identity").Image;
        editedImage.OpacityThousandthPercent = 76_000;
        editedImage.MaskPreset = "roundRect";
        editedImage.Border = null;
        editedImage.Shadow.DistanceEmu = 48_000;
        var edited = Export(editedArtifact);
        Assert.True(edited.Ok, Diagnostics(edited));
        var editedRoundTrip = Import(edited.File.ToByteArray());
        Assert.True(editedRoundTrip.Ok, Diagnostics(editedRoundTrip));
        var roundTripImage = Assert.Single(editedRoundTrip.Artifact.Presentation.Slides[0].Elements, element =>
            element.ContentCase == PresentationElement.ContentOneofCase.Image && element.Name == "evidence identity").Image;
        Assert.Equal(76_000U, roundTripImage.OpacityThousandthPercent);
        Assert.Equal("roundRect", roundTripImage.MaskPreset);
        Assert.Null(roundTripImage.Border);
        Assert.Equal(48_000L, roundTripImage.Shadow.DistanceEmu);
    }
}
