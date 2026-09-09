using System.Security.Cryptography;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PptxCodecTests
{
    [Theory]
    [InlineData("noop")]
    [InlineData("text")]
    [InlineData("frame")]
    [InlineData("semantic")]
    public void PpjPreviewCandidateMatchesFreshImportAfterEachCompilePath(string mode)
    {
        var source = PreviewSource();
        var original = source.ToArray();
        using var projected = PpjPresentationProjector.Project(source,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/candidate.pptx" }, EffectiveCodecLimits.From(null));
        var json = JsonNode.Parse(projected.Program.ProgramJson.ToStringUtf8())!;
        var elements = json["pages"]![0]!["elements"]!.AsArray();
        var text = elements.First(element => element?["text"] is not null)!;
        // The native writer supplies an empty text body even on this plain
        // rectangle; absence of a text field is not its imported identity.
        var box = elements[1]!;
        Assert.Equal("shape", box["type"]!.GetValue<string>());
        if (mode == "text")
        {
            if (text["text"] is JsonValue) text["text"] = "Candidate text changed";
            else text["text"]!["paragraphs"]![0]!["runs"]![0]!["text"] = "Candidate text changed";
        }
        if (mode == "frame") text["frame"]!["x"] = 93;
        if (mode == "semantic") box["style"]!["fill"] = JsonNode.Parse("""{"type":"solid","color":"#AB1234"}""");
        var request = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString()) };
        var ordinary = PpjPresentationCompiler.Compile(request, source, EffectiveCodecLimits.From(null));
        request.IncludePreviewScene = true;
        var requestedBytes = request.ToByteArray();
        var preview = PpjPresentationCompiler.Compile(request, source, EffectiveCodecLimits.From(null));
        Assert.Equal(ordinary.File, preview.File);
        Assert.Null(ordinary.Program.PreviewScene);
        Assert.Equal(requestedBytes, request.ToByteArray());
        Assert.Equal(original, source);
        Assert.Equal(PreviewSha(source), preview.Program.SourceSha256);
        AssertCandidateScene(preview.Program, preview.File);
        if (mode == "noop")
        {
            Assert.Equal(source, preview.File);
            Assert.Empty(preview.Program.ChangedParts);
        }
        else
        {
            Assert.NotEqual(PreviewSha(source), preview.Program.OutputSha256);
            Assert.Contains("ppt/slides/slide1.xml", preview.Program.ChangedParts);
            if (mode is "text" or "frame") Assert.Equal(new[] { "ppt/slides/slide1.xml" }, preview.Program.ChangedParts);
        }
        Assert.Equal(ZipPartPaths(source), ZipPartPaths(preview.File));
        foreach (var part in ZipPartPaths(source).Where(part => !preview.Program.ChangedParts.Contains(part)))
            Assert.Equal(ZipBytes(source, part), ZipBytes(preview.File, part));
        var before = PptxCodec.Import(source, EffectiveCodecLimits.From(null)).Artifact.Presentation;
        var scene = preview.Program.PreviewScene.Presentation;
        Assert.Equal(before.Slides[0].Elements.Select(element => element.ContentCase), scene.Slides[0].Elements.Select(element => element.ContentCase));
        var opaque = Assert.Single(scene.Slides[0].Elements, element => element.Opaque is not null);
        Assert.NotEmpty(opaque.Opaque.NativeKind);
        Assert.Empty(opaque.Opaque.RawXml);
        var sourceOpaque = Assert.Single(before.Slides[0].Elements, element => element.Opaque is not null);
        var candidateOpaque = Assert.Single(PptxCodec.Import(preview.File, EffectiveCodecLimits.From(null)).Artifact.Presentation.Slides[0].Elements,
            element => element.Opaque is not null);
        Assert.Equal(sourceOpaque.Opaque.RawXml, candidateOpaque.Opaque.RawXml);
        if (mode == "text") Assert.Contains(scene.Slides[0].Elements, element => element.Shape?.Text == "Candidate text changed");
        if (mode == "frame") Assert.Equal(93 * 12700L, scene.Slides[0].Elements[0].Shape.LeftEmu);
        if (mode == "semantic") Assert.Contains(scene.Slides[0].Elements, element => element.Shape?.FillRgb == "AB1234");
    }

    [Fact]
    public void PpjPreviewCandidateDoesNotRestoreAnEmbeddedAuthoredSnapshot()
    {
        var candidate = PreviewSource(keepEmbedded: true);
        var original = candidate.ToArray();
        using var projected = PpjPresentationProjector.Project(candidate,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/embedded.pptx" }, EffectiveCodecLimits.From(null));
        Assert.True(projected.Program.RestoredEmbeddedProgram);
        Assert.Null(projected.SourceArtifact);
        var receipt = new PresentationProgramResult
        {
            ProgramSha256 = projected.Program.ProgramSha256, OutputSha256 = PreviewSha(candidate),
        };
        using var package = new PptxPackageSource(candidate);
        PpjPreviewCandidateScene.Attach(package, receipt, EffectiveCodecLimits.From(null));
        AssertCandidateScene(receipt, candidate);
        // Native import identities, not IDs from the restored PPJ program.
        Assert.NotEqual("opening", receipt.PreviewScene.Presentation.Slides[0].Id);
        Assert.Equal("Evidence changed the decision", receipt.PreviewScene.Presentation.Slides[0].Elements[0].Shape.Text);
        Assert.Equal(original, candidate);
    }

    [Fact]
    public void PpjPreviewFileBackedNoOpKeepsReuseAndSuppliesCandidateAssets()
    {
        var source = PreviewSource();
        var directory = Directory.CreateTempSubdirectory("officekit-preview-source-");
        var path = Path.Combine(directory.FullName, "original.pptx");
        try
        {
            File.WriteAllBytes(path, source);
            using var package = new PptxPackageSource(path);
            using var projected = PpjPresentationProjector.Project(package,
                new PresentationProgramRequest { SourceUri = "deck.assets/source/file-backed.pptx" },
                EffectiveCodecLimits.From(null), retainSourceAssetData: false);
            using var validation = PpjProgramValidator.Validate(projected.Program.ProgramJson.Memory);
            Assert.True(validation.IsValid);
            var request = new PresentationProgramRequest { ProgramJson = projected.Program.ProgramJson };
            var plain = PpjPresentationCompiler.CompileValidated(request, package, EffectiveCodecLimits.From(null), validation,
                retainSourceAssetData: false);
            request.IncludePreviewScene = true;
            var preview = PpjPresentationCompiler.CompileValidated(request, package, EffectiveCodecLimits.From(null), validation,
                retainSourceAssetData: false);
            Assert.True(plain.ReuseSourceFile);
            Assert.True(preview.ReuseSourceFile);
            Assert.Empty(preview.File);
            Assert.False(package.TryGetMaterialized(out _));
            AssertCandidateScene(preview.Program, source);
            Assert.Equal(source, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
            directory.Delete();
        }
    }

    private static void AssertCandidateScene(PresentationProgramResult receipt, byte[] candidate)
    {
        var scene = receipt.PreviewScene;
        Assert.NotNull(scene);
        Assert.Equal(PresentationPreviewSceneOrigin.CandidateImport, scene.Origin);
        Assert.Equal(PreviewSha(candidate), scene.CandidateSha256);
        Assert.Equal(receipt.OutputSha256, scene.CandidateSha256);
        Assert.Equal(receipt.ProgramSha256, scene.ProgramSha256);
        var imported = PptxCodec.Import(candidate, EffectiveCodecLimits.From(null)).Artifact;
        var independentlyImported = PpjPreviewSceneBuilder.Build(imported.Presentation,
            PresentationPreviewSceneOrigin.CandidateImport, receipt.ProgramSha256, receipt.OutputSha256,
            scene.Bindings, imported.Assets, EffectiveCodecLimits.From(null));
        Assert.Equal(independentlyImported, scene);
        Assert.NotEmpty(scene.Bindings);
        Assert.NotEmpty(scene.Assets);
        foreach (var reference in scene.Assets)
        {
            var asset = Assert.Single(receipt.Assets, asset =>
                asset.Sha256.Equals(reference.Sha256, StringComparison.OrdinalIgnoreCase) &&
                asset.ContentType.Equals(reference.ContentType, StringComparison.OrdinalIgnoreCase) && !asset.Data.IsEmpty);
            Assert.Equal(reference.Sha256, PreviewSha(asset.Data.Span));
        }
    }

    private static byte[] PreviewSource(bool keepEmbedded = false)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "package.json"))) root = root.Parent;
        Assert.NotNull(root);
        var json = JsonNode.Parse(File.ReadAllBytes(Path.Combine(root!.FullName, "examples/ppj/minimum.ppj")))!;
        var canonicalPath = Path.Combine(root.FullName, "test/fixtures/presentation/evidence-ledger-canonical.ppj");
        var assets = JsonNode.Parse(File.ReadAllBytes(canonicalPath))!["assets"]!.DeepClone();
        json["assets"] = assets;
        var elements = json["pages"]![0]!["elements"]!.AsArray();
        elements.Add(JsonNode.Parse("""
            {"id":"box","type":"shape","frame":{"x":60,"y":180,"width":150,"height":100},
             "geometry":{"kind":"preset","preset":"rect"},"style":{"fill":{"type":"solid","color":"#336699"}}}
            """));
        elements.Add(JsonNode.Parse("""
            {"id":"mark","type":"image","asset":"evidence-mark","frame":{"x":260,"y":180,"width":100,"height":100}}
            """));
        var request = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString()) };
        foreach (var asset in assets.AsArray())
        {
            var data = File.ReadAllBytes(Path.Combine(Path.GetDirectoryName(canonicalPath)!, asset!["uri"]!.GetValue<string>()));
            request.Assets.Add(new Asset { Id = asset["id"]!.GetValue<string>(), ContentType = asset["mimeType"]!.GetValue<string>(),
                Data = ByteString.CopyFrom(data), Sha256 = PreviewSha(data) });
        }
        var compiled = PpjPresentationCompiler.Compile(request, [], EffectiveCodecLimits.From(null));
        if (keepEmbedded) return compiled.File;
        var source = RemoveEmbeddedPpj(compiled.File);
        return ReplaceZipText(source, "ppt/slides/slide1.xml", xml =>
        {
            var document = XDocument.Parse(xml);
            var tree = document.Descendants().Single(element => element.Name.LocalName == "spTree");
            tree.Add(XElement.Parse("""
                <p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:nvGraphicFramePr><p:cNvPr id="900" name="Opaque sibling"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
                  <p:xfrm><a:off x="5000000" y="3000000"/><a:ext cx="1000000" cy="800000"/></p:xfrm>
                  <a:graphic><a:graphicData uri="urn:officekit:test:opaque"><payload xmlns="urn:officekit:test:opaque">keep-me</payload></a:graphicData></a:graphic>
                </p:graphicFrame>
                """));
            return document.ToString(SaveOptions.DisableFormatting);
        });
    }

    private static string PreviewSha(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
