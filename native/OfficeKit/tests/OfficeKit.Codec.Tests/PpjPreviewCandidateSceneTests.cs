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
        foreach (var (element, index) in elements.Select((element, index) => (element, index)))
        {
            var binding = Assert.Single(preview.Program.PreviewScene.Bindings,
                binding => binding.SemanticId == element!["id"]!.GetValue<string>());
            Assert.Equal(PresentationPreviewAttribution.Direct, binding.Attribution);
            Assert.Equal(json["pages"]![0]!["id"]!.GetValue<string>(), binding.PageId);
            Assert.Equal($"$.pages[0].elements[{index}]", binding.ProgramPath);
            Assert.Equal($"$.presentation.slides[0].elements[{index}]", binding.ScenePath);
            Assert.Equal((uint)index, binding.ZOrder);
        }
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
    public void PpjPreviewCandidateKeepsSemanticOwnersWhenPageAndElementOrdinalsChange()
    {
        var source = PreviewSource(secondPage: true);
        using var projected = PpjPresentationProjector.Project(source,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/reorder.pptx" }, EffectiveCodecLimits.From(null));
        var json = JsonNode.Parse(projected.Program.ProgramJson.ToStringUtf8())!;
        var pages = json["pages"]!.AsArray();
        var second = pages[1]!.DeepClone();
        var first = pages[0]!.DeepClone();
        pages.Clear(); pages.Add(second); pages.Add(first);
        var elements = second["elements"]!.AsArray();
        var originalFirst = elements[0]!.DeepClone();
        var originalSecond = elements[1]!.DeepClone();
        elements[0] = originalSecond; elements[1] = originalFirst;
        second["readingOrder"] = new JsonArray(elements.Select(element => JsonValue.Create(element!["id"]!.GetValue<string>())).ToArray());
        var request = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString()), IncludePreviewScene = true };
        var compiled = PpjPresentationCompiler.Compile(request, source, EffectiveCodecLimits.From(null));
        var scene = compiled.Program.PreviewScene;
        var binding = Assert.Single(scene.Bindings, binding => binding.ScenePath == "$.presentation.slides[0].elements[0]");
        Assert.Equal(second["id"]!.GetValue<string>(), binding.PageId);
        Assert.Equal(originalSecond["id"]!.GetValue<string>(), binding.SemanticId);
        Assert.Equal("$.pages[0].elements[0]", binding.ProgramPath);
        Assert.Equal("336699", scene.Presentation.Slides[0].Elements[0].Shape.FillRgb);
        var movedText = Assert.Single(scene.Bindings, binding => binding.ScenePath == "$.presentation.slides[0].elements[1]");
        Assert.Equal(originalFirst["id"]!.GetValue<string>(), movedText.SemanticId);
        Assert.Equal("$.pages[0].elements[1]", movedText.ProgramPath);
        Assert.Equal("Second page claim", scene.Presentation.Slides[0].Elements[1].Shape.Text);
        AssertCandidateScene(compiled.Program, compiled.File);
        // The same native cNvPr IDs occur on both pages. Part identity must
        // distinguish them rather than creating an accidental cross-page join.
        var otherText = Assert.Single(scene.Bindings, binding => binding.ScenePath == "$.presentation.slides[1].elements[0]");
        Assert.Equal(first["elements"]![0]!["id"]!.GetValue<string>(), otherText.SemanticId);
        Assert.NotEqual(movedText.PageId, otherText.PageId);
    }

    [Fact]
    public void PpjPreviewNativeIdentityCollectionIsOptInAndAmbiguityDoesNotGrantOwnership()
    {
        var source = PreviewSource();
        var limits = EffectiveCodecLimits.From(null);
        var plain = PptxCodec.Import(source, limits);
        var captured = PptxCodec.Import(source, limits, includeNativeBindings: true);
        Assert.Empty(plain.NativeBindings);
        Assert.Equal(plain.Artifact, captured.Artifact);
        Assert.Equal(4, captured.NativeBindings.Count);
        using var projected = PpjPresentationProjector.Project(source,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/identity.pptx" }, limits,
            includeNativeBindings: true);
        Assert.Equal(4, projected.NativeBindings.Count);
        var identity = captured.NativeBindings[0];
        var ownership = new PpjPreviewCandidateBindings(projected.NativeBindings, projected.Validation!.Expansion!);
        Assert.NotNull(ownership.Owner(identity));
        var duplicate = projected.NativeBindings.Concat([projected.NativeBindings[0] with { ElementId = projected.NativeBindings[1].ElementId }]).ToArray();
        Assert.Null(new PpjPreviewCandidateBindings(duplicate, projected.Validation.Expansion!).Owner(identity));
        Assert.Null(ownership.Owner(identity with { NativeId = 0 }));
        Assert.Null(ownership.Owner(identity with { PartPath = "ppt/slides/not-the-source.xml" }));
    }

    [Fact]
    public void PpjPreviewNestedGroupBindsRequestedPathsSeparatelyFromNativeReadingOrder()
    {
        var source = PreviewSource(nestedGroup: true);
        var original = source.ToArray();
        using var projected = PpjPresentationProjector.Project(source,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/group.pptx" }, EffectiveCodecLimits.From(null));
        var json = JsonNode.Parse(projected.Program.ProgramJson.ToStringUtf8())!;
        var group = json["pages"]![0]!["elements"]![3]!;
        Assert.Equal("group", group["type"]!.GetValue<string>());
        var children = group["elements"]!.AsArray();
        var first = children[0]!["id"]!.GetValue<string>();
        var second = children[1]!["id"]!.GetValue<string>();
        Assert.Equal(new[] { first, second }, group["readingOrder"]!.AsArray()
            .Select(node => node!.GetValue<string>()).ToArray());
        group["readingOrder"] = new JsonArray(second, first);
        var compiled = PpjPresentationCompiler.Compile(new PresentationProgramRequest
        {
            ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString()), IncludePreviewScene = true,
        }, source, EffectiveCodecLimits.From(null));
        var scene = compiled.Program.PreviewScene;
        var groupBinding = Assert.Single(scene.Bindings, binding => binding.SemanticId == group["id"]!.GetValue<string>());
        Assert.Equal("$.pages[0].elements[3]", groupBinding.ProgramPath);
        var firstPainted = Assert.Single(scene.Bindings, binding => binding.SemanticId == second);
        Assert.Equal("$.pages[0].elements[3].elements[1]", firstPainted.ProgramPath);
        Assert.Equal("$.presentation.slides[0].elements[3].group.children[0]", firstPainted.ScenePath);
        Assert.Equal(0u, firstPainted.ZOrder);
        Assert.Equal(PresentationPreviewAttribution.Direct, firstPainted.Attribution);
        var secondPainted = Assert.Single(scene.Bindings, binding => binding.SemanticId == first);
        Assert.Equal("$.pages[0].elements[3].elements[0]", secondPainted.ProgramPath);
        Assert.Equal("$.presentation.slides[0].elements[3].group.children[1]", secondPainted.ScenePath);
        Assert.Equal(1u, secondPainted.ZOrder);
        Assert.Equal("Nested second", scene.Presentation.Slides[0].Elements[3].Group.Children[0].Shape.Text);
        Assert.Equal(original, source);
        AssertCandidateScene(compiled.Program, compiled.File);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PpjPreviewDeletionAndNewOverlaysRespectIdentityBoundaries(bool delete)
    {
        var source = PreviewSource();
        var original = source.ToArray();
        using var projected = PpjPresentationProjector.Project(source,
            new PresentationProgramRequest { SourceUri = "deck.assets/source/overlay.pptx" }, EffectiveCodecLimits.From(null));
        var json = JsonNode.Parse(projected.Program.ProgramJson.ToStringUtf8())!;
        var page = json["pages"]![0]!;
        var elements = page["elements"]!.AsArray();
        var removed = elements[1]!["id"]!.GetValue<string>();
        // Deletion and appending overlays are separate compiler profiles.
        // An overlay requires the complete original source prefix.
        if (delete) elements.RemoveAt(1);
        else elements.Add(JsonNode.Parse("""
            {"id":"new-overlay","type":"text","frame":{"x":40,"y":400,"width":300,"height":40},"text":"New candidate overlay"}
            """));
        page["readingOrder"] = new JsonArray(elements.Select(element => JsonValue.Create(element!["id"]!.GetValue<string>())).ToArray());
        var compiled = PpjPresentationCompiler.Compile(new PresentationProgramRequest
        {
            ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString()), IncludePreviewScene = true,
        }, source, EffectiveCodecLimits.From(null));
        var scene = compiled.Program.PreviewScene;
        if (delete)
        {
            Assert.DoesNotContain(scene.Bindings, binding => binding.SemanticId == removed);
            Assert.Equal(3, scene.Presentation.Slides[0].Elements.Count);
            Assert.All(scene.Bindings, binding => Assert.Equal(PresentationPreviewAttribution.Direct, binding.Attribution));
            var movedImage = Assert.Single(scene.Bindings, binding => binding.SemanticId == elements[1]!["id"]!.GetValue<string>());
            Assert.Equal("$.pages[0].elements[1]", movedImage.ProgramPath);
            Assert.Equal("$.presentation.slides[0].elements[1]", movedImage.ScenePath);
        }
        else
        {
            var overlay = Assert.Single(scene.Presentation.Slides[0].Elements, element => element.Shape?.Text == "New candidate overlay");
            var binding = Assert.Single(scene.Bindings, binding => binding.NativeId == overlay.Id);
            Assert.Equal(PresentationPreviewAttribution.Unmapped, binding.Attribution);
            Assert.Equal(string.Empty, binding.SemanticId);
            Assert.Equal("$.pages[0]", binding.ProgramPath);
            Assert.Equal(page["id"]!.GetValue<string>(), binding.PageId);
            Assert.Equal("$.presentation.slides[0].elements[4]", binding.ScenePath);
        }
        Assert.Equal(original, source);
        AssertCandidateScene(compiled.Program, compiled.File);
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

    private static byte[] PreviewSource(bool keepEmbedded = false, bool secondPage = false, bool nestedGroup = false)
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
        if (nestedGroup) elements.Add(JsonNode.Parse("""
            {"id":"nested-group","type":"group","frame":{"x":400,"y":180,"width":300,"height":200},
             "childFrame":{"x":0,"y":0,"width":300,"height":200},
             "elements":[
               {"id":"nested-first","type":"text","frame":{"x":0,"y":0,"width":200,"height":40},"text":"Nested first"},
               {"id":"nested-second","type":"text","frame":{"x":0,"y":80,"width":200,"height":40},"text":"Nested second"}]}
            """));
        if (secondPage)
        {
            var page = json["pages"]![0]!.DeepClone();
            page["id"] = "second";
            foreach (var element in page["elements"]!.AsArray()) element!["id"] = element["id"]!.GetValue<string>() + "-second";
            page["elements"]![0]!["text"] = "Second page claim";
            json["pages"]!.AsArray().Add(page);
        }
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
