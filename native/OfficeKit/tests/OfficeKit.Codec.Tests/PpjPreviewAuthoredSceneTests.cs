using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class PpjPreviewAuthoredSceneTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RealAuthoredCompilationReturnsResolvedSceneWithoutChangingCandidate(bool ppjOnly, bool canonical)
    {
        var fixture = Fixture(canonical ? "test/fixtures/presentation/evidence-ledger-canonical.ppj" : "examples/ppj/minimum.ppj");
        var request = Request(fixture);
        var original = request.ToByteArray();
        var ordinary = Invoke(request, ppjOnly);
        request.PresentationProgram.IncludePreviewScene = true;
        var withScene = Invoke(request, ppjOnly);
        Assert.Equal(ordinary.File, withScene.File);
        Assert.Equal(ordinary.PresentationProgram.ProgramJson, withScene.PresentationProgram.ProgramJson);
        Assert.Equal(ordinary.PresentationProgram.ProgramSha256, withScene.PresentationProgram.ProgramSha256);
        var scene = withScene.PresentationProgram.PreviewScene;
        Assert.NotNull(scene);
        Assert.Null(ordinary.PresentationProgram.PreviewScene);
        Assert.Equal(PresentationPreviewSceneOrigin.AuthoredLowering, scene.Origin);
        Assert.Equal(withScene.PresentationProgram.OutputSha256, scene.CandidateSha256);
        Assert.Equal(Sha(withScene.File.Span), scene.CandidateSha256);
        Assert.Equal(withScene.PresentationProgram.ProgramSha256, scene.ProgramSha256);
        Assert.Equal(canonical ? 2 : 1, scene.Presentation.Slides.Count);
        Assert.NotEmpty(scene.Bindings);
        request.PresentationProgram.IncludePreviewScene = false;
        Assert.Equal(original, request.ToByteArray());
        if (canonical)
        {
            var all = scene.Presentation.Slides.SelectMany(slide => Walk(slide.Elements)).ToArray();
            Assert.Contains(all, element => element.Id == "decision-flow-gate" && element.Shape?.Geometry == "flowChartDecision");
            Assert.Contains(all, element => element.Id == "evidence-chart-main" && element.Chart is not null);
            Assert.Contains(scene.Bindings, binding => binding.ComponentId.Length > 0);
            foreach (var image in all.Where(element => element.Image is not null))
            {
                var reference = Assert.Single(scene.Assets, asset => asset.NativeId == image.Image.AssetId);
                var bytes = Assert.Single(withScene.PresentationProgram.Assets, asset => asset.Id == reference.NativeId);
                Assert.Equal(Sha(bytes.Data.Span), reference.Sha256);
            }
        }
        else
        {
            var shape = scene.Presentation.Slides[0].Elements[0].Shape;
            Assert.Equal(48 * 12700L, shape.LeftEmu);
            Assert.Equal(42 * 12700L, shape.TopEmu);
            Assert.Equal("Evidence changed the decision", shape.Text);
        }
    }

    [Fact]
    public void RealAuthoredSceneRetainsMasterLayoutAndResolvedTextPrecedence()
    {
        var request = Request(Fixture("examples/ppj/minimum.ppj"));
        var json = JsonNode.Parse(request.PresentationProgram.ProgramJson.ToStringUtf8())!;
        json["design"]!["masters"] = JsonNode.Parse("""
            [{"id":"master","name":"Master","background":{"type":"solid","color":"#113355"},
              "style":{"size":19,"bold":true},
              "textStyles":{"title":[{"level":0,"alignment":"center"}]}}]
            """);
        json["design"]!["layouts"] = JsonNode.Parse("""
            [{"id":"layout","name":"Layout","master":"master","layoutType":"blank",
              "background":{"type":"solid","color":"#CCDDEE"},"style":{"size":31}}]
            """);
        json["pages"]![0]!["layout"] = "layout";
        json["design"]!["grammar"]!["stylePrecedence"] = JsonNode.Parse("""
            [{"target":"text.size","sources":["layout","master"]},
             {"target":"text.bold","sources":["layout","master"]}]
            """);
        request.PresentationProgram.ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString());
        var ordinary = Invoke(request, false);
        request.PresentationProgram.IncludePreviewScene = true;
        var compiled = Invoke(request, false);
        Assert.Equal(ordinary.File, compiled.File);
        var presentation = compiled.PresentationProgram.PreviewScene.Presentation;
        Assert.Equal("113355", Assert.Single(presentation.Masters).Background.ColorRgb);
        Assert.Equal("center", Assert.Single(presentation.Masters[0].TextStyles.TitleLevels).Alignment);
        Assert.Equal("CCDDEE", Assert.Single(presentation.Layouts).Background.ColorRgb);
        var slide = Assert.Single(presentation.Slides);
        Assert.Equal("layout", slide.LayoutId);
        var shape = Assert.Single(slide.Elements).Shape;
        var run = Assert.Single(Assert.Single(shape.TextBody.Paragraphs).Runs);
        Assert.Equal(31, run.FontSizePoints);
        Assert.True(run.Bold);
        Assert.NotNull(presentation.AuthoredTheme);
    }

    [Fact]
    public void UppercaseAssetHashDoesNotChangeAuthoredSceneOrCandidate()
    {
        var request = Request(Fixture("test/fixtures/presentation/evidence-ledger-canonical.ppj"));
        foreach (var asset in request.PresentationProgram.Assets) asset.Sha256 = asset.Sha256.ToUpperInvariant();
        var ordinary = Invoke(request, false);
        request.PresentationProgram.IncludePreviewScene = true;
        var compiled = Invoke(request, false);
        Assert.Equal(ordinary.File, compiled.File);
        Assert.All(compiled.PresentationProgram.PreviewScene.Assets,
            asset => Assert.Equal(asset.Sha256.ToLowerInvariant(), asset.Sha256));
        Assert.All(request.PresentationProgram.Assets, asset => Assert.Equal(asset.Sha256.ToUpperInvariant(), asset.Sha256));
    }

    [Fact]
    public void AuthoredVectorAndNativeChartSceneRetainsDifferentRepresentationsAndMissingIndexes()
    {
        var request = Request(Fixture("examples/ppj/minimum.ppj"));
        var json = JsonNode.Parse(request.PresentationProgram.ProgramJson.ToStringUtf8())!;
        var elements = json["pages"]![0]!["elements"]!.AsArray();
        elements.Add(JsonNode.Parse("""
            {"id":"heat","type":"chart","chartType":"heatmap","frame":{"x":40,"y":140,"width":400,"height":240},
             "style":{"heatmap":{"colors":["#FFFFFF","#114477"],"showColorBar":false}},
             "data":{"categories":["a","b","c"],"series":[{"id":"heat-a","name":"A","values":[1,null,0]},{"id":"heat-b","name":"B","values":[3,2,1]}]}}
            """));
        elements.Add(JsonNode.Parse("""
            {"id":"line","type":"chart","chartType":"line","frame":{"x":480,"y":140,"width":400,"height":240},
             "data":{"categories":["a","b","c"],"series":[{"id":"line-a","name":"A","values":[1,null,0]}]}}
            """));
        request.PresentationProgram.ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString());
        var ordinary = Invoke(request, false);
        request.PresentationProgram.IncludePreviewScene = true;
        var compiled = Invoke(request, false);
        Assert.Equal(ordinary.File, compiled.File);
        var scene = compiled.PresentationProgram.PreviewScene;
        var heat = Assert.Single(scene.Presentation.Slides[0].Elements, element => element.Id == "heat");
        Assert.NotNull(heat.Group);
        Assert.True(heat.Group.Children.Count > 6);
        Assert.Contains(heat.Group.Children, element => element.Shape?.Text == "a");
        var line = Assert.Single(scene.Presentation.Slides[0].Elements, element => element.Id == "line");
        Assert.Equal(new[] { 1d, 0d, 0d }, line.Chart.Series[0].Values);
        Assert.Equal(new uint[] { 1 }, line.Chart.Series[0].MissingValueIndexes);
        Assert.Contains(scene.Bindings, binding => binding.SemanticId == "heat" &&
            binding.Attribution == PresentationPreviewAttribution.Generated && binding.ScenePath.Contains(".group.children["));
        Assert.All(scene.Bindings.Where(binding => binding.SemanticId == "heat"), binding =>
        {
            Assert.Equal("$.pages[0].elements[1]", binding.ProgramPath);
            Assert.Equal("heat", binding.SourceId);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedRepeatAndReplacedSlotsKeepOriginalOwnersWithoutChangingExpansion(bool ppjOnly)
    {
        var request = NestedOriginRequest();
        var original = request.PresentationProgram.ProgramJson;
        var json = JsonNode.Parse(original.ToStringUtf8())!;
        using var ordinaryValidation = PpjProgramValidator.Validate(original.Memory);
        using var previewValidation = PpjProgramValidator.Validate(original.Memory, includePreviewOrigins: true);
        Assert.True(ordinaryValidation.IsValid, string.Join("\n", ordinaryValidation.Diagnostics));
        Assert.True(previewValidation.IsValid, string.Join("\n", previewValidation.Diagnostics));
        Assert.Null(ordinaryValidation.Expansion!.PreviewOrigins);
        Assert.Equal(ordinaryValidation.CanonicalJson, previewValidation.CanonicalJson);
        Assert.Equal(ordinaryValidation.Expansion.NodeMapJson, previewValidation.Expansion!.NodeMapJson);
        Assert.Equal(ordinaryValidation.Expansion.NodeMapSha256, previewValidation.Expansion.NodeMapSha256);
        Assert.Equal(previewValidation.Expansion.Nodes.Count, previewValidation.Expansion.PreviewOrigins!.Count);
        request.PresentationProgram.IncludeNodeMap = true;
        var ordinary = Invoke(request, ppjOnly);
        request.PresentationProgram.IncludePreviewScene = true;
        var compiled = Invoke(request, ppjOnly);
        Assert.Equal(ordinary.File, compiled.File);
        Assert.Equal(ordinary.PresentationProgram.NodeMapJson, compiled.PresentationProgram.NodeMapJson);
        Assert.Equal(original, request.PresentationProgram.ProgramJson);
        var scene = compiled.PresentationProgram.PreviewScene;
        Assert.All(scene.Bindings, binding =>
        {
            Assert.DoesNotContain("#component", binding.ProgramPath);
            Assert.Equal(binding.SourceId, ResolveOrigin(json, binding.ProgramPath)["id"]!.GetValue<string>());
            Assert.Equal(PresentationPreviewAttribution.Generated, binding.Attribution);
            Assert.Equal("opening", binding.PageId);
            Assert.Equal(binding.NativeId, ResolveSceneNode(scene, binding.ScenePath).Id);
        });
        Assert.Equal(scene.Bindings.Count, scene.Bindings.Select(binding => binding.ScenePath).Distinct().Count());
        var after = scene.Bindings.Where(binding => binding.SourceId == "after").ToArray();
        Assert.Equal(2, after.Length);
        // The two supplied nodes replace one nested slot. The following node
        // moves to scene index 3, but its original definition index stays 2.
        Assert.All(after, binding =>
        {
            Assert.Equal("$.components[1].elements[0].elements[2]", binding.ProgramPath);
            Assert.EndsWith(".group.children[3]", binding.ScenePath);
            Assert.Equal(3u, binding.ZOrder);
        });
        foreach (var (sourceId, slotIndex) in new[] { ("slot-a", 0), ("slot-b", 1) })
        {
            var supplied = scene.Bindings.Where(binding => binding.SourceId == sourceId).ToArray();
            Assert.Equal(2, supplied.Length);
            Assert.All(supplied, binding => Assert.Equal($"$.pages[0].elements[0].slots[\"body\"][{slotIndex}]", binding.ProgramPath));
        }
        var nestedLeaves = scene.Bindings.Where(binding => binding.SourceId == "leaf").ToArray();
        Assert.Equal(4, nestedLeaves.Length);
        Assert.All(nestedLeaves, binding => Assert.Equal("$.components[0].elements[0]", binding.ProgramPath));
        Assert.Equal(4, nestedLeaves.Select(binding => binding.InstanceId).Distinct().Count());
        Assert.Equal(new[] { "inner-one", "inner-two" }, nestedLeaves.Select(binding => binding.RepeatKey).Distinct().Order().ToArray());
        var nestedSupplied = scene.Bindings.Where(binding => binding.SourceId == "inner-supplied").ToArray();
        Assert.Equal(4, nestedSupplied.Length);
        Assert.All(nestedSupplied, binding => Assert.Equal(
            "$.components[1].elements[0].elements[3].slots[\"extra\"][0]", binding.ProgramPath));
        Assert.Equal(new[] { "outer-one", "outer-two" }, after.Select(binding => binding.RepeatKey).Order().ToArray());
    }

    [Fact]
    public void ValidatedComponentPreviewRequiresOriginalPassProvenance()
    {
        var request = NestedOriginRequest().PresentationProgram;
        using var validation = PpjProgramValidator.Validate(request.ProgramJson.Memory);
        Assert.True(validation.IsValid, string.Join("\n", validation.Diagnostics));
        request.IncludePreviewScene = true;
        var failure = Assert.Throws<CodecException>(() => PpjPresentationCompiler.CompileValidated(
            request, Array.Empty<byte>(), EffectiveCodecLimits.From(null), validation));
        Assert.Equal("ppj.preview.originsRequired", failure.Code);
    }

    private static CodecRequest NestedOriginRequest()
    {
        var request = Request(Fixture("examples/ppj/minimum.ppj"));
        var json = JsonNode.Parse(request.PresentationProgram.ProgramJson.ToStringUtf8())!;
        json["components"] = JsonNode.Parse("""
            [{"id":"inner","frame":{"x":0,"y":0,"width":120,"height":50},
              "slots":[{"name":"extra","accepts":["text"],"minItems":1,"maxItems":1}],
              "elements":[
                {"id":"leaf","type":"text","frame":{"x":0,"y":0,"width":60,"height":20},"text":"inner leaf"},
                {"id":"extra-slot","type":"slot","frame":{"x":60,"y":0,"width":60,"height":20},"slot":"extra"}]},
             {"id":"outer","frame":{"x":0,"y":0,"width":400,"height":240},
              "slots":[{"name":"body","accepts":["text"],"minItems":2,"maxItems":2}],
              "elements":[{"id":"container","type":"group","frame":{"x":0,"y":0,"width":400,"height":240},
                "childFrame":{"x":0,"y":0,"width":400,"height":240},"elements":[
                  {"id":"before","type":"text","frame":{"x":0,"y":0,"width":100,"height":20},"text":"before"},
                  {"id":"body-slot","type":"slot","frame":{"x":0,"y":30,"width":100,"height":40},"slot":"body"},
                  {"id":"after","type":"text","frame":{"x":0,"y":90,"width":100,"height":20},"text":"after"},
                  {"id":"nested","type":"component","component":"inner","frame":{"x":20,"y":130,"width":240,"height":100},
                   "repeat":{"items":[{"key":"inner-one","arguments":{}},{"key":"inner-two","arguments":{}}],"layout":{"direction":"vertical","gap":0}},
                   "slots":{"extra":[{"id":"inner-supplied","type":"text","frame":{"x":60,"y":0,"width":60,"height":20},"text":"supplied inside definition"}]}}
                ]}]}]
            """);
        json["pages"]![0]!["elements"] = JsonNode.Parse("""
            [{"id":"outer-instance","type":"component","component":"outer","frame":{"x":40,"y":30,"width":800,"height":240},
              "repeat":{"items":[{"key":"outer-one","arguments":{}},{"key":"outer-two","arguments":{}}],"layout":{"direction":"horizontal","gap":0}},
              "slots":{"body":[
                {"id":"slot-a","type":"text","frame":{"x":0,"y":30,"width":100,"height":20},"text":"first supplied"},
                {"id":"slot-b","type":"text","frame":{"x":0,"y":60,"width":100,"height":20},"text":"second supplied"}]}}]
            """);
        request.PresentationProgram.ProgramJson = ByteString.CopyFromUtf8(json.ToJsonString());
        return request;
    }

    private static JsonNode ResolveOrigin(JsonNode root, string path)
    {
        var tokens = Regex.Matches(path[1..], "\\.([a-zA-Z][a-zA-Z0-9]*)|\\[(\\d+|\"(?:[^\"\\\\]|\\\\.)*\")\\]");
        Assert.Equal(path[1..], string.Concat(tokens.Select(token => token.Value)));
        foreach (Match token in tokens)
        {
            if (token.Groups[1].Success) root = root[token.Groups[1].Value]!;
            else if (token.Groups[2].Value.StartsWith('"')) root = root[JsonSerializer.Deserialize<string>(token.Groups[2].Value)!]!;
            else root = root[int.Parse(token.Groups[2].Value)]!;
            Assert.NotNull(root);
        }
        return root;
    }

    private static PresentationElement ResolveSceneNode(PresentationPreviewScene scene, string path)
    {
        var indexes = Regex.Matches(path, "\\[(\\d+)\\]").Select(match => int.Parse(match.Groups[1].Value)).ToArray();
        var element = scene.Presentation.Slides[indexes[0]].Elements[indexes[1]];
        foreach (var index in indexes.Skip(2)) element = element.Group.Children[index];
        return element;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriterObserverCapturesExactInputOnceAndPreservesPreviousSlideLifetime(bool morph)
    {
        var plain = ExportObserved(morph, false);
        var preview = ExportObserved(morph, true);
        Assert.Equal(plain.File, preview.File);
        Assert.Equal(new[] { 1, 1, 1 }, plain.Plan.Materializations);
        Assert.Equal(plain.Plan.Materializations, preview.Plan.Materializations);
        Assert.Equal(new[] { false, morph, false }, plain.Plan.ReceivedPrevious);
        Assert.Equal(plain.Plan.ReceivedPrevious, preview.Plan.ReceivedPrevious);
        Assert.Equal(preview.Plan.WriterInputs, preview.Scene!.Presentation.Slides);
        Assert.Equal(preview.Plan.Presentation.Masters, preview.Scene.Presentation.Masters);
        Assert.Equal(preview.Plan.Presentation.Layouts, preview.Scene.Presentation.Layouts);
        Assert.Equal(preview.Plan.Presentation.AuthoredTheme, preview.Scene.Presentation.AuthoredTheme);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        Assert.All(plain.Plan.LiveSlides, reference => Assert.False(reference.TryGetTarget(out _)));
        Assert.All(preview.Plan.LiveSlides, reference => Assert.False(reference.TryGetTarget(out _)));
        Assert.All(preview.Plan.Presentation.Slides, slide => Assert.Empty(slide.Elements));
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (byte[] File, TrackingPlan Plan, PresentationPreviewScene? Scene) ExportObserved(bool morph, bool preview)
    {
        var plan = new TrackingPlan(morph);
        var wrapper = preview ? new PpjPreviewSourceFreeBuildPlan(plan, [], EffectiveCodecLimits.From(null)) : null;
        var exported = PptxCodec.ExportSourceFree((IPptxSourceFreeBuildPlan?)wrapper ?? plan, [],
            EffectiveCodecLimits.From(null), _ => { });
        return (exported.File, plan, wrapper?.Complete(new string('a', 64), Sha(exported.File), []));
    }

    private sealed class TrackingPlan : IPptxSourceFreeBuildPlan
    {
        private readonly bool morph;
        internal int[] Materializations { get; } = new int[3];
        internal bool[] ReceivedPrevious { get; } = new bool[3];
        internal List<PresentationSlide> WriterInputs { get; } = [];
        internal List<WeakReference<PresentationSlide>> LiveSlides { get; } = [];
        public PresentationArtifact Presentation { get; } = new()
        {
            Id = "observed", SlideWidthEmu = 12192000, SlideHeightEmu = 6858000,
            AuthoredTheme = new PresentationThemeArtifact { Name = "Preview theme" },
            Masters = { new PresentationMaster { Id = "master", Name = "Master" } },
            Layouts = { new PresentationLayout { Id = "layout", MasterId = "master", Name = "Layout", Type = "blank" } },
        };

        internal TrackingPlan(bool morph)
        {
            this.morph = morph;
            for (var i = 0; i < 3; i++) Presentation.Slides.Add(new PresentationSlide { Id = $"page-{i}", LayoutId = "layout" });
        }
        public bool RequiresPreviousSlide(int index) => morph && index == 1;
        public PresentationSlide MaterializeSlide(int index, PresentationSlide? previousSlide)
        {
            Materializations[index]++;
            ReceivedPrevious[index] = previousSlide is not null;
            if (previousSlide is not null) Assert.Equal("page-0", previousSlide.Id);
            var slide = Presentation.Slides[index].Clone();
            slide.Elements.Add(new PresentationElement { Id = $"shape-{index}", Name = "!!same",
                Shape = new PresentationShape { Geometry = "rect", WidthEmu = 1270000, HeightEmu = 1270000, FillRgb = "446688" } });
            if (morph && index == 1) slide.Morph = new PresentationMorph
            {
                FromSlideId = "page-0", DurationMs = 800,
                Pairs = { new PresentationMorphPair { Key = "same", FromId = "shape-0", ToId = "shape-1" } },
            };
            LiveSlides.Add(new(slide));
            return slide;
        }
        public void RecordNativeBindings(int index, PresentationSlide slide, IReadOnlyList<PresentationElement> flattenedElements)
        {
            Assert.True(LiveSlides[index].TryGetTarget(out var actual));
            Assert.Same(actual, slide);
            Assert.Same(slide.Elements[0], flattenedElements[0]);
            WriterInputs.Add(slide.Clone());
        }
    }

    private static string Fixture(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(directory!.FullName, relative);
    }

    private static CodecRequest Request(string file)
    {
        var json = File.ReadAllBytes(file);
        var node = JsonNode.Parse(json)!;
        var request = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion, Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFrom(json) },
        };
        foreach (var asset in node["assets"]?.AsArray() ?? [])
        {
            var bytes = File.ReadAllBytes(Path.Combine(Path.GetDirectoryName(file)!, asset!["uri"]!.GetValue<string>()));
            request.PresentationProgram.Assets.Add(new Asset
            {
                Id = asset["id"]!.GetValue<string>(), ContentType = asset["mimeType"]!.GetValue<string>(),
                Data = ByteString.CopyFrom(bytes), Sha256 = Sha(bytes),
            });
        }
        return request;
    }

    private static IEnumerable<PresentationElement> Walk(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.Group is not null) foreach (var child in Walk(element.Group.Children)) yield return child;
        }
    }

    private static CodecResponse Invoke(CodecRequest request, bool ppjOnly)
    {
        var bytes = request.ToByteArray();
        var response = ppjOnly ? PpjCodecProtocol.InvokeResponse(ref bytes, null) : CodecProtocol.InvokeResponse(ref bytes);
        Assert.True(response.Ok, string.Join("\n", response.Diagnostics.Select(d => $"{d.Code}: {d.Message}")));
        return response;
    }
    private static string Sha(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
