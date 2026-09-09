using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
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
