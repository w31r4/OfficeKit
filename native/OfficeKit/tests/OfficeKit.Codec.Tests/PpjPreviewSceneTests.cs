using System.Collections;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using OfficeKit.Artifact.Wire.V1;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class PpjPreviewSceneTests
{
    [Fact]
    public void CollectorSnapshotsEachWriterSlideOnceAndRequiresCompleteOrderedCapture()
    {
        var fixture = SceneFixture();
        var headers = fixture.Clone();
        headers.Slides[0].Elements.Clear();
        var collector = new PpjPreviewSceneBuilder.Collector(headers, EffectiveCodecLimits.From(null));
        Assert.Equal("invalid_preview_scene", Assert.Throws<CodecException>(() => collector.Complete(
            new string('a', 64), new string('b', 64), [], [])).Code);
        Assert.Throws<CodecException>(() => collector.CaptureSlide(1, fixture.Slides[0]));
        collector.CaptureSlide(0, fixture.Slides[0]);
        Assert.Throws<CodecException>(() => collector.CaptureSlide(0, fixture.Slides[0]));
        fixture.Slides[0].Elements[0].Group.Children[0].Shape.Text = "writer graph mutated after capture";
        var scene = collector.Complete(new string('a', 64), new string('b', 64), [], []);
        Assert.Equal("visual text", scene.Presentation.Slides[0].Elements[0].Group.Children[0].Shape.Text);
        Assert.Empty(headers.Slides[0].Elements);
        Assert.Throws<CodecException>(() => collector.Complete(new string('a', 64), new string('b', 64), [], []));
        Assert.Equal(Build(SceneFixture()), scene);
    }

    [Fact]
    public void EveryCurrentNativeContentCaseRetainsItsVisualFields()
    {
        var presentation = new PresentationArtifact { Slides = { new PresentationSlide { Id = "page" } } };
        foreach (var field in PresentationElement.Descriptor.Oneofs.Single(oneof => oneof.Name == "content").Fields)
        {
            var element = new PresentationElement { Id = field.Name };
            field.Accessor.SetValue(element, PopulateVisualMessage(field.MessageType, 0));
            presentation.Slides[0].Elements.Add(element);
        }
        var scene = Build(presentation);
        // Exact protobuf equality includes every populated scalar, optional
        // presence, collection order and nested field, including unpainted IR.
        Assert.Equal(presentation, scene.Presentation);
        Assert.Equal(9, scene.Presentation.Slides[0].Elements.Count);
    }

    [Fact]
    public void SceneBindingsKeepGeneratedOwnerPathsWithoutGrantingAuthority()
    {
        var binding = new PresentationPreviewNodeBinding
        {
            PageId = "page", SemanticId = "chart", NativeId = "chart-generated-child",
            ProgramPath = "$.pages[0].elements[0]", ScenePath = "$.presentation.slides[0].elements[0].group.children[0]",
            SourceId = "definition-child", ComponentId = "component", InstanceId = "instance", RepeatKey = "0",
            Attribution = PresentationPreviewAttribution.Generated, ZOrder = 0,
        };
        var scene = PpjPreviewSceneBuilder.Build(SceneFixture(), PresentationPreviewSceneOrigin.AuthoredLowering,
            new string('a', 64), new string('b', 64), [binding], [], EffectiveCodecLimits.From(null));
        Assert.Equal(binding, Assert.Single(scene.Bindings));
        binding.NativeId = "changed";
        Assert.Equal("chart-generated-child", scene.Bindings[0].NativeId);
        Assert.Equal("invalid_preview_scene", Assert.Throws<CodecException>(() => PpjPreviewSceneBuilder.Build(
            SceneFixture(), PresentationPreviewSceneOrigin.AuthoredLowering, new string('a', 64), new string('b', 64),
            [binding, binding], [], EffectiveCodecLimits.From(null))).Code);
    }

    private static IMessage PopulateVisualMessage(MessageDescriptor descriptor, int depth)
    {
        var message = descriptor.Parser.ParseFrom(ByteString.Empty);
        if (depth >= 7) return message;
        foreach (var field in descriptor.Fields.InFieldNumberOrder())
        {
            // Exclude only nonvisual export authority/payload from this fixture.
            // Opaque descriptors and native owner-local visual leaves remain.
            if (field.Name is "raw_xml" or "replacement_asset_id" or "element_deletions") continue;
            if (field.FieldType == FieldType.Message &&
                (field.MessageType.Name.EndsWith("SourceBinding", StringComparison.Ordinal) ||
                 field.MessageType.Name.EndsWith("Capability", StringComparison.Ordinal))) continue;
            if (field.ContainingOneof is { IsSynthetic: false } oneof && field != oneof.Fields[0]) continue;
            object value = field.FieldType switch
            {
                FieldType.Message => PopulateVisualMessage(field.MessageType, depth + 1),
                FieldType.String => field.Name,
                FieldType.Bool => false,
                FieldType.Double => 0d,
                FieldType.Float => 0f,
                FieldType.Int64 or FieldType.SInt64 or FieldType.SFixed64 => 7L,
                FieldType.UInt64 or FieldType.Fixed64 => 7UL,
                FieldType.Int32 or FieldType.SInt32 or FieldType.SFixed32 => 7,
                FieldType.UInt32 or FieldType.Fixed32 => 7U,
                FieldType.Enum => field.EnumType.Values[0].Number,
                _ => throw new InvalidOperationException($"Unclassified visual fixture field: {field.FullName}"),
            };
            if (field.IsRepeated) ((IList)field.Accessor.GetValue(message)).Add(value);
            else field.Accessor.SetValue(message, value);
        }
        return message;
    }

    [Fact]
    public void SceneBuilderRetainsVisualFieldsButNotExportAuthorityOrOpaqueXml()
    {
        var source = SceneFixture();
        var original = source.ToByteArray();
        var scene = Build(source);
        Assert.Equal(original, source.ToByteArray());
        Assert.Null(scene.Presentation.Slides[0].Source);
        var group = scene.Presentation.Slides[0].Elements[0];
        Assert.Null(group.Source);
        Assert.False(group.Hidden);
        Assert.True(group.HasHidden);
        var children = group.Group.Children;
        Assert.Null(children[0].Source);
        Assert.Equal(source.Slides[0].Elements[0].Group.Children[0].Shape, children[0].Shape);
        Assert.True(children[0].Shape.Shadow.HasOpacityThousandthPercent);
        Assert.Equal(0U, children[0].Shape.Shadow.OpacityThousandthPercent);
        Assert.False(children[0].Shape.Shadow.HasDistanceEmu);
        Assert.Equal(string.Empty, children[1].Opaque.RawXml);
        Assert.Equal("opaque source text", children[1].Opaque.Text);
        Assert.Equal("oleObject", children[1].Opaque.NativeKind);
        Assert.Equal(123L, children[1].Opaque.WidthEmu);
        Assert.Equal("", children[1].Opaque.OleWorkbook.ReplacementAssetId);
        Assert.Equal(source.Slides[0].Elements[0].Group.Children[2].Chart, children[2].Chart);
        Assert.Equal(0, children[2].Chart.YAxis.Minimum);
        Assert.True(children[2].Chart.YAxis.HasMinimum);
        Assert.False(children[2].Chart.YAxis.HasMinorUnit);
        Assert.Equal(source.Slides[0].Elements[0].Group.Children[3].Connector, children[3].Connector);
        Assert.Equal("target-b", children[3].Connector.EndTargetId);
        Assert.True(PpjPreviewSceneBuilder.Serialize(scene).Length < 10000);
        source.Slides[0].Elements[0].Group.Children[0].Shape.Text = "mutated later";
        Assert.Equal("visual text", children[0].Shape.Text);
    }

    [Fact]
    public void SceneDigestIsDeterministicAndIncludesProvenance()
    {
        var first = Build(SceneFixture());
        var second = Build(SceneFixture());
        Assert.Equal(first.Sha256, second.Sha256);
        Assert.Equal(PpjPreviewSceneBuilder.Serialize(first), PpjPreviewSceneBuilder.Serialize(second));
        var claimedHash = first.Sha256;
        first.Sha256 = "";
        Assert.Equal(claimedHash, Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
            PpjPreviewSceneBuilder.Serialize(first))).ToLowerInvariant());
        var changed = Build(SceneFixture(), candidateHash: new string('c', 64));
        Assert.NotEqual(claimedHash, changed.Sha256);
    }

    [Fact]
    public void SceneAssetReferencesValidateBytesAndHaveStableOrder()
    {
        var bytes = ByteString.CopyFromUtf8("asset bytes");
        var sha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes.Span)).ToLowerInvariant();
        var a = new Asset { Id = "native-a", ContentType = "image/png", Data = bytes, Sha256 = sha };
        var b = a.Clone(); b.Id = "native-b";
        var first = Build(SceneFixture(), assets: [b, a]);
        var second = Build(SceneFixture(), assets: [a, b]);
        Assert.Equal(first, second);
        Assert.Equal(new[] { "native-a", "native-b" }, first.Assets.Select(asset => asset.NativeId));
        a.Sha256 = sha.ToUpperInvariant();
        Assert.Equal(first, Build(SceneFixture(), assets: [a, b]));
        Assert.Equal(sha.ToUpperInvariant(), a.Sha256);
        b.Data = ByteString.CopyFromUtf8("changed");
        Assert.Equal("preview_scene_asset_mismatch", Assert.Throws<CodecException>(() => Build(SceneFixture(), assets: [b])).Code);
        Assert.Equal("invalid_preview_scene", Assert.Throws<CodecException>(() => Build(SceneFixture(), assets: [a, a])).Code);
    }

    [Fact]
    public void SceneBudgetsFailInsteadOfReturningTruncatedEvidence()
    {
        var fixture = SceneFixture();
        Assert.Equal("preview_scene_budget_exceeded", Assert.Throws<CodecException>(() => Build(fixture,
            limits: EffectiveCodecLimits.From(new CodecLimits { MaxUncompressedBytes = 128 }))).Code);
        var large = new PresentationArtifact { Slides = { new PresentationSlide { Id = "many" } } };
        for (var i = 0; i <= PpjProgramValidator.MaxExpandedElements; i++)
            large.Slides[0].Elements.Add(new PresentationElement { Id = i.ToString() });
        Assert.Equal("preview_scene_budget_exceeded", Assert.Throws<CodecException>(() => Build(large)).Code);
        var deep = new PresentationArtifact { Slides = { new PresentationSlide { Id = "deep" } } };
        var node = new PresentationElement { Group = new PresentationGroup() };
        deep.Slides[0].Elements.Add(node);
        for (var i = 0; i <= PpjPreviewSceneBuilder.MaxDepth; i++)
        {
            var child = new PresentationElement { Group = new PresentationGroup() };
            node.Group.Children.Add(child); node = child;
        }
        Assert.Equal("preview_scene_budget_exceeded", Assert.Throws<CodecException>(() => Build(deep)).Code);
    }

    private static PresentationPreviewScene Build(PresentationArtifact presentation,
        string? candidateHash = null, Asset[]? assets = null, EffectiveCodecLimits? limits = null) =>
        PpjPreviewSceneBuilder.Build(presentation, PresentationPreviewSceneOrigin.AuthoredLowering,
            new string('a', 64), candidateHash ?? new string('b', 64), [], assets ?? [],
            limits ?? EffectiveCodecLimits.From(null));

    private static PresentationArtifact SceneFixture() => new()
    {
        Id = "deck", SlideWidthEmu = 12000000, SlideHeightEmu = 6750000,
        Slides = { new PresentationSlide
        {
            Id = "page", Source = new PresentationSlideSourceBinding(),
            Elements = { new PresentationElement
            {
                Id = "group", Hidden = false, Source = new PresentationElementSourceBinding { Editable = true },
                Group = new PresentationGroup { Children =
                {
                    new PresentationElement { Id = "shape", Source = new PresentationElementSourceBinding { Editable = true },
                        Shape = new PresentationShape { Text = "visual text", Geometry = "diamond", FillRgb = "445566",
                            LeftEmu = 1234, WidthEmu = 5678,
                            Shadow = new PresentationShadow { OpacityThousandthPercent = 0, RotateWithShape = false } } },
                    new PresentationElement { Id = "opaque", Opaque = new PresentationOpaqueElement {
                        RawXml = new string('x', 1024 * 1024), Text = "opaque source text", NativeKind = "oleObject", WidthEmu = 123,
                        OleWorkbook = new PresentationOleWorkbook { ReplacementAssetId = "not-preview-authority" } } },
                    new PresentationElement { Id = "chart", Chart = new PresentationChart {
                        YAxis = new SpreadsheetChartAxisArtifact { Minimum = 0, Maximum = 10, Reverse = false } } },
                    new PresentationElement { Id = "connector", Connector = new PresentationConnector {
                        StartTargetId = "target-a", EndTargetId = "target-b", EndArrow = "triangle", LineOpacityThousandthPercent = 0 } },
                } },
            } },
        } },
    };

    [Fact]
    public void PreviewWireIsOptInAndPreservesNativePresence()
    {
        Assert.False(PresentationProgramRequest.Parser.ParseFrom(ByteString.Empty).IncludePreviewScene);
        Assert.Null(PresentationProgramResult.Parser.ParseFrom(ByteString.Empty).PreviewScene);
        var request = new PresentationProgramRequest { IncludePreviewScene = true };
        Assert.True(PresentationProgramRequest.Parser.ParseFrom(request.ToByteArray()).IncludePreviewScene);

        var original = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFromUtf8("{ \"canonical\": true }"),
            ProgramSha256 = new string('a', 64),
            OutputSha256 = new string('b', 64),
            PreviewScene = new PresentationPreviewScene
            {
                Version = 1,
                Origin = PresentationPreviewSceneOrigin.AuthoredLowering,
                ProgramSha256 = new string('a', 64),
                CandidateSha256 = new string('b', 64),
                Presentation = new PresentationArtifact
                {
                    Id = "deck",
                    Slides =
                    {
                        new PresentationSlide
                        {
                            Id = "page", Hidden = false,
                            Elements =
                            {
                                new PresentationElement { Id = "explicit", Hidden = false, Locked = false },
                                new PresentationElement { Id = "absent" },
                            },
                        },
                    },
                },
                Bindings = { new PresentationPreviewNodeBinding
                {
                    PageId = "page", SemanticId = "explicit", NativeId = "explicit",
                    ProgramPath = "$.pages[0].elements[0]",
                    ScenePath = "$.presentation.slides[0].elements[0]",
                    Attribution = PresentationPreviewAttribution.Direct, ZOrder = 0,
                } },
                Assets = { new PresentationPreviewAssetReference
                {
                    NativeId = "native-image", ContentType = "image/png", Sha256 = new string('c', 64),
                } },
            },
        };
        var restored = PresentationProgramResult.Parser.ParseFrom(original.ToByteArray());
        Assert.Equal(original, restored);
        Assert.Equal(original.ProgramJson, restored.ProgramJson);
        var elements = restored.PreviewScene.Presentation.Slides[0].Elements;
        Assert.True(elements[0].HasHidden);
        Assert.False(elements[0].Hidden);
        Assert.True(elements[0].HasLocked);
        Assert.False(elements[1].HasHidden);
        Assert.False(elements[1].HasLocked);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void BothProfilesRejectInvalidSceneOperationsBeforeParsing(bool ppjOnly, bool projection)
    {
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Operation = projection ? CodecOperation.ProjectPptxToPpj : CodecOperation.CompilePpjToPptx,
            File = projection ? ByteString.CopyFromUtf8("not a zip") : ByteString.Empty,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFromUtf8("not json"),
                IncludePreviewScene = true,
                ValidationOnly = !projection,
            },
        }, ppjOnly);
        Assert.False(response.Ok);
        Assert.True(response.File.IsEmpty);
        Assert.Null(response.PresentationProgram);
        var diagnostic = Assert.Single(response.Diagnostics);
        Assert.Equal("invalid_preview_scene_request", diagnostic.Code);
        Assert.Equal("presentation_program.include_preview_scene", diagnostic.SourcePath);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ExistingBuildAndCheckRemainSceneFree(bool ppjOnly, bool validationOnly)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        var response = Invoke(new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Operation = CodecOperation.CompilePpjToPptx,
            PresentationProgram = new PresentationProgramRequest
            {
                ProgramJson = ByteString.CopyFrom(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples", "ppj", "minimum.ppj"))),
                ValidationOnly = validationOnly,
            },
        }, ppjOnly);
        Assert.True(response.Ok, string.Join("\n", response.Diagnostics.Select(d => $"{d.Code}: {d.Message}")));
        Assert.Null(response.PresentationProgram.PreviewScene);
        Assert.False(response.PresentationProgram.ProgramJson.IsEmpty);
        Assert.Equal(validationOnly, response.File.IsEmpty);
    }

    private static CodecResponse Invoke(CodecRequest request, bool ppjOnly)
    {
        var bytes = request.ToByteArray();
        return ppjOnly ? PpjCodecProtocol.InvokeResponse(ref bytes, null) : CodecProtocol.InvokeResponse(ref bytes);
    }
}
