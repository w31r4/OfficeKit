using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.IO.Compression;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class PpjConnectorObjectAnchorTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(40)]
    public void DirectedAnchorsCompileAtTargetsAndFollowMovement(int shift)
    {
        var program = Program(
            Edge("edge", Anchor("from-box", "left"), Anchor("to-box", "right")),
            Box("from-box", 750, 300, 60, 40), Box("to-box", 450, 100 + shift, 60, 40));
        var original = program.ToJsonString();
        var result = Compile(program);
        var connector = Find(result, "edge").Connector;
        Coordinates(connector, 750, 320, 510, 120 + shift);
        Assert.Equal("triangle", connector.EndArrow);
        using var stream = new MemoryStream(result.File.ToByteArray());
        using var document = PresentationDocument.Open(stream, false);
        var native = Assert.Single(document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.ConnectionShape>());
        var nativeIds = document.PresentationPart.SlideParts.Single().Slide.Descendants<P.NonVisualDrawingProperties>()
            .ToDictionary(properties => properties.Id!.Value, properties => properties.Name!.Value!);
        Assert.True(PptxConnectorCodec.TryRead(native, nativeIds, out var reread));
        Coordinates(reread!, 750, 320, 510, 120 + shift);
        Assert.Equal("from-box", reread.StartFrameAnchor.TargetId);
        Assert.Equal("left", reread.StartFrameAnchor.Anchor);
        Assert.Equal("right", reread.EndFrameAnchor.Anchor);
        Assert.Empty(native.Descendants<A.StartConnection>());
        Assert.Empty(native.Descendants<A.EndConnection>());
        Assert.Equal(connector.EndArrow, reread!.EndArrow);
        var withoutScene = Compile(program, preview: false);
        Assert.Null(withoutScene.PresentationProgram.PreviewScene);
        Assert.Equal(result.File, withoutScene.File);
        Assert.Equal(original, program.ToJsonString());
    }

    [Theory]
    [InlineData("top", 0, false, false, 140, 100)]
    [InlineData("right", 0, false, false, 180, 120)]
    [InlineData("bottom", 0, false, false, 140, 140)]
    [InlineData("left", 0, false, false, 100, 120)]
    [InlineData("center", 0, false, false, 140, 120)]
    [InlineData("left", 90, false, false, 140, 80)]
    [InlineData("left", 90, true, false, 140, 160)]
    [InlineData("top", 0, false, true, 140, 140)]
    [InlineData("center", 37, true, true, 140, 120)]
    public void ExplicitAnchorUsesTargetLocalFrame(string anchor, double rotation, bool flipH, bool flipV, double x, double y)
    {
        var box = Box("target", 100, 100, 80, 40);
        box["frame"]!["rotation"] = rotation;
        box["frame"]!["flipH"] = flipH;
        box["frame"]!["flipV"] = flipV;
        var result = Compile(Program(box, Edge("edge", Anchor("target", anchor), Literal(700, 300))));
        Coordinates(Find(result, "edge").Connector, x, y, 700, 300);
    }

    [Fact]
    public void AutoAnchorsChooseShortestPairWithStableTiesAndNoDeclarationOrderDependency()
    {
        foreach (var reversed in new[] { false, true })
        {
            var a = Box("a", 100, 100, 80, 40);
            var b = Box("b", 300, 100, 80, 40);
            var program = Program(reversed ? b : a, reversed ? a : b,
                Edge("pair", Anchor("a", "auto"), Anchor("b", "auto")),
                Edge("tie", Anchor("a", "auto"), Anchor("a", "auto")),
                Edge("mixed", Anchor("a", "auto"), Literal(700, 120)));
            var result = Compile(program);
            Coordinates(Find(result, "pair").Connector, 180, 120, 300, 120);
            Coordinates(Find(result, "tie").Connector, 140, 100, 140, 100);
            Coordinates(Find(result, "mixed").Connector, 180, 120, 700, 120);
        }
    }

    [Fact]
    public void NestedGroupAnchorsConvertBetweenTargetAndConnectorChildSpaces()
    {
        var target = Box("target", 2, 2, 5, 4);
        var inner = Group("inner", Frame(30, 20, 40, 20), Frame(0, 0, 20, 10), target,
            Edge("inside", Anchor("target", "right"), Anchor("outside", "center")));
        var outer = Group("outer", Frame(100, 100, 400, 200), Frame(10, 10, 100, 50), inner);
        var program = Program(outer, Box("outside", 500, 300, 20, 20),
            Edge("cross", Anchor("target", "left"), Literal(500, 300)));
        var result = Compile(program);
        Coordinates(Find(result, "cross").Connector, 196, 172, 500, 300);
        Coordinates(Find(result, "inside").Connector, 7, 4, 41.25, 21.25);

        // Outer horizontal flip followed by 90-degree rotation around (300,200).
        // The original point (196,172) maps to (328,304).
        outer["frame"]!["flipH"] = true;
        outer["frame"]!["rotation"] = 90;
        result = Compile(program);
        Coordinates(Find(result, "cross").Connector, 328, 304, 500, 300);
        Coordinates(Find(result, "inside").Connector, 7, 4, 1.25, -18.75);
        var signedSource = PptxCodecTests.RemoveEmbeddedPpj(result.File.ToByteArray());
        var signedProjection = Projected(signedSource);
        Assert.Equal("center", ByName(signedProjection, "inside")["to"]!["anchor"]!.GetValue<string>());
        Assert.Equal(signedSource, Compile(signedProjection, source: signedSource).File.ToByteArray());
        // Moving the outside target recomputes the same local-space endpoint.
        program["pages"]![0]!["elements"]![1]!["frame"] = Frame(200, 200, 20, 20);
        result = Compile(program);
        Coordinates(Find(result, "cross").Connector, 328, 304, 500, 300);
        Coordinates(Find(result, "inside").Connector, 7, 4, 13.75, 18.75);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ComponentPlacementTransformsLiteralsAndKeepsGroupChildrenLocal(bool explicitChildFrame)
    {
        var children = new JsonArray(Box("target", 20, 20, 20, 10),
            Edge("bound", Anchor("target", "right"), Literal(70, 25)),
            Edge("literal", Literal(40, 25), Literal(70, 25)));
        var group = Group("group", Frame(10, 10, 80, 40), explicitChildFrame ? Frame(10, 10, 80, 40) : null,
            Box("nested-target", 20, 20, 20, 10),
            Edge("nested-bound", Anchor("nested-target", "right"), Literal(70, 25)));
        children.Add(group);
        var program = Program(new JsonObject
        {
            ["type"] = "component", ["id"] = "instance", ["component"] = "definition", ["frame"] = Frame(100, 100, 400, 100),
        });
        program["components"] = new JsonArray(new JsonObject
        {
            ["id"] = "definition", ["frame"] = Frame(0, 0, 100, 50), ["elements"] = children,
        });
        var result = Compile(program);
        Coordinates(FindSuffix(result, "::bound").Connector, 260, 150, 380, 150);
        Coordinates(FindSuffix(result, "::literal").Connector, 260, 150, 380, 150);
        Coordinates(FindSuffix(result, "::nested-bound").Connector, 40, 25, 70, 25);
        var nativeGroup = FindSuffix(result, "::group").Group;
        Assert.Equal(140 * 12_700L, nativeGroup.LeftEmu);
        Assert.Equal(120 * 12_700L, nativeGroup.TopEmu);
        Assert.Equal(10 * 12_700L, nativeGroup.ChildLeftEmu);
        Assert.Equal(20 * 12_700L, FindSuffix(result, "::nested-target").Shape.LeftEmu);
    }

    [Fact]
    public void UnsupportedOrMissingTargetsNeverUseTheConnectorFrame()
    {
        var literal = Edge("literal", Literal(10, 10), Literal(20, 20));
        var linked = Edge("linked", Anchor("literal", "left"), Literal(70, 25));
        var result = Compile(Program(literal, linked), success: false);
        Assert.Contains(result.Diagnostics, d => d.Code == "ppj.connector.endpoint" && d.Message.Contains("unsupported type connector"));
        linked["from"]!["element"] = "missing";
        result = Compile(Program(literal.DeepClone().AsObject(), linked.DeepClone().AsObject()), success: false);
        Assert.Contains(result.Diagnostics, d => d.Code == "ppj.connector.target");

        var model = (PpjConnectorElementModel)PpjProgramParser.ParseElement(JsonSerializer.SerializeToElement(linked));
        var error = Assert.Throws<CodecException>(() => PpjConnectorEndpointResolver.ResolveLiteral(model));
        Assert.Equal("ppj.connector.endpoint", error.Code);
    }

    [Theory]
    [InlineData("top")]
    [InlineData("right")]
    [InlineData("bottom")]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("auto")]
    public void FrameAnchorSurvivesFreshProjectionAndByteIdenticalNoop(string anchor)
    {
        var authored = Compile(Program(Box("target", 100, 100, 80, 40),
            Edge("edge", Anchor("target", anchor), Literal(700, 300))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Project(source);
        Assert.False(projected.PresentationProgram.RestoredEmbeddedProgram);
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        var elements = program["pages"]![0]!["elements"]!.AsArray();
        var target = elements.Single(e => e!["type"]!.GetValue<string>() == "shape")!;
        var edge = elements.Single(e => e!["type"]!.GetValue<string>() == "connector")!;
        Assert.Equal(target["id"]!.GetValue<string>(), edge["from"]!["element"]!.GetValue<string>());
        Assert.Equal(anchor, edge["from"]!["anchor"]!.GetValue<string>());
        Assert.Equal(700, edge["to"]!["x"]!.GetValue<double>());
        Assert.DoesNotContain(target["nativeRef"]!["capabilities"]!.AsArray(),
            capability => capability!["operation"]!.GetValue<string>() == "delete");
        var noop = Compile(program, source: source);
        Assert.Equal(source, noop.File.ToByteArray());
    }

    [Fact]
    public void MalformedFrameAnchorProjectsAsOpaqueAndNoopPreservesItsBytes()
    {
        var authored = Compile(Program(Box("target", 100, 100, 80, 40),
            Edge("edge", Anchor("target", "center"), Literal(700, 300))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var shape = Assert.Single(document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.ConnectionShape>());
            var start = shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start");
            start.SetAttribute(new OpenXmlAttribute("anchor", "", "future-anchor"));
        }
        var malformed = stream.ToArray();
        var projected = Project(malformed);
        var program = JsonNode.Parse(projected.PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
        Assert.Single(program["pages"]![0]!["elements"]!.AsArray(), e => e!["type"]!.GetValue<string>() == "opaque");
        Assert.Equal(malformed, Compile(program, source: malformed).File.ToByteArray());
    }

    [Fact]
    public void FrameAnchorMetadataRejectsMalformedOrConflictingTopologyAndProtectsTargets()
    {
        var element = new PresentationElement { Id = "edge", Name = "edge", Connector = new PresentationConnector
        {
            ConnectorType = "straight", StartXEmu = 12700, StartYEmu = 12700, EndXEmu = 25400, EndYEmu = 25400,
            LineRgb = "112233", LineWidthEmu = 12700,
            StartFrameAnchor = new PresentationConnectorFrameAnchor { TargetId = "target", Anchor = "center" },
        } };
        var ids = new Dictionary<string, uint> { ["target"] = 2, ["edge"] = 3 };
        var reverse = new Dictionary<uint, string> { [2] = "target", [3] = "edge" };
        var native = PptxConnectorCodec.Build(element, 3, ids);
        Assert.True(PptxConnectorCodec.TryRead(native, reverse, out var decoded));
        Assert.Equal(element.Connector.StartFrameAnchor, decoded.StartFrameAnchor);
        Assert.True(PptxConnectorAnchorCodec.References(native, 2));
        Assert.False(PptxConnectorAnchorCodec.References(native, 5));
        foreach (var mutation in new Action<P.ConnectionShape>[]
        {
            shape => shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start")
                .SetAttribute(new OpenXmlAttribute("anchor", "", "unknown")),
            shape => shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start")
                .SetAttribute(new OpenXmlAttribute("target", "", "99")),
            shape => shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start")
                .SetAttribute(new OpenXmlAttribute("target", "", "3")),
            shape => shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start")
                .SetAttribute(new OpenXmlAttribute("unexpected", "", "yes")),
            shape => { var start = shape.Descendants().Single(n => n.NamespaceUri == PptxConnectorAnchorCodec.Namespace && n.LocalName == "start"); start.Parent!.Append(start.CloneNode(true)); },
            shape => { var ext = shape.Descendants<A.Extension>().Single(); ext.Parent!.Append(ext.CloneNode(true)); },
            shape => shape.NonVisualConnectionShapeProperties!.NonVisualConnectorShapeDrawingProperties!
                .PrependChild(new A.StartConnection { Id = 2, Index = 0 }),
        })
        {
            var candidate = (P.ConnectionShape)native.CloneNode(true);
            mutation(candidate);
            Assert.False(PptxConnectorCodec.TryRead(candidate, reverse, out _));
        }
        element.Connector.StartTargetId = "target";
        Assert.Throws<CodecException>(() => PptxConnectorCodec.Build(element, 3, ids));
        element.Connector.StartTargetId = "";
        element.Connector.StartFrameAnchor = null;
        PptxConnectorCodec.Apply(native, element, ids);
        Assert.Empty(native.Descendants<A.Extension>());
        Assert.False(PptxConnectorAnchorCodec.References(native, 2));
    }

    [Theory]
    [InlineData("move-frame", 510, 160)]
    [InlineData("move-leaf", 510, 160)]
    [InlineData("width-leaf", 550, 120)]
    [InlineData("rotation-leaf", 450, 120)]
    [InlineData("flip-leaf", 510, 120)]
    public void SourceTargetFrameEditsUpdateAttachedConnector(string mode, double x, double y)
    {
        var target = Box("to-box", 450, 100, 60, 40);
        if (mode == "rotation-leaf") target["frame"]!["rotation"] = 90;
        if (mode == "flip-leaf") target["frame"]!["flipH"] = true;
        var authored = Compile(Program(Box("from-box", 750, 300, 60, 40), target,
            Edge("edge", Anchor("from-box", "left"), Anchor("to-box", "right"))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var request = Projected(source);
        var box = ByName(request, "to-box");
        switch (mode)
        {
            case "move-frame": box["frame"]!["y"] = 140; break;
            case "move-leaf": Leaf(box, "topEmu")["value"] = 140 * 12_700L; break;
            case "width-leaf": Leaf(box, "widthEmu")["value"] = 100 * 12_700L; break;
            case "rotation-leaf": Leaf(box, "rotationDegrees")["value"] = 180; break;
            case "flip-leaf": Leaf(box, "flipHorizontal")["value"] = false; break;
        }
        var result = Compile(request, source: source);
        Coordinates(Assert.Single(Walk(result.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Connector is not null).Connector,
            750, 320, x, y);
        var second = Projected(result.File.ToByteArray());
        var edge = ByName(second, "edge");
        Assert.Equal("right", edge["to"]!["anchor"]!.GetValue<string>());
        Assert.Equal(ByName(second, "to-box")["id"]!.GetValue<string>(), edge["to"]!["element"]!.GetValue<string>());
        var recompiled = Compile(second, source: result.File.ToByteArray());
        Assert.Equal(result.File, recompiled.File);
        AssertSlideOnly(source, result.File.ToByteArray());
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("anchor", 780, 320, 510, 120)]
    [InlineData("retarget", 480, 120, 510, 120)]
    [InlineData("detach-start", 840, 400, 510, 120)]
    [InlineData("detach-both", 840, 400, 90, 20)]
    [InlineData("attach", 750, 320, 90, 20)]
    public void SourceEndpointEditsPreserveTheOtherEndpointAndReproject(string mode, double sx, double sy, double ex, double ey)
    {
        var authored = Compile(Program(Box("from-box", 750, 300, 60, 40), Box("to-box", 450, 100, 60, 40),
            Edge("edge", mode == "attach" ? Literal(840, 400) : Anchor("from-box", "left"),
                mode == "attach" ? Literal(90, 20) : Anchor("to-box", "right"))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var request = Projected(source);
        var edge = ByName(request, "edge");
        Assert.Contains(edge["nativeRef"]!["capabilities"]!.AsArray(), c => c!["operation"]!.GetValue<string>() == "setConnectorEndpoints");
        switch (mode)
        {
            case "anchor": edge["from"]!["anchor"] = "center"; break;
            case "retarget": edge["from"] = Anchor(ByName(request, "to-box")["id"]!.GetValue<string>(), "center"); break;
            case "detach-start": edge["from"] = Literal(840, 400); break;
            case "detach-both": edge["from"] = Literal(840, 400); edge["to"] = Literal(90, 20); break;
            case "attach": edge["from"] = Anchor(ByName(request, "from-box")["id"]!.GetValue<string>(), "left"); break;
        }
        var result = Compile(request, source: source);
        var connector = Assert.Single(Walk(result.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Connector is not null).Connector;
        Coordinates(connector, sx, sy, ex, ey);
        Assert.Equal("triangle", connector.EndArrow);
        var projected = Projected(result.File.ToByteArray());
        var projectedEdge = ByName(projected, "edge");
        Assert.True(JsonNode.DeepEquals(edge["from"], projectedEdge["from"]));
        Assert.True(JsonNode.DeepEquals(edge["to"], projectedEdge["to"]));
        if (mode == "detach-both")
        {
            Assert.Null(connector.StartFrameAnchor);
            Assert.Null(connector.EndFrameAnchor);
        }
        AssertSlideOnly(source, result.File.ToByteArray());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SourceGroupFrameAndChildFrameEditsUpdateDescendantAnchors(bool leaf, bool childFrame)
    {
        var authored = Compile(Program(
            Group("group", Frame(100, 100, 400, 200), Frame(10, 10, 100, 50), Box("target", 20, 20, 20, 10)),
            Edge("edge", Literal(10, 300), Anchor("target", "right"))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var request = Projected(source);
        var group = ByName(request, "group");
        if (leaf) Leaf(group, childFrame ? "childLeftEmu" : "leftEmu")["value"] = (childFrame ? 20 : 140) * 12_700L;
        else group[childFrame ? "childFrame" : "frame"]!["x"] = childFrame ? 20 : 140;
        var result = Compile(request, source: source);
        Coordinates(Assert.Single(Walk(result.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Connector is not null).Connector,
            10, 300, childFrame ? 180 : 260, 160);
        var second = Projected(result.File.ToByteArray());
        Assert.Equal("right", ByName(second, "edge")["to"]!["anchor"]!.GetValue<string>());
        AssertSlideOnly(source, result.File.ToByteArray());
    }

    [Fact]
    public void SourceAutoRecomputesAfterTargetMovesAndConflictingFrameEditsFail()
    {
        var authored = Compile(Program(Box("from-box", 750, 300, 60, 40), Box("to-box", 450, 100, 60, 40),
            Edge("edge", Anchor("from-box", "auto"), Anchor("to-box", "auto"))));
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var request = Projected(source);
        ByName(request, "to-box")["frame"] = Frame(850, 300, 60, 40);
        var result = Compile(request, source: source);
        Coordinates(Assert.Single(Walk(result.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Connector is not null).Connector,
            810, 320, 850, 320);
        Assert.Equal("auto", ByName(Projected(result.File.ToByteArray()), "edge")["from"]!["anchor"]!.GetValue<string>());
        AssertSlideOnly(source, result.File.ToByteArray());

        request = Projected(source);
        var target = ByName(request, "to-box");
        target["frame"]!["y"] = 150;
        Leaf(target, "topEmu")["value"] = 140 * 12_700L;
        var rejected = Compile(request, source: source, success: false);
        Assert.Contains(rejected.Diagnostics, d => d.Message.Contains("conflicting semantic and native frame edits"));
        Assert.Empty(rejected.File);
    }

    [Fact]
    public void NativeGeometrySiteAttachmentsKeepTheirSeparateSourceAuthority()
    {
        var authored = Compile(Program(Box("target", 100, 100, 80, 40), Edge("edge", Literal(140, 100), Literal(700, 300))));
        using var stream = new MemoryStream();
        stream.Write(PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray()));
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            var targetId = slide.Descendants<P.NonVisualDrawingProperties>().Single(p => p.Name?.Value == "target").Id!.Value;
            slide.Descendants<P.ConnectionShape>().Single().NonVisualConnectionShapeProperties!.NonVisualConnectorShapeDrawingProperties!
                .Append(new A.StartConnection { Id = targetId, Index = 0 });
        }
        var source = stream.ToArray();
        var request = Projected(source);
        var edge = ByName(request, "edge");
        Assert.DoesNotContain(edge["nativeRef"]!["capabilities"]!.AsArray(), c => c!["operation"]!.GetValue<string>() == "setConnectorEndpoints");
        Assert.Equal(source, Compile(request, source: source).File.ToByteArray());
        edge["from"]!["anchor"] = "left";
        var rejected = Compile(request, source: source, success: false);
        Assert.Contains(rejected.Diagnostics, d => d.Code == "ppj.nativeRef.capabilityMissing");
    }

    [Fact]
    public void SignedLiteralEndpointsReprojectAndEditWithoutChangingOtherParts()
    {
        var authored = Compile(Program(Edge("signed", Literal(-20, -10), Literal(40, 30))));
        Coordinates(Find(authored, "signed").Connector, -20, -10, 40, 30);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        using (var document = PresentationDocument.Open(new MemoryStream(source), false))
        {
            var transform = document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.ConnectionShape>().Single()
                .ShapeProperties!.GetFirstChild<A.Transform2D>()!;
            Assert.Equal(-20 * 12_700L, transform.Offset!.X!.Value);
            Assert.Equal(-10 * 12_700L, transform.Offset.Y!.Value);
        }
        var request = Projected(source);
        Assert.Equal(source, Compile(request, source: source).File.ToByteArray());
        ByName(request, "signed")["from"]!["x"] = -40;
        var edited = Compile(request, source: source);
        var fresh = ByName(Projected(edited.File.ToByteArray()), "signed");
        Assert.Equal(-40, fresh["from"]!["x"]!.GetValue<double>());
        Assert.Equal(-10, fresh["from"]!["y"]!.GetValue<double>());
        Assert.Equal(40, fresh["to"]!["x"]!.GetValue<double>());
        AssertSlideOnly(source, edited.File.ToByteArray());

        request = Projected(source);
        ByName(request, "signed")["frame"]!["x"] = -40;
        edited = Compile(request, source: source);
        fresh = ByName(Projected(edited.File.ToByteArray()), "signed");
        Assert.Equal(-40, fresh["from"]!["x"]!.GetValue<double>());
        Assert.Equal(20, fresh["to"]!["x"]!.GetValue<double>());
        AssertSlideOnly(source, edited.File.ToByteArray());

    }

    [Fact]
    public void CoordinateBoundsRejectOverflowAndPreserveMalformedNativeSource()
    {
        var rejected = Compile(Program(Edge("bad", Literal(-3e9, 0), Literal(40, 30))), success: false);
        Assert.Contains(rejected.Diagnostics, d => d.Code == "ppj.connector.endpoint");
        Assert.Empty(rejected.File);
        var oversized = new PresentationConnector { ConnectorType = "straight", StartXEmu = PptxConnectorCodec.MinimumCoordinate,
            EndXEmu = PptxConnectorCodec.MaximumCoordinate };
        Assert.Throws<CodecException>(() => PptxConnectorCodec.Validate(oversized, "bad", ""));
        oversized.StartXEmu = long.MinValue; oversized.EndXEmu = long.MaxValue;
        Assert.Throws<CodecException>(() => PptxConnectorCodec.Validate(oversized, "bad", ""));
        Assert.True(PptxConnectorCodec.IsEndpointPair(PptxConnectorCodec.MinimumCoordinate, 0, PptxConnectorCodec.MinimumCoordinate, 1));
        Assert.True(PptxConnectorCodec.IsEndpointPair(PptxConnectorCodec.MaximumCoordinate, 0, PptxConnectorCodec.MaximumCoordinate, 1));

        var valid = Compile(Program(Edge("bad", Literal(10, 10), Literal(40, 30))));
        using var stream = new MemoryStream();
        stream.Write(PptxCodecTests.RemoveEmbeddedPpj(valid.File.ToByteArray())); stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
            document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Transform2D>().Single().Offset!.X = long.MaxValue;
        var source = stream.ToArray();
        var limits = EffectiveCodecLimits.From(null);
        var imported = PptxCodec.Import(source, limits).Artifact;
        Assert.NotNull(Assert.Single(imported.Presentation.Slides[0].Elements).Opaque);
        Assert.Equal(source, PptxCodec.Export(imported, limits).File);
        var projection = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion, Operation = CodecOperation.ProjectPptxToPpj, Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source), PresentationProgram = new PresentationProgramRequest { SourceUri = "malformed.pptx" },
        }.ToByteArray();
        var response = PpjCodecProtocol.InvokeResponse(ref projection, null);
        Assert.False(response.Ok);
        Assert.Contains(response.Diagnostics, d => d.Code == "ppj.schema.maximum");
    }

    private static JsonObject Projected(byte[] source) => JsonNode.Parse(Project(source).PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    private static JsonObject ByName(JsonObject program, string name)
    {
        IEnumerable<JsonObject> WalkJson(JsonArray elements)
        {
            foreach (var element in elements.OfType<JsonObject>())
            {
                yield return element;
                if (element["elements"] is JsonArray children) foreach (var child in WalkJson(children)) yield return child;
            }
        }
        return Assert.Single(WalkJson(program["pages"]![0]!["elements"]!.AsArray()), e => e["name"]?.GetValue<string>() == name);
    }
    private static JsonObject Leaf(JsonObject owner, string kind) =>
        Assert.Single(owner["nativeRef"]!["leaves"]!.AsArray().OfType<JsonObject>(), leaf => leaf["kind"]!.GetValue<string>() == kind);
    private static void AssertSlideOnly(byte[] source, byte[] candidate)
    {
        Dictionary<string, byte[]> Parts(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
            return zip.Entries.ToDictionary(entry => entry.FullName, entry =>
            {
                using var input = entry.Open();
                using var output = new MemoryStream();
                input.CopyTo(output);
                return output.ToArray();
            });
        }
        var before = Parts(source);
        var after = Parts(candidate);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        Assert.Equal(new[] { "ppt/slides/slide1.xml" }, before.Keys.Where(key => !before[key].SequenceEqual(after[key])).Order());
    }

    private static JsonObject Frame(double x, double y, double width, double height) =>
        new() { ["x"] = x, ["y"] = y, ["width"] = width, ["height"] = height };
    private static JsonObject Anchor(string id, string anchor) => new() { ["element"] = id, ["anchor"] = anchor };
    private static JsonObject Literal(double x, double y) => new() { ["x"] = x, ["y"] = y };
    private static JsonObject Box(string id, double x, double y, double width, double height) => new()
    {
        ["id"] = id, ["type"] = "shape", ["frame"] = Frame(x, y, width, height),
        ["geometry"] = new JsonObject { ["kind"] = "preset", ["preset"] = "rect" },
        ["style"] = new JsonObject { ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = "#224466" } },
    };
    private static JsonObject Edge(string id, JsonObject from, JsonObject to) => new()
    {
        ["id"] = id, ["type"] = "connector", ["connectorType"] = "straight", ["frame"] = Frame(0, 0, 1, 1),
        ["from"] = from, ["to"] = to, ["stroke"] = new JsonObject { ["color"] = "#114477", ["width"] = 2 },
        ["endArrow"] = "triangle",
    };
    private static JsonObject Group(string id, JsonObject frame, JsonObject? childFrame, params JsonObject[] children)
    {
        var group = new JsonObject { ["id"] = id, ["type"] = "group", ["frame"] = frame, ["elements"] = new JsonArray(children.Cast<JsonNode>().ToArray()) };
        if (childFrame is not null) group["childFrame"] = childFrame;
        return group;
    }
    private static JsonObject Program(params JsonObject[] elements)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        Assert.NotNull(directory);
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples/ppj/minimum.ppj")))!.AsObject();
        program["pages"]![0]!["elements"] = new JsonArray(elements.Cast<JsonNode>().ToArray());
        return program;
    }
    private static CodecResponse Compile(JsonObject program, bool preview = true, bool success = true, byte[]? source = null)
    {
        var request = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion, Operation = CodecOperation.CompilePpjToPptx, Family = ArtifactFamily.Presentation,
            File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
            PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()), IncludePreviewScene = preview },
        };
        var bytes = request.ToByteArray();
        var result = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(result.Ok == success, string.Join("\n", result.Diagnostics.Select(d => $"{d.Code}: {d.Message}")));
        return result;
    }
    private static CodecResponse Project(byte[] source)
    {
        var request = new CodecRequest
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion, Operation = CodecOperation.ProjectPptxToPpj, Family = ArtifactFamily.Presentation,
            File = ByteString.CopyFrom(source), PresentationProgram = new PresentationProgramRequest { SourceUri = "connector/source.pptx" },
        };
        var bytes = request.ToByteArray();
        var result = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(result.Ok, string.Join("\n", result.Diagnostics.Select(d => $"{d.Code}: {d.Message}")));
        return result;
    }
    private static IEnumerable<PresentationElement> Walk(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.Group is { } group) foreach (var child in Walk(group.Children)) yield return child;
        }
    }
    private static PresentationElement Find(CodecResponse response, string id) =>
        Assert.Single(Walk(response.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Id == id);
    private static PresentationElement FindSuffix(CodecResponse response, string suffix) =>
        Assert.Single(Walk(response.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements), e => e.Id.EndsWith(suffix, StringComparison.Ordinal));
    private static void Coordinates(PresentationConnector connector, double sx, double sy, double ex, double ey)
    {
        Assert.Equal((long)Math.Round(sx * 12_700), connector.StartXEmu);
        Assert.Equal((long)Math.Round(sy * 12_700), connector.StartYEmu);
        Assert.Equal((long)Math.Round(ex * 12_700), connector.EndXEmu);
        Assert.Equal((long)Math.Round(ey * 12_700), connector.EndYEmu);
    }
}
