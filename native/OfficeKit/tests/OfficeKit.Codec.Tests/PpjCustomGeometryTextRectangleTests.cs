using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Nodes;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed class PpjCustomGeometryTextRectangleTests
{
    [Fact]
    public void TextRectangleAuthorsAndEditsWithReferenceIdentityAndOptionalPresence()
    {
        var program = Program();
        var shape = Element(program);
        var authored = Compile(program);
        var native = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.TextRectangle;
        Assert.Equal(10 * 12700L, native.LeftEmu);
        Assert.Equal(5 * 12700L, native.TopEmu);
        Assert.Equal("r", native.RightReference);
        Assert.Equal("b", native.BottomReference);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(shape["geometry"]!["textRectangle"], Element(projected)["geometry"]!["textRectangle"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var mode in new[] { "numeric", "reference", "remove" })
        {
            var request = Project(source);
            var element = Element(request);
            var geometry = element["geometry"]!.AsObject();
            if (mode == "remove") geometry.Remove("textRectangle");
            else if (mode == "numeric") { geometry["textRectangle"]!["right"] = 180; geometry["textRectangle"]!["bottom"] = 90; }
            else { geometry["textRectangle"]!["left"] = "l"; geometry["textRectangle"]!["top"] = "t"; }
            var result = Compile(request, source);
            var fresh = Element(Project(result.File.ToByteArray()));
            Assert.True(JsonNode.DeepEquals(geometry["textRectangle"], fresh["geometry"]!["textRectangle"]));
            foreach (var retained in new[] { "frame", "text" }) Assert.True(JsonNode.DeepEquals(element[retained], fresh[retained]), retained);
            Assert.True(JsonNode.DeepEquals(geometry["paths"], fresh["geometry"]!["paths"]));
            Assert.Equal(PathXml(source), PathXml(result.File.ToByteArray()));
            AssertSlideOnly(source, result.File.ToByteArray());
            if (mode == "remove")
            {
                using var document = PresentationDocument.Open(new MemoryStream(result.File.ToByteArray()), false);
                Assert.Empty(document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Rectangle>());
                var add = Project(result.File.ToByteArray());
                Element(add)["geometry"]!["textRectangle"] = new JsonObject { ["left"] = "l", ["top"] = "t", ["right"] = "r", ["bottom"] = "b" };
                var added = Compile(add, result.File.ToByteArray());
                Assert.Equal("l", Element(Project(added.File.ToByteArray()))["geometry"]!["textRectangle"]!["left"]!.GetValue<string>());
                AssertSlideOnly(result.File.ToByteArray(), added.File.ToByteArray());
            }
        }
        var invalid = Project(source);
        Element(invalid)["geometry"]!["textRectangle"]!["left"] = "unknownGuide";
        Assert.Empty(Compile(invalid, source, false).File);
        Element(invalid)["geometry"]!["textRectangle"]!["left"] = 250;
        Assert.Empty(Compile(invalid, source, false).File);
        var denied = Project(source);
        var fields = Element(denied)["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setGeometry")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "geometry.textRectangle"));
        Element(denied)["geometry"]!["textRectangle"]!["left"] = 20;
        Assert.Empty(Compile(denied, source, false).File);
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(),
            JsonSerializer.SerializeToElement(shape["geometry"]), "mask"));
        Assert.Equal(original, source);
    }

    [Fact]
    public void CustomGeometryGuidesRetainDependenciesAndSupportCoherentSourceEdits()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        geometry["guides"] = JsonNode.Parse("""
            [{"name":"inset","formula":"*/ w 1 10"},{"name":"rightBound","formula":"+- w 0 inset"}]
            """);
        geometry["textRectangle"]!["left"] = "inset";
        geometry["textRectangle"]!["right"] = "rightBound";
        var authored = Compile(program);
        var native = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
        Assert.Equal(2, native.CustomGuides.Count);
        Assert.True(PptxCustomGeometryFormulaCodec.Validate(native, "custom").TryResolveReference("inset", out var inset));
        Assert.Equal(20 * 12700d, inset);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry["guides"], Element(projected)["geometry"]!["guides"]));
        Assert.Equal("inset", Element(projected)["geometry"]!["textRectangle"]!["left"]!.GetValue<string>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var mode in new[] { "edit", "add", "remove" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!.AsObject();
            if (mode == "edit") requested["guides"]![0]!["formula"] = "*/ w 1 5";
            else if (mode == "add") requested["guides"]!.AsArray().Add(new JsonObject { ["name"] = "unused", ["formula"] = "val 0" });
            else { requested.Remove("guides"); requested.Remove("textRectangle"); }
            var candidate = Compile(request, source);
            var fresh = Element(Project(candidate.File.ToByteArray()));
            Assert.True(JsonNode.DeepEquals(requested["guides"], fresh["geometry"]!["guides"]));
            Assert.True(JsonNode.DeepEquals(requested["textRectangle"], fresh["geometry"]!["textRectangle"]));
            foreach (var field in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[field], fresh[field]));
            Assert.Equal(PathXml(source), PathXml(candidate.File.ToByteArray()));
            AssertSlideOnly(source, candidate.File.ToByteArray());
            if (mode == "edit")
            {
                var changed = candidate.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
                Assert.True(PptxCustomGeometryFormulaCodec.Validate(changed, "custom").TryResolveReference("inset", out var result));
                Assert.Equal(40 * 12700d, result);
            }
        }
        foreach (var invalid in new[] { "forward", "duplicate", "dangling" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!.AsObject();
            if (invalid == "forward") requested["guides"]![0]!["formula"] = "val rightBound";
            else if (invalid == "duplicate") requested["guides"]![1]!["name"] = "inset";
            else requested.Remove("guides");
            Assert.Empty(Compile(request, source, false).File);
        }
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(),
            JsonSerializer.SerializeToElement(geometry), "mask"));
        geometry.Remove("textRectangle"); geometry["guides"] = new JsonArray();
        var empty = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        Assert.Null(Element(Project(empty))["geometry"]!["guides"]);
        Assert.Equal(original, source);
    }

    [Fact]
    public void CustomGeometryAdjustmentsFeedGuidesAndRetainSourceLifecycle()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        geometry["adjustments"] = JsonNode.Parse("""
            [{"name":"padding","formula":"val 127000"},{"name":"inset","formula":"val padding"}]
            """);
        geometry["guides"] = JsonNode.Parse("""[{"name":"rightBound","formula":"+- w 0 inset"}]""");
        geometry["textRectangle"]!["left"] = "inset";
        geometry["textRectangle"]!["right"] = "rightBound";
        var authored = Compile(program);
        var native = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
        Assert.Equal(2, native.CustomAdjustments.Count);
        Assert.True(PptxCustomGeometryFormulaCodec.Validate(native, "custom").TryResolveReference("inset", out var inset));
        Assert.Equal(10 * 12700d, inset);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry["adjustments"], Element(projected)["geometry"]!["adjustments"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var mode in new[] { "edit", "add", "remove" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!.AsObject();
            if (mode == "edit") requested["adjustments"]![0]!["formula"] = "val 254000";
            else if (mode == "add") requested["adjustments"]!.AsArray().Add(new JsonObject { ["name"] = "unused", ["formula"] = "val 0" });
            else { requested.Remove("adjustments"); requested.Remove("guides"); requested.Remove("textRectangle"); }
            var candidate = Compile(request, source);
            var fresh = Element(Project(candidate.File.ToByteArray()));
            foreach (var field in new[] { "adjustments", "guides", "textRectangle" })
                Assert.True(JsonNode.DeepEquals(requested[field], fresh["geometry"]![field]), field);
            foreach (var field in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[field], fresh[field]));
            Assert.Equal(PathXml(source), PathXml(candidate.File.ToByteArray()));
            AssertSlideOnly(source, candidate.File.ToByteArray());
            if (mode == "edit")
            {
                var changed = candidate.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
                Assert.True(PptxCustomGeometryFormulaCodec.Validate(changed, "custom").TryResolveReference("inset", out var result));
                Assert.Equal(20 * 12700d, result);
            }
        }
        foreach (var invalid in new[] { "forward", "duplicate", "dangling" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!.AsObject();
            if (invalid == "forward") requested["adjustments"]![0]!["formula"] = "val rightBound";
            else if (invalid == "duplicate") requested["guides"]![0]!["name"] = "inset";
            else requested.Remove("adjustments");
            Assert.Empty(Compile(request, source, false).File);
        }
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(),
            JsonSerializer.SerializeToElement(geometry), "mask"));
        geometry.Remove("guides"); geometry.Remove("textRectangle"); geometry["adjustments"] = new JsonArray();
        var empty = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        Assert.Null(Element(Project(empty))["geometry"]!["adjustments"]);
        Element(program)["geometry"] = new JsonObject { ["kind"] = "preset", ["preset"] = "roundRect", ["adjustments"] = new JsonArray(10000) };
        var preset = Compile(program);
        Assert.Equal(10000, Assert.Single(preset.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.PresetAdjustments));
        Assert.Equal(10000, Element(Project(PptxCodecTests.RemoveEmbeddedPpj(preset.File.ToByteArray())))["geometry"]!["adjustments"]![0]!.GetValue<int>());
    }

    [Fact]
    public void CustomGeometryPathExtrusionPreservesPresenceAndSourceEdits()
    {
        foreach (bool? initial in new bool?[] { null, false, true })
        {
            var program = Program();
            var path = Element(program)["geometry"]!["paths"]![0]!.AsObject();
            path["fill"] = false; path["stroke"] = true;
            if (initial.HasValue) path["extrusionOk"] = initial.Value;
            var authored = Compile(program);
            var native = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.CustomPaths[0];
            Assert.Equal(initial.HasValue, native.HasExtrusionAllowed);
            if (initial.HasValue) Assert.Equal(initial.Value, native.ExtrusionAllowed);
            var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
            var original = source.ToArray();
            var projected = Project(source);
            Assert.True(JsonNode.DeepEquals(path, Element(projected)["geometry"]!["paths"]![0]));
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            foreach (var mode in new[] { "true", "false", "remove", "coordinate" })
            {
                var request = Project(source);
                var requested = Element(request)["geometry"]!["paths"]![0]!.AsObject();
                if (mode == "remove") requested.Remove("extrusionOk");
                else if (mode == "coordinate") requested["commands"]![1]!["x"] = 90;
                else requested["extrusionOk"] = mode == "true";
                var candidate = Compile(request, source).File.ToByteArray();
                var fresh = Element(Project(candidate));
                Assert.True(JsonNode.DeepEquals(Element(request)["geometry"], fresh["geometry"]));
                foreach (var field in new[] { "text", "frame" })
                    Assert.True(JsonNode.DeepEquals(Element(request)[field], fresh[field]));
                if (JsonNode.DeepEquals(Element(projected)["geometry"], Element(request)["geometry"]))
                    Assert.Equal(source, candidate);
                else AssertSlideOnly(source, candidate);
            }
            Assert.Equal(original, source);
        }
        var invalid = Program();
        Element(invalid)["geometry"]!["paths"]![0]!["extrusionOk"] = "false";
        Assert.Empty(Compile(invalid, null, false).File);
    }

    [Fact]
    public void CustomGeometryConnectionSitesPreserveReferencesAndSourceIdentity()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        geometry["adjustments"] = JsonNode.Parse("""[{"name":"inset","formula":"val 127000"}]""");
        geometry["guides"] = JsonNode.Parse("""[{"name":"siteX","formula":"val inset"}]""");
        geometry["connectionSites"] = JsonNode.Parse("""[{"angle":90,"x":10,"y":"vc"},{"angle":"cd2","x":"siteX","y":20}]""");
        var authored = Compile(program);
        var sites = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.CustomConnectionSites;
        Assert.Equal(2, sites.Count);
        Assert.Equal(5_400_000, sites[0].Angle60000);
        Assert.Equal(127_000, sites[0].XEmu);
        Assert.Equal("vc", sites[0].YReference);
        Assert.Equal("siteX", sites[1].XReference);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry, Element(projected)["geometry"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var field in new[] { "angle", "x", "y" })
        {
            var request = Project(source);
            request["pages"]![0]!["elements"]![0]!["geometry"]!["connectionSites"]![0]![field] = field == "angle" ? JsonValue.Create("cd4") : JsonValue.Create(30);
            var candidate = Compile(request, source).File.ToByteArray();
            var fresh = Element(Project(candidate));
            Assert.True(JsonNode.DeepEquals(Element(request)["geometry"], fresh["geometry"]));
            foreach (var retained in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[retained], fresh[retained]));
            Assert.Equal(PathXml(source), PathXml(candidate));
            AssertSlideOnly(source, candidate);
        }
        foreach (var mode in new[] { "remove", "add", "reference", "bounds", "angle", "denied" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!.AsObject();
            if (mode == "remove") requested.Remove("connectionSites");
            else if (mode == "add") requested["connectionSites"]!.AsArray().Add(requested["connectionSites"]![0]!.DeepClone());
            else if (mode == "reference") requested["connectionSites"]![0]!["x"] = "missing";
            else if (mode == "bounds") requested["connectionSites"]![0]!["x"] = 201;
            else if (mode == "angle") requested["connectionSites"]![0]!["angle"] = 361;
            else
            {
                requested["connectionSites"]![0]!["x"] = 30;
                foreach (var capability in Element(request)["nativeRef"]!["capabilities"]!.AsArray())
                    if (capability!["operation"]?.GetValue<string>() == "setGeometry")
                    {
                        var fields = capability["fields"]!.AsArray();
                        fields.Remove(fields.Single(f => f!.GetValue<string>() == "geometry.connectionSites"));
                    }
            }
            Assert.Empty(Compile(request, source, false).File);
        }
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(),
            JsonSerializer.SerializeToElement(geometry), "mask"));
        geometry["connectionSites"] = new JsonArray();
        Assert.Null(Element(Project(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray())))["geometry"]!["connectionSites"]);
        Assert.Equal(original, source);
    }

    [Fact]
    public void CustomGeometryAdjustmentHandlesRetainKindsRangesAndSourceIdentity()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        geometry["adjustments"] = JsonNode.Parse("""[{"name":"ax","formula":"val 127000"},{"name":"ay","formula":"val 254000"},{"name":"radius","formula":"val 127000"},{"name":"angle","formula":"val 5400000"}]""");
        geometry["adjustmentHandles"] = JsonNode.Parse("""
            [{"kind":"xy","xAdjustment":"ax","minX":0,"maxX":"w","yAdjustment":"ay","position":{"x":"ax","y":20}},
             {"kind":"polar","radialAdjustment":"radius","angleAdjustment":"angle","minAngle":0,"maxAngle":"cd2","position":{"x":10,"y":"ay"}}]
            """);
        var authored = Compile(program);
        var handles = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.CustomAdjustmentHandles;
        Assert.Equal(2, handles.Count);
        Assert.Equal("ax", handles[0].Xy.XAdjustment);
        Assert.True(handles[0].Xy.HasMinX);
        Assert.Equal(0, handles[0].Xy.MinX);
        Assert.False(handles[0].Xy.HasMinY);
        Assert.Equal(254_000, handles[0].Xy.Position.Y);
        Assert.Equal("cd2", handles[1].Polar.MaxAngleReference);
        Assert.Equal(127_000, handles[1].Polar.Position.X);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry, Element(projected)["geometry"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var mode in new[] { "add", "edit", "remove", "position" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!["adjustmentHandles"]!.AsArray();
            var xy = requested[0]!.AsObject(); var polar = requested[1]!.AsObject();
            if (mode == "add") { xy["minY"] = 0; xy["maxY"] = "h"; polar["minRadius"] = 0; polar["maxRadius"] = 100; }
            else if (mode == "edit") { xy["maxX"] = 150; polar["maxAngle"] = 180; }
            else if (mode == "remove") { xy.Remove("minX"); xy.Remove("maxX"); polar.Remove("minAngle"); polar.Remove("maxAngle"); }
            else { xy["position"]!["x"] = 30; xy["position"]!["y"] = "ay"; polar["position"]!["x"] = "ax"; polar["position"]!["y"] = 30; }
            var response = Compile(request, source);
            if (mode == "edit")
            {
                var changed = response.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.CustomAdjustmentHandles;
                Assert.Equal(150 * 12700, changed[0].Xy.MaxX);
                Assert.Equal(180 * 60000, changed[1].Polar.MaxAngle60000);
            }
            var candidate = response.File.ToByteArray();
            var fresh = Element(Project(candidate));
            Assert.True(JsonNode.DeepEquals(Element(request)["geometry"], fresh["geometry"]));
            foreach (var field in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[field], fresh[field]));
            Assert.Equal(PathXml(source), PathXml(candidate));
            AssertSlideOnly(source, candidate);
        }
        foreach (var mode in new[] { "remove", "reorder", "identity", "pair", "range", "reference", "position" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!["adjustmentHandles"]!.AsArray();
            if (mode == "remove") requested.RemoveAt(1);
            else if (mode == "reorder") { var first = requested[0]!.DeepClone(); var second = requested[1]!.DeepClone(); requested[0] = second; requested[1] = first; }
            else if (mode == "identity") requested[0]!["xAdjustment"] = "ay";
            else if (mode == "pair") requested[0]!.AsObject().Remove("maxX");
            else if (mode == "range") requested[1]!["maxAngle"] = 45;
            else if (mode == "reference") requested[0]!["position"]!["x"] = "missing";
            else requested[1]!["position"]!["y"] = 101;
            Assert.Empty(Compile(request, source, false).File);
        }
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(),
            JsonSerializer.SerializeToElement(geometry), "mask"));
        geometry["adjustmentHandles"] = new JsonArray();
        Assert.Null(Element(Project(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray())))["geometry"]!["adjustmentHandles"]);
        Assert.Equal(original, source);
    }

    [Fact]
    public void CustomGeometryPathReferencesRetainEverySlotAndSourceDependencies()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        geometry["adjustments"] = JsonNode.Parse("""[{"name":"ax","formula":"val 20000"}]""");
        geometry["guides"] = JsonNode.Parse("""[{"name":"edge","formula":"val ax"},{"name":"zero","formula":"val 0"}]""");
        geometry["paths"]![0]!["commands"] = JsonNode.Parse("""
            [{"op":"moveTo","x":"l","y":"t"},{"op":"lineTo","x":"edge","y":20},
             {"op":"quadraticTo","x1":"edge","y1":"edge","x":"edge","y":"edge"},
             {"op":"cubicTo","x1":"edge","y1":"edge","x2":"edge","y2":"edge","x":"edge","y":"edge"},
             {"op":"arcTo","radiusX":"edge","radiusY":"edge","startAngle":"zero","sweepAngle":"cd4"},{"op":"close"}]
            """);
        var authored = Compile(program);
        var native = authored.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
        Assert.Equal("edge", native.CustomPaths[0].Commands[2].QuadraticBezierTo.Control.XReference);
        Assert.True(PptxCustomGeometryFormulaCodec.Validate(native, "custom").TryResolveReference("edge", out var initial));
        Assert.Equal(20000d, initial);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry, Element(projected)["geometry"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        var commands = geometry["paths"]![0]!["commands"]!.AsArray();
        for (var index = 0; index < commands.Count; index++)
        foreach (var field in commands[index]!.AsObject().Where(p => p.Key != "op").Select(p => p.Key))
        {
            var request = Project(source);
            var target = Element(request)["geometry"]!["paths"]![0]!["commands"]![index]!;
            target[field] = index == 1 && field == "y" ? JsonValue.Create("edge") : JsonValue.Create(10);
            var candidate = Compile(request, source).File.ToByteArray();
            var fresh = Element(Project(candidate));
            Assert.True(JsonNode.DeepEquals(Element(request)["geometry"], fresh["geometry"]), $"{index}.{field}");
            foreach (var retained in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[retained], fresh[retained]));
            AssertSlideOnly(source, candidate);
        }
        var dependency = Project(source);
        Element(dependency)["geometry"]!["adjustments"]![0]!["formula"] = "val 30000";
        var changed = Compile(dependency, source);
        var changedShape = changed.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape;
        Assert.True(PptxCustomGeometryFormulaCodec.Validate(changedShape, "custom").TryResolveReference("edge", out var resolved));
        Assert.Equal(30000d, resolved);
        Assert.Equal(PathXml(source), PathXml(changed.File.ToByteArray()));
        AssertSlideOnly(source, changed.File.ToByteArray());
        Assert.True(JsonNode.DeepEquals(Element(dependency)["geometry"], Element(Project(changed.File.ToByteArray()))["geometry"]));
        foreach (var mode in new[] { "missing", "radius", "sweep" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!["paths"]![0]!["commands"]!;
            if (mode == "missing") requested[1]!["x"] = "unknown";
            else requested[4]![mode == "radius" ? "radiusX" : "sweepAngle"] = "zero";
            Assert.Empty(Compile(request, source, false).File);
        }
        var mask = geometry.DeepClone().AsObject(); mask.Remove("adjustments"); mask.Remove("guides"); mask.Remove("textRectangle");
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(), JsonSerializer.SerializeToElement(mask), "mask"));
        var lineProgram = Program();
        Element(lineProgram).Remove("geometry"); Element(lineProgram).Remove("text"); Element(lineProgram).Remove("fill");
        Element(lineProgram)["type"] = "line";
        Element(lineProgram)["stroke"] = JsonNode.Parse("""{"color":"#112233","width":1}""");
        Element(lineProgram)["path"] = JsonNode.Parse("""{"viewBox":{"x":0,"y":0,"width":100,"height":100},"commands":[{"op":"moveTo","x":0,"y":0},{"op":"lineTo","x":"w","y":10}]}""");
        Assert.Empty(Compile(lineProgram, null, false).File);
        Element(lineProgram)["path"]!["commands"]![1]!["x"] = 50;
        Assert.NotEmpty(Compile(lineProgram).File);
        Assert.Equal(original, source);
    }

    [Fact]
    public void CustomGeometryPathViewportsRetainIndependentAndDefaultAxes()
    {
        var program = Program();
        var geometry = Element(program)["geometry"]!.AsObject();
        var paths = geometry["paths"]!.AsArray();
        paths[0]!["fill"] = false; paths[0]!["stroke"] = true; paths[0]!["extrusionOk"] = false;
        foreach (var pair in new[] { (50, 200), (0, 25), (30, 0), (0, 0) })
        {
            var path = paths[0]!.DeepClone();
            path["viewport"] = new JsonObject { ["width"] = pair.Item1, ["height"] = pair.Item2 };
            paths.Add(path);
        }
        var authored = Compile(program);
        AssertPathExtents(program, authored);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(JsonNode.DeepEquals(geometry, Element(projected)["geometry"]));
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var mode in new[] { "edit", "remove", "firstDefault", "firstLarge" })
        {
            var request = Project(source);
            var requested = Element(request)["geometry"]!["paths"]!.AsArray();
            if (mode == "edit") requested[1]!["viewport"] = new JsonObject { ["width"] = 75, ["height"] = 0 };
            else if (mode == "remove") requested[1]!.AsObject().Remove("viewport");
            else requested[0]!["viewport"] = new JsonObject { ["width"] = mode == "firstDefault" ? 0 : 2147483.647, ["height"] = mode == "firstDefault" ? 0 : 2147483.647 };
            var candidate = Compile(request, source);
            AssertPathExtents(request, candidate);
            var fresh = Project(candidate.File.ToByteArray());
            if (mode == "firstLarge")
            {
                var otherKind = fresh.DeepClone().AsObject();
                var leaf = Element(otherKind)["nativeRef"]!["leaves"]!.AsArray()
                    .First(l => l!["kind"]!.GetValue<string>() == "customGeometryPathWidth")!;
                leaf["kind"] = "customGeometryPathArcWidthRadius";
                Assert.False(PpjProgramValidator.Validate(JsonSerializer.SerializeToUtf8Bytes(otherKind)).IsValid);
            }
            AssertPathExtents(request, Compile(fresh, candidate.File.ToByteArray()));
            var freshShape = Element(fresh);
            foreach (var field in new[] { "text", "frame" }) Assert.True(JsonNode.DeepEquals(Element(request)[field], freshShape[field]));
            Assert.True(JsonNode.DeepEquals(Element(request)["geometry"]!["textRectangle"], freshShape["geometry"]!["textRectangle"]));
            for (var index = 0; index < paths.Count; index++)
            {
                var expected = requested[index]!.DeepClone().AsObject(); expected.Remove("viewport");
                var actual = freshShape["geometry"]!["paths"]![index]!.DeepClone().AsObject(); actual.Remove("viewport");
                Assert.True(JsonNode.DeepEquals(expected, actual));
            }
            AssertSlideOnly(source, candidate.File.ToByteArray());
        }
        foreach (var invalid in new[] { -1d, 2147483.648 })
        {
            var request = Project(source);
            Element(request)["geometry"]!["paths"]![1]!["viewport"]!["width"] = invalid;
            Assert.Empty(Compile(request, source, false).File);
        }
        var mask = geometry.DeepClone().AsObject(); mask.Remove("textRectangle");
        Assert.Throws<CodecException>(() => PpjAuthoredPresentationCompiler.ApplyCustomGeometry(new PresentationShape(), JsonSerializer.SerializeToElement(mask), "mask"));
        Assert.Equal(original, source);
    }

    private static void AssertPathExtents(JsonObject request, CodecResponse response)
    {
        var geometry = Element(request)["geometry"]!;
        var paths = geometry["paths"]!.AsArray();
        var native = response.PresentationProgram.PreviewScene.Presentation.Slides[0].Elements[0].Shape.CustomPaths;
        Assert.Equal(paths.Count, native.Count);
        for (var index = 0; index < paths.Count; index++)
        {
            var extent = JsonSerializer.SerializeToElement(paths[index]!["viewport"] ?? geometry["viewBox"]!);
            Assert.Equal((long)Math.Round(extent.GetProperty("width").GetDouble() * 1000), native[index].Width);
            Assert.Equal((long)Math.Round(extent.GetProperty("height").GetDouble() * 1000), native[index].Height);
        }
    }

    private static JsonObject Program()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples/ppj/minimum.ppj")))!.AsObject();
        var shape = JsonNode.Parse("""
            {"id":"custom","type":"shape","frame":{"x":30,"y":40,"width":200,"height":100},
             "geometry":{"kind":"custom","viewBox":{"x":0,"y":0,"width":100,"height":100},
               "paths":[{"commands":[{"op":"moveTo","x":0,"y":0},{"op":"lineTo","x":100,"y":0},
                 {"op":"lineTo","x":100,"y":100},{"op":"close"}]}],
               "textRectangle":{"left":10,"top":5,"right":"r","bottom":"b"}},
             "text":"Rectangle text","style":{"fill":{"type":"solid","color":"#AACCEE"}}}
            """)!.AsObject();
        program["pages"]![0]!["elements"] = new JsonArray(shape);
        return program;
    }

    private static JsonObject Element(JsonObject program) => program["pages"]![0]!["elements"]![0]!.AsObject();
    private static CodecResponse Compile(JsonObject program, byte[]? source = null, bool success = true) => Invoke(new CodecRequest
    {
        Operation = CodecOperation.CompilePpjToPptx, File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()), IncludePreviewScene = true },
    }, success);
    private static JsonObject Project(byte[] source) => JsonNode.Parse(Invoke(new CodecRequest
    {
        Operation = CodecOperation.ProjectPptxToPpj, File = ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { SourceUri = "source.pptx", AssetRootUri = "assets" },
    }).PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    private static CodecResponse Invoke(CodecRequest request, bool success = true)
    {
        request.ProtocolVersion = CodecProtocol.ProtocolVersion; request.Family = ArtifactFamily.Presentation;
        var bytes = request.ToByteArray(); var result = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(result.Ok == success, string.Join("\n", result.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        return result;
    }
    private static string PathXml(byte[] bytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.PathList>().Single().OuterXml;
    }
    private static void AssertSlideOnly(byte[] source, byte[] candidate)
    {
        Dictionary<string, byte[]> Parts(byte[] bytes)
        {
            using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            return zip.Entries.ToDictionary(e => e.FullName, e => { using var input = e.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray(); });
        }
        var before = Parts(source); var after = Parts(candidate);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        Assert.Equal(new[] { "ppt/slides/slide1.xml" }, before.Keys.Where(k => !before[k].SequenceEqual(after[k])).Order());
    }
}
