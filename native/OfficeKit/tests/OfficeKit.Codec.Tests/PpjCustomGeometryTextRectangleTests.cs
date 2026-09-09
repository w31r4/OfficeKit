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
