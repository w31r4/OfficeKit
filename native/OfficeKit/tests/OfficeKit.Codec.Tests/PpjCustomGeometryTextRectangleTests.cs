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
