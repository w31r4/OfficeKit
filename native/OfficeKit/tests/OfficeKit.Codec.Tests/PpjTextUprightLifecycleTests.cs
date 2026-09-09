using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed class PpjTextUprightLifecycleTests
{
    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void UprightOnlyStyleCanBeRemovedWithExistingAuthority(string kind)
    {
        var program = Program(kind);
        Style(program, kind).Clear(); Style(program, kind)["upright"] = true;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var request = Project(source);
        Assert.Equal(new[] { "upright" }, Style(request, kind).Select(p => p.Key));
        JsonObject StyleOwner(JsonObject input) => (kind == "shape" ? input["pages"]![0]!["elements"]![0] :
            input["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
        StyleOwner(request).Remove(kind == "shape" ? "textStyle" : "style");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).UpRight);
        var fresh = Project(deleted);
        if (kind == "shape") Assert.False(StyleOwner(fresh).ContainsKey("textStyle"));
        else
        {
            var cell = fresh["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!;
            Assert.Equal("Retain this text", cell["text"]!.GetValue<string>());
            cell["text"] = new JsonObject { ["style"] = new JsonObject { ["upright"] = false },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }) };
            var restored = Compile(fresh, deleted).File.ToByteArray();
            Assert.False(Body(restored, kind).UpRight!.Value);
            Assert.False(Style(Project(restored), kind)["upright"]!.GetValue<bool>());
            AssertOnlyUprightChanged(deleted, restored, kind);
        }
        AssertOnlyUprightChanged(source, deleted, kind);

        var denied = Project(source);
        StyleOwner(denied).Remove(kind == "shape" ? "textStyle" : "style");
        var capabilities = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray();
        capabilities.Remove(capabilities.Single(c => c!["operation"]!.GetValue<string>() == (kind == "shape" ? "setTextBodyStyle" : "setTableCellStyle")));
        Assert.Empty(Compile(denied, source, success: false).File);

        var otherSource = PptxCodecTests.RemoveEmbeddedPpj(Compile(Program(kind)).File.ToByteArray());
        var otherEdit = Project(otherSource);
        StyleOwner(otherEdit).Remove(kind == "shape" ? "textStyle" : "style");
        Assert.Empty(Compile(otherEdit, otherSource, success: false).File);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void UprightSourceRemovalAndRestorationPreserveOtherState(string kind)
    {
        var program = Program(kind);
        var authored = Compile(program);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(Style(projected, kind)["upright"]!.GetValue<bool>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        var baseline = Body(source, kind);
        Assert.True(baseline.UpRight!.Value);
        var request = Project(source);
        Style(request, kind).Remove("upright");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).UpRight);
        Assert.False(Style(Project(deleted), kind).ContainsKey("upright"));
        AssertOnlyUprightChanged(source, deleted, kind);
        foreach (var value in new[] { false, true })
        {
            var restore = Project(deleted);
            Style(restore, kind)["upright"] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(value, Body(restored, kind).UpRight!.Value);
            Assert.Equal(value, Style(Project(restored), kind)["upright"]!.GetValue<bool>());
            AssertOnlyUprightChanged(deleted, restored, kind);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove("upright");
            Assert.Null(Body(Compile(removeAgain, restored).File.ToByteArray(), kind).UpRight);
        }
        Assert.Equal(original, source);
    }

    private static JsonObject Program(string kind)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples/ppj/minimum.ppj")))!.AsObject();
        var element = JsonNode.Parse("""
            {"id":"owner","type":"shape","frame":{"x":30,"y":40,"width":200,"height":100},
             "geometry":{"kind":"preset","preset":"rect"},"text":"Retain this text",
             "textStyle":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}}}
            """)!.AsObject();
        if (kind == "table")
        {
            element = JsonNode.Parse("""
                {"id":"owner","type":"table","frame":{"x":30,"y":40,"width":200,"height":100},
                 "columns":[{"id":"column","width":200}],"rows":[{"id":"row","height":100,"cells":[{"id":"cell",
                 "text":{"style":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}},
                 "paragraphs":[{"runs":[{"text":"Retain this text"}]}]}}]}]}
                """)!.AsObject();
        }
        else if (kind == "text")
        {
            element["type"] = kind;
            element["style"] = element["textStyle"]!.DeepClone(); element.Remove("textStyle"); element.Remove("geometry");
        }
        program["pages"]![0]!["elements"] = new JsonArray(element);
        if (kind is "master" or "layout")
        {
            program["design"]!["masters"] = JsonNode.Parse("""[{"id":"master","name":"Master","placeholders":[]}]""");
            program["design"]!["layouts"] = JsonNode.Parse("""[{"id":"layout","name":"Layout","master":"master","layoutType":"blank","placeholders":[]}]""");
            var placeholder = new JsonObject { ["id"] = "placeholder", ["name"] = "Owner", ["placeholderType"] = "body", ["index"] = 0,
                ["frame"] = element["frame"]!.DeepClone(), ["text"] = element["text"]!.DeepClone(), ["style"] = element["textStyle"]!.DeepClone() };
            program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"] = new JsonArray(placeholder);
            program["pages"]![0]!["layout"] = "layout";
            program["pages"]![0]!["elements"] = new JsonArray();
        }
        return program;
    }

    private static JsonObject Style(JsonObject program, string kind)
    {
        if (kind is "master" or "layout") return program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"]![0]!["style"]!.AsObject();
        var element = program["pages"]![0]!["elements"]![0]!;
        return (kind == "table" ? element["rows"]![0]!["cells"]![0]!["text"]!["style"] : element[kind == "shape" ? "textStyle" : "style"])!.AsObject();
    }
    private static CodecResponse Compile(JsonObject program, byte[]? source = null, bool success = true) => Invoke(new CodecRequest
    {
        Operation = CodecOperation.CompilePpjToPptx, File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
    }, success);
    private static JsonObject Project(byte[] source) => JsonNode.Parse(Invoke(new CodecRequest
    {
        Operation = CodecOperation.ProjectPptxToPpj, File = ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { SourceUri = "source.pptx" },
    }).PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    private static CodecResponse Invoke(CodecRequest request, bool success = true)
    {
        request.ProtocolVersion = CodecProtocol.ProtocolVersion; request.Family = ArtifactFamily.Presentation;
        var bytes = request.ToByteArray(); var response = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(response.Ok == success, string.Join("\n", response.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        return response;
    }
    private static OpenXmlPartRootElement Owner(PresentationDocument document, string kind) => kind switch
    {
        "master" => document.PresentationPart!.SlideMasterParts.Single().SlideMaster,
        "layout" => document.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout,
        _ => document.PresentationPart!.SlideParts.Single().Slide,
    };
    private static A.BodyProperties Body(byte[] bytes, string kind)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return (A.BodyProperties)Assert.Single(Owner(document, kind).Descendants<A.BodyProperties>()).CloneNode(true);
    }
    private static void AssertOnlyUprightChanged(byte[] before, byte[] after, string kind)
    {
        using var left = PresentationDocument.Open(new MemoryStream(before), false);
        using var right = PresentationDocument.Open(new MemoryStream(after), false);
        var oldSlide = Owner(left, kind);
        var newSlide = Owner(right, kind);
        Assert.Single(oldSlide.Descendants<A.BodyProperties>()).UpRight = null;
        Assert.Single(newSlide.Descendants<A.BodyProperties>()).UpRight = null;
        XElement Content(string xml)
        {
            var node = XElement.Parse(xml);
            foreach (var attribute in node.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).ToArray()) attribute.Remove();
            return node;
        }
        // The table writer may repeat namespace declarations. Compare expanded
        // XML names/values while keeping all non-namespace attributes/content.
        Assert.True(XNode.DeepEquals(Content(oldSlide.OuterXml), Content(newSlide.OuterXml)));
        using var oldZip = new ZipArchive(new MemoryStream(before));
        using var newZip = new ZipArchive(new MemoryStream(after));
        Assert.Equal(oldZip.Entries.Select(e => e.FullName).Order(), newZip.Entries.Select(e => e.FullName).Order());
        var ownerPart = kind switch { "master" => "ppt/slideMasters/slideMaster1.xml", "layout" => "ppt/slideLayouts/slideLayout1.xml", _ => "ppt/slides/slide1.xml" };
        foreach (var entry in oldZip.Entries.Where(e => e.FullName != ownerPart))
        {
            using var a = entry.Open(); using var b = newZip.GetEntry(entry.FullName)!.Open();
            using var aa = new MemoryStream(); using var bb = new MemoryStream(); a.CopyTo(aa); b.CopyTo(bb);
            Assert.Equal(aa.ToArray(), bb.ToArray());
        }
    }
}
