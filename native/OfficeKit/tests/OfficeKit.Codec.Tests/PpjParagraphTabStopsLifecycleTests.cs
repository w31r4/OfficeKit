using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    private const string InitialTabs = """[{"position":30},{"position":90,"alignment":"decimal"}]""";
    private const string ChangedTabs = """[{"position":0},{"position":72.12345,"alignment":"center"},{"position":144,"alignment":"right"},{"position":190,"alignment":"decimal"}]""";

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphTabStopsLifecyclePreservesSource(string kind)
    {
        var program = BulletFontProgram(kind); SetTabs(program, InitialTabs);
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["tabStops"] = JsonNode.Parse(InitialTabs);
        using var stream = new MemoryStream(); stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
        using (var document = PresentationDocument.Open(stream, true))
        {
            foreach (var paragraph in Owner(document, kind).Descendants<A.Paragraph>())
            {
                var properties = paragraph.ParagraphProperties!;
                var tab = properties.GetFirstChild<A.TabStopList>()!.GetFirstChild<A.TabStop>()!;
                tab.SetAttribute(new OpenXmlAttribute("pos", "", "+00381000")); tab.Alignment = null;
                properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
            }
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var choice in new string?[] { ChangedTabs, "[]", InitialTabs, null, ChangedTabs, "clear", InitialTabs })
        {
            var request = Project(before); SetTabs(request, choice); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertTabs(candidate, kind, choice);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.tabStops");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        foreach (var choice in new string?[] { null, "[]", "clear" })
        {
            var request = Project(source); SetTabs(request, choice);
            var candidate = Compile(request, source).File.ToByteArray();
            AssertTabs(candidate, kind, null);
            AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraph.tabStops");
        }
        foreach (var field in new[] { "tabStops", "noTabStops" })
        {
            var request = Project(source); SetTabs(request, "clear");
            var fields = request["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
            fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style." + field));
            Assert.Empty(Compile(request, source, success: false).File);
        }
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{},"runs":[{"text":"Tab only"}]}]}""");
        SetTabs(minimal, InitialTabs); var onlySource = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        var removeStyle = Project(onlySource); FirstTextParagraph(removeStyle).Remove("style");
        AssertTabs(Compile(removeStyle, onlySource).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphTabStopsClearsEmptyNativeList()
    {
        var program = BulletFontProgram("text"); SetTabs(program, InitialTabs);
        using var stream = new MemoryStream();
        stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
        using (var document = PresentationDocument.Open(stream, true))
            Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.TabStopList>()!.RemoveAllChildren();
        var source = stream.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), "text", "paragraphDefault.bold");
        foreach (var choice in new[] { "[]", "clear" })
        {
            var request = Project(source); SetTabs(request, choice);
            var candidate = Compile(request, source).File.ToByteArray();
            AssertTabs(candidate, "text", null);
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraph.tabStops");
            // Clearing an absent list is also valid.
            var repeat = Project(candidate); SetTabs(repeat, choice);
            AssertTabs(Compile(repeat, candidate).File.ToByteArray(), "text", null);
        }
    }

    [Fact]
    public void ParagraphTabStopsGroupClearRunsNativeValidation()
    {
        var program = BulletFontProgram("text"); SetTabs(program, InitialTabs);
        var elements = program["pages"]![0]!["elements"]!.AsArray();
        var child = elements[0]!.DeepClone();
        elements[0] = new JsonObject { ["id"] = "tab-group", ["type"] = "group",
            ["frame"] = child["frame"]!.DeepClone(), ["elements"] = new JsonArray(child) };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var malformed in new[] { false, true })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var list = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.TabStopList>()!;
                if (malformed) list.Elements<A.TabStop>().First().Position = null;
                else list.RemoveAllChildren();
            }
            var source = stream.ToArray(); var request = Project(source);
            Assert.Equal(source, Compile(request, source).File.ToByteArray());
            var paragraph = request["pages"]![0]!["elements"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!;
            paragraph["style"]!["noTabStops"] = true;
            var result = Compile(request, source, success: !malformed);
            if (malformed) { Assert.Empty(result.File); continue; }
            var candidate = result.File.ToByteArray();
            using var output = PresentationDocument.Open(new MemoryStream(candidate), false);
            Assert.Null(Owner(output, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.TabStopList>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraph.tabStops");
            var projected = Project(candidate);
            var style = projected["pages"]![0]!["elements"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!["style"];
            Assert.Null(style?["tabStops"]); Assert.Null(style?["noTabStops"]);
            Assert.Equal(candidate, Compile(projected, candidate).File.ToByteArray());
        }
    }

    [Fact]
    public void ParagraphTabStopsBoundsAndPrecision()
    {
        var program = BulletFontProgram("text");
        foreach (var choice in new[] { "[]", InitialTabs, ChangedTabs,
            new JsonArray(new JsonObject { ["position"] = int.MaxValue / 12700d }).ToJsonString(),
            new JsonArray(Enumerable.Range(0, 32).Select(i => (JsonNode)new JsonObject { ["position"] = i }).ToArray()).ToJsonString() })
        {
            SetTabs(program, choice); var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertTabs(bytes, "text", choice);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        foreach (var invalid in new[] { "null", "false", "{}", """[{"position":-0.00001}]""", """[{"position":169094}]""",
            """[{"position":"1"}]""", """[{}]""", """[{"position":1,"alignment":"invalid"}]""",
            """[{"position":2},{"position":1}]""", """[{"position":1},{"position":1}]""",
            """[{"position":1},{"position":1.00001}]""",
            new JsonArray(Enumerable.Range(0, 33).Select(i => (JsonNode)new JsonObject { ["position"] = i }).ToArray()).ToJsonString() })
        {
            SetTabs(program, invalid); Assert.Empty(Compile(program, success: false).File);
        }
        foreach (var invalid in new[] { "false", "null", "\"true\"" })
        {
            SetTabs(program, null); FirstTextParagraph(program)["style"]!["noTabStops"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
        foreach (var choice in new[] { "[]", InitialTabs })
        {
            SetTabs(program, choice); FirstTextParagraph(program)["style"]!["noTabStops"] = true;
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphTabStopsPrecedenceSelectsWholeChoice()
    {
        var program = Program("text"); var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["tabStops"] = JsonNode.Parse(InitialTabs) };
        element["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{},"runs":[{"text":"Tabs"}]}]}""");
        foreach (var clear in new[] { "clear", "[]" })
        {
            SetTabs(program, clear); AssertTabs(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", null);
        }
        element["style"]!["paragraph"] = new JsonObject { ["noTabStops"] = true };
        SetTabs(program, ChangedTabs); AssertTabs(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", ChangedTabs);
        FirstTextParagraph(program).Remove("style");
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["tabStops"] = JsonNode.Parse(InitialTabs) } };
        AssertTabs(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", InitialTabs);
    }

    [Fact]
    public void ParagraphTabStopsRetainsUnmodeledSource()
    {
        var program = BulletFontProgram("text"); SetTabs(program, InitialTabs);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var profile in new[] { "missing", "number", "overflow", "negative", "alignment", "list-attribute", "tab-attribute",
            "foreign-position", "child", "duplicate-list", "duplicate-position", "reversed", "over-budget" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var list = properties.GetFirstChild<A.TabStopList>()!; var tab = list.Elements<A.TabStop>().First();
                if (profile == "missing") tab.Position = null;
                else if (profile is "number" or "overflow" or "negative") tab.SetAttribute(new OpenXmlAttribute("pos", "", profile == "number" ? "oops" : profile == "overflow" ? "9999999999999" : "-1"));
                else if (profile == "alignment") tab.SetAttribute(new OpenXmlAttribute("algn", "", "unknown"));
                else if (profile == "list-attribute") list.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "yes"));
                else if (profile == "tab-attribute") tab.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "yes"));
                else if (profile == "foreign-position") { tab.Position = null; tab.SetAttribute(new OpenXmlAttribute("future", "pos", "urn:officekit:test", "381000")); }
                else if (profile == "child") list.InnerXml += """<future:content xmlns:future="urn:officekit:test"/>""";
                else if (profile == "duplicate-list") properties.Append(list.CloneNode(true));
                else if (profile == "duplicate-position") list.Elements<A.TabStop>().Last().Position = tab.Position;
                else if (profile == "over-budget")
                {
                    list.RemoveAllChildren();
                    for (var i = 0; i < 33; i++) list.Append(new A.TabStop { Position = i });
                }
                else list.Elements<A.TabStop>().Last().Position = 1;
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["tabStops"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            foreach (var choice in new[] { ChangedTabs, "[]", "clear" })
            {
                var request = Project(source); SetTabs(request, choice);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                AssertOnlyBodyPropertyChanged(source, Compile(request, source).File.ToByteArray(), "text", "paragraphDefault.bold");
            }
        }
    }

    private static void SetTabs(JsonObject program, string? choice)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject(); var style = paragraph["style"]!.AsObject();
        style.Remove("tabStops"); style.Remove("noTabStops");
        if (choice == "clear") style["noTabStops"] = true;
        else if (choice is not null) style["tabStops"] = JsonNode.Parse(choice);
    }

    private static void AssertTabs(byte[] bytes, string kind, string? choice)
    {
        var expected = choice is null or "clear" ? new JsonArray() : JsonNode.Parse(choice)!.AsArray();
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var list = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.GetFirstChild<A.TabStopList>();
        var style = FirstTextParagraph(Project(bytes))["style"];
        Assert.Null(style?["noTabStops"]);
        if (expected.Count == 0) { Assert.Null(list); Assert.Null(style?["tabStops"]); return; }
        var tabs = list!.Elements<A.TabStop>().ToArray(); var projected = style!["tabStops"]!.AsArray();
        Assert.Equal(expected.Count, tabs.Length); Assert.Equal(expected.Count, projected.Count);
        for (var i = 0; i < expected.Count; i++)
        {
            var position = (int)Math.Round(expected[i]!["position"]!.GetValue<double>() * 12700);
            var alignment = expected[i]!["alignment"]?.GetValue<string>() ?? "left";
            Assert.Equal(position, tabs[i].Position!.Value);
            Assert.Equal(position, (int)Math.Round(projected[i]!["position"]!.GetValue<double>() * 12700));
            Assert.Equal(alignment, projected[i]!["alignment"]!.GetValue<string>());
        }
    }
}
