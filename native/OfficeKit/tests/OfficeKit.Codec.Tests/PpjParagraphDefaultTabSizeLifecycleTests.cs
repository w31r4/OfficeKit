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
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultTabSizeLifecyclePreservesSource(string kind)
    {
        var program = DefaultTabSizeProgram(kind);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        AssertDefaultTabSize(authored, kind, 18.25);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("defTabSz", "", "+00231775"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("defTabSz", "", "+000000"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var value in new double?[] { 0, -25.12345, null, 40, 0 })
        {
            var request = Project(before); SetDefaultTabSize(request, value); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertDefaultTabSize(candidate, kind, value);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.defaultTabSize");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetDefaultTabSize(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertDefaultTabSize(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.defaultTabSize");
        var restore = Project(removed); SetDefaultTabSize(restore, 18.25);
        var restored = Compile(restore, removed).File.ToByteArray();
        AssertDefaultTabSize(restored, kind, 18.25);
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.defaultTabSize");
        var clearTabs = Project(source);
        FirstTextParagraph(clearTabs)["style"]!.AsObject().Remove("tabStops");
        FirstTextParagraph(clearTabs)["style"]!["noTabStops"] = true;
        var withoutTabs = Compile(clearTabs, source).File.ToByteArray();
        AssertDefaultTabSize(withoutTabs, kind, 18.25);
        AssertOnlyBodyPropertyChanged(source, withoutTabs, kind, "paragraph.tabStops");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetDefaultTabSize(denied, 0);
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.defaultTabSize"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultTabSize":0},"runs":[{"text":""}]}]}
            """);
        var onlySize = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertDefaultTabSize(onlySize, kind, 0);
        var removeStyle = Project(onlySize); FirstTextParagraph(removeStyle).Remove("style");
        AssertDefaultTabSize(Compile(removeStyle, onlySize).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphDefaultTabSizeBoundsPrecisionAndPrecedence()
    {
        var program = DefaultTabSizeProgram("text");
        foreach (var value in new[] { int.MinValue / 12700d, -25.12345, -1.5 / 12700, -0.5 / 12700, 0, 0.5 / 12700, 1.5 / 12700, 25.12345, int.MaxValue / 12700d })
        {
            SetDefaultTabSize(program, value);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertDefaultTabSize(bytes, "text", value);
            Assert.Equal(bytes, Compile(Project(bytes), bytes).File.ToByteArray());
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var paragraphs = Owner(document, "text").Descendants<A.Paragraph>().ToArray();
            Assert.Equal(new[] { 304_800, 609_600 }, paragraphs[0].ParagraphProperties!.GetFirstChild<A.TabStopList>()!
                .Elements<A.TabStop>().Select(tab => tab.Position!.Value));
            Assert.Null(paragraphs[1].ParagraphProperties!.GetFirstChild<A.TabStopList>());
            Assert.Equal("A\tB\tC", paragraphs[0].InnerText);
            Assert.Equal(152_400, paragraphs[0].ParagraphProperties!.LeftMargin!.Value);
            Assert.Equal(-50_800, paragraphs[0].ParagraphProperties!.Indent!.Value);
        }
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["defaultTabSize"] = 50 };
        SetDefaultTabSize(program, 0);
        AssertDefaultTabSize(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 0);
        SetDefaultTabSize(program, null);
        AssertDefaultTabSize(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 50);
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["defaultTabSize"] = 25.5 } };
        AssertDefaultTabSize(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 25.5);
        foreach (var invalid in new[] { "null", "false", "\"12\"", "-169093.2007", "169093.2006", "{}", "[]" })
        {
            FirstTextParagraph(program)["style"]!["defaultTabSize"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphDefaultTabSizeRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(DefaultTabSizeProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "-2147483649", "2147483648", "1.5", "invalid", "9223372036854775808", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("defTabSz", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["defaultTabSize"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetDefaultTabSize(projected, 0);
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["defaultTabSize"]);
        }
    }

    private static JsonObject DefaultTabSizeProgram(string kind)
    {
        var program = DirectionProgram(kind);
        var style = FirstTextParagraph(program)["style"]!;
        style["defaultTabSize"] = 18.25; style["indent"] = 12; style["hanging"] = 4;
        style["tabStops"] = JsonNode.Parse("""[{"position":24,"alignment":"left"},{"position":48,"alignment":"decimal"}]""");
        FirstTextParagraph(program)["runs"] = JsonNode.Parse("""[{"text":"A\tB\tC"}]""");
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["noTabStops"] = true;
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["defaultTabSize"] = 0;
        return program;
    }

    private static void SetDefaultTabSize(JsonObject program, double? value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (value is null) paragraph["style"]!.AsObject().Remove("defaultTabSize");
        else paragraph["style"]!["defaultTabSize"] = value.Value;
    }

    private static void AssertDefaultTabSize(byte[] bytes, string kind, double? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var margin = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.DefaultTabSize;
        var projected = FirstTextParagraph(Project(bytes))["style"]?["defaultTabSize"];
        if (expected is null) { Assert.Null(margin); Assert.Null(projected); return; }
        var emu = (long)Math.Round(expected.Value * 12700);
        Assert.NotNull(margin); Assert.Equal(emu, margin.Value);
        Assert.NotNull(projected); Assert.Equal(emu, (long)Math.Round(projected.GetValue<double>() * 12700));
    }
}
