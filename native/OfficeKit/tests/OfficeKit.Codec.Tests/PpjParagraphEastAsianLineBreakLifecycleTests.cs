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
    public void ParagraphEastAsianLineBreakLifecyclePreservesSource(string kind)
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(EastAsianLineBreakProgram(kind)).File.ToByteArray());
        AssertEastAsianLineBreak(authored, kind, true);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("eaLnBrk", "", "true"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("eaLnBrk", "", "false"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var value in new bool?[] { false, null, true, false })
        {
            var request = Project(before); SetEastAsianLineBreak(request, value); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertEastAsianLineBreak(candidate, kind, value);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.eastAsianLineBreak");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetEastAsianLineBreak(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertEastAsianLineBreak(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.eastAsianLineBreak");
        var restore = Project(removed); SetEastAsianLineBreak(restore, true);
        var restored = Compile(restore, removed).File.ToByteArray();
        AssertEastAsianLineBreak(restored, kind, true);
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.eastAsianLineBreak");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetEastAsianLineBreak(denied, false);
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.eastAsianLineBreak"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"eastAsianLineBreak":false},"runs":[{"text":""}]}]}
            """);
        var onlyLineBreak = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertEastAsianLineBreak(onlyLineBreak, kind, false);
        var removeStyle = Project(onlyLineBreak); FirstTextParagraph(removeStyle).Remove("style");
        AssertEastAsianLineBreak(Compile(removeStyle, onlyLineBreak).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphEastAsianLineBreakAuthoringAndPrecedence()
    {
        var program = EastAsianLineBreakProgram("text");
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["eastAsianLineBreak"] = true };
        foreach (var value in new[] { false, true })
        {
            element["style"]!["wrap"] = value ? "none" : "square";
            SetEastAsianLineBreak(program, value);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertEastAsianLineBreak(bytes, "text", value);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Equal(-152_400, properties.Indent!.Value);
            Assert.Equal(304_800, properties.LeftMargin!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Left, properties.Alignment!.Value);
            Assert.True(properties.RightToLeft!.Value);
            Assert.True(properties.Height!.Value);
            Assert.True(properties.LatinLineBreak!.Value);
            Assert.Single(Owner(document, "text").Descendants<A.Break>());
            Assert.Equal(value ? A.TextWrappingValues.None : A.TextWrappingValues.Square,
                Owner(document, "text").Descendants<A.BodyProperties>().First().Wrap!.Value);
        }
        SetEastAsianLineBreak(program, null);
        AssertEastAsianLineBreak(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", true);
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["eastAsianLineBreak"] = false } };
        AssertEastAsianLineBreak(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", false);
        foreach (var invalid in new[] { "null", "0", "1", "\"false\"", "\"true\"", "{}", "[]" })
        {
            FirstTextParagraph(program)["style"]!["eastAsianLineBreak"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphEastAsianLineBreakRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(EastAsianLineBreakProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "TRUE", "2", "invalid", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("eaLnBrk", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["eastAsianLineBreak"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetEastAsianLineBreak(projected, false);
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["eastAsianLineBreak"]);
        }
    }

    private static JsonObject EastAsianLineBreakProgram(string kind)
    {
        var program = DirectionProgram(kind);
        SetEastAsianLineBreak(program, true);
        var paragraph = FirstTextParagraph(program);
        paragraph["style"]!["hanging"] = 12;
        paragraph["style"]!["indent"] = 24;
        paragraph["style"]!["hangingPunctuation"] = true;
        paragraph["style"]!["latinLineBreak"] = true;
        paragraph["runs"] = JsonNode.Parse("""
            [{"text":"「中文标点」与（日文）、句末。","style":{"language":"zh-CN","color":"#334455"}},
             {"break":true},{"text":"已有手动换行。"}]
            """);
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["eastAsianLineBreak"] = false;
        return program;
    }

    private static void SetEastAsianLineBreak(JsonObject program, bool? value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (value is null) paragraph["style"]!.AsObject().Remove("eastAsianLineBreak");
        else paragraph["style"]!["eastAsianLineBreak"] = value.Value;
    }

    private static void AssertEastAsianLineBreak(byte[] bytes, string kind, bool? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var value = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.EastAsianLineBreak;
        Assert.Equal(expected, value?.Value);
        Assert.Equal(expected, FirstTextParagraph(Project(bytes))["style"]?["eastAsianLineBreak"]?.GetValue<bool>());
    }
}
