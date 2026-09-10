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
    public void ParagraphHangingPunctuationLifecyclePreservesSource(string kind)
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(HangingPunctuationProgram(kind)).File.ToByteArray());
        AssertHangingPunctuation(authored, kind, true);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("hangingPunct", "", "true"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("hangingPunct", "", "false"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var value in new bool?[] { false, null, true, false })
        {
            var request = Project(before); SetHangingPunctuation(request, value); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertHangingPunctuation(candidate, kind, value);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.hangingPunctuation");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetHangingPunctuation(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertHangingPunctuation(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.hangingPunctuation");
        var restore = Project(removed); SetHangingPunctuation(restore, true);
        var restored = Compile(restore, removed).File.ToByteArray();
        AssertHangingPunctuation(restored, kind, true);
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.hangingPunctuation");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetHangingPunctuation(denied, false);
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.hangingPunctuation"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"hangingPunctuation":false},"runs":[{"text":""}]}]}
            """);
        var onlyPunctuation = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertHangingPunctuation(onlyPunctuation, kind, false);
        var removeStyle = Project(onlyPunctuation); FirstTextParagraph(removeStyle).Remove("style");
        AssertHangingPunctuation(Compile(removeStyle, onlyPunctuation).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphHangingPunctuationAuthoringAndPrecedence()
    {
        var program = HangingPunctuationProgram("text");
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["hangingPunctuation"] = true };
        foreach (var value in new[] { false, true })
        {
            SetHangingPunctuation(program, value);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertHangingPunctuation(bytes, "text", value);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Equal(-152_400, properties.Indent!.Value);
            Assert.Equal(304_800, properties.LeftMargin!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Left, properties.Alignment!.Value);
            Assert.True(properties.RightToLeft!.Value);
        }
        SetHangingPunctuation(program, null);
        AssertHangingPunctuation(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", true);
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["hangingPunctuation"] = false } };
        AssertHangingPunctuation(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", false);
        foreach (var invalid in new[] { "null", "0", "1", "\"false\"", "\"true\"", "{}", "[]" })
        {
            FirstTextParagraph(program)["style"]!["hangingPunctuation"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphHangingPunctuationRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(HangingPunctuationProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "TRUE", "2", "invalid", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("hangingPunct", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["hangingPunctuation"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetHangingPunctuation(projected, false);
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["hangingPunctuation"]);
        }
    }

    private static JsonObject HangingPunctuationProgram(string kind)
    {
        var program = DirectionProgram(kind);
        SetHangingPunctuation(program, true);
        var paragraph = FirstTextParagraph(program);
        paragraph["style"]!["hanging"] = 12;
        paragraph["style"]!["indent"] = 24;
        paragraph["runs"]![0]!["text"] = "中文标点，句末。";
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["hangingPunctuation"] = false;
        return program;
    }

    private static void SetHangingPunctuation(JsonObject program, bool? value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (value is null) paragraph["style"]!.AsObject().Remove("hangingPunctuation");
        else paragraph["style"]!["hangingPunctuation"] = value.Value;
    }

    private static void AssertHangingPunctuation(byte[] bytes, string kind, bool? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var value = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.Height;
        Assert.Equal(expected, value?.Value);
        Assert.Equal(expected, FirstTextParagraph(Project(bytes))["style"]?["hangingPunctuation"]?.GetValue<bool>());
    }
}
