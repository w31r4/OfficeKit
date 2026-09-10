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
    public void ParagraphLatinLineBreakLifecyclePreservesSource(string kind)
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(LatinLineBreakProgram(kind)).File.ToByteArray());
        AssertLatinLineBreak(authored, kind, true);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("latinLnBrk", "", "true"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("latinLnBrk", "", "false"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var value in new bool?[] { false, null, true, false })
        {
            var request = Project(before); SetLatinLineBreak(request, value); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertLatinLineBreak(candidate, kind, value);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.latinLineBreak");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetLatinLineBreak(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertLatinLineBreak(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.latinLineBreak");
        var restore = Project(removed); SetLatinLineBreak(restore, true);
        var restored = Compile(restore, removed).File.ToByteArray();
        AssertLatinLineBreak(restored, kind, true);
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.latinLineBreak");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetLatinLineBreak(denied, false);
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.latinLineBreak"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"latinLineBreak":false},"runs":[{"text":""}]}]}
            """);
        var onlyLineBreak = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertLatinLineBreak(onlyLineBreak, kind, false);
        var removeStyle = Project(onlyLineBreak); FirstTextParagraph(removeStyle).Remove("style");
        AssertLatinLineBreak(Compile(removeStyle, onlyLineBreak).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphLatinLineBreakAuthoringAndPrecedence()
    {
        var program = LatinLineBreakProgram("text");
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["latinLineBreak"] = true };
        foreach (var value in new[] { false, true })
        {
            element["style"]!["wrap"] = value ? "none" : "square";
            SetLatinLineBreak(program, value);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertLatinLineBreak(bytes, "text", value);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
            Assert.Equal(-152_400, properties.Indent!.Value);
            Assert.Equal(304_800, properties.LeftMargin!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Left, properties.Alignment!.Value);
            Assert.True(properties.RightToLeft!.Value);
            Assert.True(properties.Height!.Value);
            Assert.Single(Owner(document, "text").Descendants<A.Break>());
            Assert.Equal(value ? A.TextWrappingValues.None : A.TextWrappingValues.Square,
                Owner(document, "text").Descendants<A.BodyProperties>().First().Wrap!.Value);
        }
        SetLatinLineBreak(program, null);
        AssertLatinLineBreak(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", true);
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["latinLineBreak"] = false } };
        AssertLatinLineBreak(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", false);
        foreach (var invalid in new[] { "null", "0", "1", "\"false\"", "\"true\"", "{}", "[]" })
        {
            FirstTextParagraph(program)["style"]!["latinLineBreak"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphLatinLineBreakRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(LatinLineBreakProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "TRUE", "2", "invalid", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("latinLnBrk", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["latinLineBreak"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetLatinLineBreak(projected, false);
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["latinLineBreak"]);
        }
    }

    private static JsonObject LatinLineBreakProgram(string kind)
    {
        var program = DirectionProgram(kind);
        SetLatinLineBreak(program, true);
        var paragraph = FirstTextParagraph(program);
        paragraph["style"]!["hanging"] = 12;
        paragraph["style"]!["indent"] = 24;
        paragraph["style"]!["hangingPunctuation"] = true;
        paragraph["runs"] = JsonNode.Parse("""
            [{"text":"extraordinarilylongword","style":{"language":"en-US","color":"#334455"}},
             {"break":true},{"text":"已有手动换行。"}]
            """);
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["latinLineBreak"] = false;
        return program;
    }

    private static void SetLatinLineBreak(JsonObject program, bool? value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (value is null) paragraph["style"]!.AsObject().Remove("latinLineBreak");
        else paragraph["style"]!["latinLineBreak"] = value.Value;
    }

    private static void AssertLatinLineBreak(byte[] bytes, string kind, bool? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var value = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.LatinLineBreak;
        Assert.Equal(expected, value?.Value);
        Assert.Equal(expected, FirstTextParagraph(Project(bytes))["style"]?["latinLineBreak"]?.GetValue<bool>());
    }
}
