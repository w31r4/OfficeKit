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
    public void ParagraphFontAlignmentLifecyclePreservesSource(string kind)
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(FontAlignmentProgram(kind)).File.ToByteArray());
        AssertFontAlignment(authored, kind, "baseline");
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
            Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!
                .SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var alignment in new[] { "auto", "top", "center", "bottom", null, "baseline", "auto" })
        {
            var request = Project(before); SetFontAlignment(request, alignment); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertFontAlignment(candidate, kind, alignment);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.fontAlignment");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetFontAlignment(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertFontAlignment(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.fontAlignment");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetFontAlignment(denied, "top");
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.fontAlignment"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"fontAlignment":"auto"},"runs":[{"text":""}]}]}
            """);
        var onlyAlignment = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertFontAlignment(onlyAlignment, kind, "auto");
        var removeStyle = Project(onlyAlignment); FirstTextParagraph(removeStyle).Remove("style");
        AssertFontAlignment(Compile(removeStyle, onlyAlignment).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphFontAlignmentAuthoringAndPrecedence()
    {
        var program = FontAlignmentProgram("text");
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["fontAlignment"] = "bottom" };
        foreach (var alignment in new[] { "auto", "top", "center", "baseline", "bottom" })
        {
            SetFontAlignment(program, alignment);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertFontAlignment(bytes, "text", alignment);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var paragraph = Owner(document, "text").Descendants<A.Paragraph>().First();
            Assert.Equal(A.TextAlignmentTypeValues.Left, paragraph.ParagraphProperties!.Alignment!.Value);
            Assert.True(paragraph.ParagraphProperties.RightToLeft!.Value);
            Assert.Equal(new[] { 1400, 3200 }, paragraph.Elements<A.Run>().Select(run => run.RunProperties!.FontSize!.Value));
        }
        SetFontAlignment(program, null);
        AssertFontAlignment(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "bottom");
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["fontAlignment"] = "center" } };
        AssertFontAlignment(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "center");
        foreach (var invalid in new[] { "null", "false", "0", "\"t\"", "\"automatic\"", "\"\"", "{}" })
        {
            FirstTextParagraph(program)["style"]!["fontAlignment"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphFontAlignmentRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(FontAlignmentProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "AUTO", "top", "invalid", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("fontAlgn", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["fontAlignment"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetFontAlignment(projected, "auto");
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["fontAlignment"]);
        }
    }

    private static JsonObject FontAlignmentProgram(string kind)
    {
        var program = DirectionProgram(kind);
        SetFontAlignment(program, "baseline");
        FirstTextParagraph(program)["runs"] = JsonNode.Parse("""
            [{"text":"Small 中文","style":{"size":14,"color":"#334455"}},
             {"text":" Large","style":{"size":32}}]
            """);
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["fontAlignment"] = "auto";
        return program;
    }

    private static void SetFontAlignment(JsonObject program, string? alignment)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (alignment is null) paragraph["style"]!.AsObject().Remove("fontAlignment");
        else paragraph["style"]!["fontAlignment"] = alignment;
    }

    private static void AssertFontAlignment(byte[] bytes, string kind, string? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties;
        var token = properties?.GetAttributes().FirstOrDefault(a => a.NamespaceUri.Length == 0 && a.LocalName == "fontAlgn").Value;
        Assert.Equal(expected switch { "auto" => "auto", "top" => "t", "center" => "ctr", "baseline" => "base", "bottom" => "b", _ => null }, token);
        Assert.Equal(expected, FirstTextParagraph(Project(bytes))["style"]?["fontAlignment"]?.GetValue<string>());
    }
}
