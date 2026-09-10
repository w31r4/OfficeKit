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
    public void ParagraphDirectionLifecyclePreservesSource(string kind)
    {
        var program = DirectionProgram(kind);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        AssertDirection(authored, kind, "right-to-left");
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("rtl", "", "true"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("rtl", "", "false"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var direction in new[] { "left-to-right", null, "right-to-left", "left-to-right" })
        {
            var request = Project(before); SetDirection(request, direction); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertDirection(candidate, kind, direction);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.direction");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        // Remove from the original bytes, not a previously edited artifact.
        var remove = Project(source); SetDirection(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertDirection(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.direction");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetDirection(denied, "left-to-right");
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.direction"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"direction":"left-to-right"},"runs":[{"text":""}]}]}
            """);
        var onlyDirection = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertDirection(onlyDirection, kind, "left-to-right");
        var removeStyle = Project(onlyDirection); FirstTextParagraph(removeStyle).Remove("style");
        AssertDirection(Compile(removeStyle, onlyDirection).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphDirectionAuthoringAndPrecedence()
    {
        var program = DirectionProgram("text");
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["direction"] = "right-to-left" };
        element["style"]!["columnDirection"] = "right-to-left";
        SetDirection(program, "left-to-right");
        var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        AssertDirection(bytes, "text", "left-to-right");
        using (var document = PresentationDocument.Open(new MemoryStream(bytes), false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
            var body = Owner(document, "text").Descendants<A.BodyProperties>().First();
            Assert.True(body.RightToLeftColumns!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Left, Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.Alignment!.Value);
        }
        SetDirection(program, null);
        AssertDirection(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "right-to-left");
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["direction"] = "left-to-right" } };
        AssertDirection(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "left-to-right");
        foreach (var invalid in new[] { "null", "false", "0", "\"rtl\"", "\"auto\"", "\"\"", "{}" })
        {
            FirstTextParagraph(program)["style"]!["direction"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphDirectionRetainsUnmodeledSource()
    {
        var program = DirectionProgram("text");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "TRUE", "2", "invalid", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("rtl", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["direction"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetDirection(projected, "left-to-right");
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["direction"]);
        }
    }

    private static JsonObject DirectionProgram(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"direction":"right-to-left","alignment":"left","defaultText":{"bold":true}},
             "runs":[{"text":"مرحبا שלום Office","style":{"color":"#334455"}}]},
             {"style":{"direction":"left-to-right"},"runs":[{"text":"Neighbor"}]}]}
            """);
        return program;
    }

    private static void SetDirection(JsonObject program, string? direction)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (direction is null) paragraph["style"]!.AsObject().Remove("direction");
        else paragraph["style"]!["direction"] = direction;
    }

    private static void AssertDirection(byte[] bytes, string kind, string? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var rtl = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.RightToLeft;
        if (expected is null) Assert.Null(rtl);
        else { Assert.NotNull(rtl); Assert.Equal(expected == "right-to-left", rtl.Value); }
        Assert.Equal(expected, FirstTextParagraph(Project(bytes))["style"]?["direction"]?.GetValue<string>());
    }
}
