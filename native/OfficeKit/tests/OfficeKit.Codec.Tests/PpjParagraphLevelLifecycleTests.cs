using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphLevelLifecyclePreservesSource(string kind)
    {
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{"level":2},"runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
                {"style":{"level":8},"runs":[{"text":"Neighbor"}]}]}
                """);
            Style(program, kind)["paragraph"] = new JsonObject { ["level"] = 4 };
            AssertParagraphLevel(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), kind, 2);
            var inherited = program.DeepClone().AsObject(); FirstTextParagraph(inherited).Remove("style");
            AssertParagraphLevel(PptxCodecTests.RemoveEmbeddedPpj(Compile(inherited).File.ToByteArray()), kind, 4);
            if (mixed)
            {
                var style = FirstTextParagraph(program)["style"]!;
                style["alignment"] = "thaiDistributed"; style["indent"] = 7; style["hanging"] = 3.5;
                style["spaceBefore"] = 0; style["spaceAfter"] = 9; style["lineSpacingMultiplier"] = 1.25;
                style["bullet"] = new JsonObject { ["type"] = "character", ["character"] = "•" };
                style["defaultText"] = JsonNode.Parse("""{"bold":true,"shadow":{"color":"#112233","scaleY":-0.5}}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
                paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("lvl", "", "+0002"));
                paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("lvl", "", "+0008"));
                if (mixed)
                {
                    var properties = paragraphs[0].ParagraphProperties!;
                    properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
                    properties.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                    properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            foreach (var level in new[] { 0, 3, 8 })
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["level"] = level;
                var changed = Compile(request, source).File.ToByteArray();
                AssertParagraphLevel(changed, kind, level);
                AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.level");
                var remove = Project(changed); FirstTextParagraph(remove)["style"]!.AsObject().Remove("level");
                var removed = Compile(remove, changed).File.ToByteArray();
                AssertParagraphLevel(removed, kind, null);
                AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph.level");
                var restore = Project(removed); FirstTextParagraph(restore)["style"] ??= new JsonObject();
                FirstTextParagraph(restore)["style"]!["level"] = level;
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertParagraphLevel(restored, kind, level);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.level");
            }
            if (!mixed)
            {
                var request = Project(source); FirstTextParagraph(request).Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertParagraphLevel(removed, kind, null);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.level");
            }
            else
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                AssertOnlyBodyPropertyChanged(source, Compile(request, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            }
            var denied = Project(source); FirstTextParagraph(denied)["style"]!["level"] = 0;
            var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
            fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.level"));
            Assert.Empty(Compile(denied, source, success: false).File);
            foreach (var invalid in new[] { "null", "false", "-1", "9", "1.5", "2147483648", "\"2\"", "{}" })
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["level"] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void ParagraphLevelRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"level":2,"defaultText":{"bold":true}},"runs":[{"text":"Retain source level"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "-1", "9", "invalid", "1.5", "1e0", "4294967296", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("lvl", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["level"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["level"] = 0;
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var changed = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(changed))["style"]?["level"]);
            }
        }
    }

    private static void AssertParagraphLevel(byte[] file, string kind, int? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties;
        Assert.Equal(expected, properties?.Level?.Value);
        Assert.Equal(expected, FirstTextParagraph(Project(file))["style"]?["level"]?.GetValue<int>());
    }
}
