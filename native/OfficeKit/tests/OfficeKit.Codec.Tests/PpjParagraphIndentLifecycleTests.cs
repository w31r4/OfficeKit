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
    public void ParagraphIndentLifecyclePreservesSource(string kind)
    {
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{"indent":7.25},"runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
                {"style":{"indent":0.000079},"runs":[{"text":"Neighbor"}]}]}
                """);
            Style(program, kind)["paragraph"] = new JsonObject { ["indent"] = 18 };
            AssertParagraphIndent(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), kind, 92_075);
            var inherited = program.DeepClone().AsObject(); FirstTextParagraph(inherited).Remove("style");
            AssertParagraphIndent(PptxCodecTests.RemoveEmbeddedPpj(Compile(inherited).File.ToByteArray()), kind, 228_600);
            if (mixed)
            {
                var style = FirstTextParagraph(program)["style"]!;
                style["hanging"] = 3.5; style["spaceBefore"] = 0; style["spaceAfter"] = 9; style["lineSpacingMultiplier"] = 1.25;
                style["defaultText"] = JsonNode.Parse("""{"bold":true,"shadow":{"color":"#112233","scaleY":-0.5}}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
                paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("marL", "", "+0092075"));
                paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("marL", "", "+0000001"));
                if (mixed)
                {
                    var properties = paragraphs[0].ParagraphProperties!;
                    properties.SetAttribute(new OpenXmlAttribute("indent", "", "-044450"));
                    properties.SetAttribute(new OpenXmlAttribute("marR", "", "+076200"));
                    properties.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                    properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            foreach (var (points, emu) in new[] { (0d, 0L), (0.000001, 0L), (0.005, 64L), (0.015, 190L), (1.234567, 15_679L), (4032d, 51_206_400L) })
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["indent"] = points;
                var changed = Compile(request, source).File.ToByteArray();
                AssertParagraphIndent(changed, kind, emu);
                AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.indent");
                var remove = Project(changed); FirstTextParagraph(remove)["style"]!.AsObject().Remove("indent");
                var removed = Compile(remove, changed).File.ToByteArray();
                AssertParagraphIndent(removed, kind, null);
                AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph.indent");
                var restore = Project(removed); FirstTextParagraph(restore)["style"] ??= new JsonObject();
                FirstTextParagraph(restore)["style"]!["indent"] = points;
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertParagraphIndent(restored, kind, emu);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.indent");
            }
            if (!mixed)
            {
                var request = Project(source); FirstTextParagraph(request).Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertParagraphIndent(removed, kind, null);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.indent");
            }
            else
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                AssertOnlyBodyPropertyChanged(source, Compile(request, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            }
            var denied = Project(source); FirstTextParagraph(denied)["style"]!["indent"] = 2.5;
            var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
            fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.indent"));
            Assert.Empty(Compile(denied, source, success: false).File);
            foreach (var invalid in new[] { "null", "false", "-0.000001", "4032.000001", "\"18\"", "{}" })
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["indent"] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void ParagraphIndentRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"indent":7,"defaultText":{"bold":true}},"runs":[{"text":"Retain source margin"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "-1", "51206401", "2147483648", "invalid", "1.5", "1e3", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("marL", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["indent"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["indent"] = 3;
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var changed = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(changed))["style"]?["indent"]);
            }
        }
    }

    private static void AssertParagraphIndent(byte[] file, string kind, long? expectedEmu)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var native = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.LeftMargin;
        var projected = FirstTextParagraph(Project(file))["style"]?["indent"];
        if (expectedEmu is null) { Assert.Null(native); Assert.Null(projected); return; }
        Assert.NotNull(native); Assert.Equal(expectedEmu.Value, native.Value);
        Assert.Equal(Math.Round(expectedEmu.Value / 12700d, 6, MidpointRounding.AwayFromZero), projected!.GetValue<double>(), 6);
    }
}
