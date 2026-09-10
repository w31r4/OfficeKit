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
    public void ParagraphAlignmentLifecyclePreservesSource(string kind)
    {
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{"alignment":"center"},"runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
                {"style":{"alignment":"distributed"},"runs":[{"text":"Neighbor"}]}]}
                """);
            Style(program, kind)["paragraph"] = new JsonObject { ["alignment"] = "right" };
            AssertParagraphAlignment(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), kind, "center", "ctr");
            var inherited = program.DeepClone().AsObject(); FirstTextParagraph(inherited).Remove("style");
            AssertParagraphAlignment(PptxCodecTests.RemoveEmbeddedPpj(Compile(inherited).File.ToByteArray()), kind, "right", "r");
            if (mixed)
            {
                var style = FirstTextParagraph(program)["style"]!;
                style["indent"] = 7; style["hanging"] = 3.5; style["spaceBefore"] = 0;
                style["spaceAfter"] = 9; style["lineSpacingMultiplier"] = 1.25;
                style["defaultText"] = JsonNode.Parse("""{"bold":true,"shadow":{"color":"#112233","scaleY":-0.5}}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            if (mixed)
            {
                using var document = PresentationDocument.Open(stream, true);
                var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
                var properties = paragraphs[0].ParagraphProperties!;
                properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
                properties.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("algn", "", "futureAlignment"));
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            foreach (var (alignment, token) in new[] { ("left", "l"), ("center", "ctr"), ("right", "r"), ("justify", "just"), ("distributed", "dist"),
                ("justifyLow", "justLow"), ("thaiDistributed", "thaiDist") })
            {
                var authoredRequest = program.DeepClone().AsObject(); FirstTextParagraph(authoredRequest)["style"]!["alignment"] = alignment;
                AssertParagraphAlignment(PptxCodecTests.RemoveEmbeddedPpj(Compile(authoredRequest).File.ToByteArray()), kind, alignment, token);
                var request = Project(source); FirstTextParagraph(request)["style"]!["alignment"] = alignment;
                var changed = Compile(request, source).File.ToByteArray();
                AssertParagraphAlignment(changed, kind, alignment, token);
                AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.alignment");
                var remove = Project(changed); FirstTextParagraph(remove)["style"]!.AsObject().Remove("alignment");
                var removed = Compile(remove, changed).File.ToByteArray();
                AssertParagraphAlignment(removed, kind, null, null);
                AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph.alignment");
                var restore = Project(removed); FirstTextParagraph(restore)["style"] ??= new JsonObject();
                FirstTextParagraph(restore)["style"]!["alignment"] = alignment;
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertParagraphAlignment(restored, kind, alignment, token);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.alignment");
            }
            if (!mixed)
            {
                var request = Project(source); FirstTextParagraph(request).Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertParagraphAlignment(removed, kind, null, null);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.alignment");
            }
            else
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                AssertOnlyBodyPropertyChanged(source, Compile(request, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            }
            var denied = Project(source); FirstTextParagraph(denied)["style"]!["alignment"] = "left";
            var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
            fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.alignment"));
            Assert.Empty(Compile(denied, source, success: false).File);
            foreach (var invalid in new[] { "null", "false", "0", "\"l\"", "\"Center\"", "\"\"", "{}" })
            {
                var request = Project(source); FirstTextParagraph(request)["style"]!["alignment"] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void ParagraphAlignmentRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"alignment":"center","defaultText":{"bold":true}},"runs":[{"text":"Retain source alignment"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "justHigh", "thaiDistributed", "future", "center", "CTR", " ctr ", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("algn", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["alignment"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["alignment"] = "left";
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var changed = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraphDefault.bold");
                AssertParagraphAlignment(changed, "text", null, token);
            }
        }
    }

    [Theory]
    [InlineData("justifyLow", "justLow")]
    [InlineData("thaiDistributed", "thaiDist")]
    public void ParagraphAlignmentModesRetainTableAndMasterDefaults(string alignment, string token)
    {
        var master = Program("master");
        master["design"]!["masters"]![0]!["textStyles"] = new JsonObject
        {
            ["title"] = new JsonArray(new JsonObject { ["level"] = 0, ["alignment"] = alignment }),
        };
        var masterSource = PptxCodecTests.RemoveEmbeddedPpj(Compile(master).File.ToByteArray());
        var masterProjected = Project(masterSource);
        Assert.Equal(alignment, masterProjected["design"]!["masters"]![0]!["textStyles"]!["title"]![0]!["alignment"]!.GetValue<string>());
        Assert.Equal(masterSource, Compile(masterProjected, masterSource).File.ToByteArray());
        using (var document = PresentationDocument.Open(new MemoryStream(masterSource), false))
            Assert.Equal(token, document.PresentationPart!.SlideMasterParts.Single().SlideMaster
                .Descendants<A.Level1ParagraphProperties>().Single().Alignment!.InnerText);

        static JsonNode CellParagraph(JsonObject program) => program["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!;
        var table = Program("table"); CellParagraph(table)["style"] = new JsonObject { ["alignment"] = alignment };
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(table).File.ToByteArray());
        var projected = Project(source);
        Assert.Equal("table", projected["pages"]![0]!["elements"]![0]!["type"]!.GetValue<string>());
        Assert.Equal(alignment, CellParagraph(projected)["style"]!["alignment"]!.GetValue<string>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        var next = alignment == "justifyLow" ? "thaiDistributed" : "justifyLow";
        CellParagraph(projected)["style"]!["alignment"] = next;
        var changed = Compile(projected, source).File.ToByteArray();
        Assert.Equal(next, CellParagraph(Project(changed))["style"]!["alignment"]!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(source, changed, "table", "paragraph.alignment");
        using var result = PresentationDocument.Open(new MemoryStream(changed), false);
        Assert.Equal(next == "justifyLow" ? "justLow" : "thaiDist",
            Owner(result, "table").Descendants<A.Paragraph>().First().ParagraphProperties!.Alignment!.InnerText);
    }

    private static void AssertParagraphAlignment(byte[] file, string kind, string? expected, string? nativeToken)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties;
        var actual = properties?.GetAttributes().FirstOrDefault(a => a.NamespaceUri.Length == 0 && a.LocalName == "algn").Value;
        Assert.Equal(nativeToken, actual);
        Assert.Equal(expected, FirstTextParagraph(Project(file))["style"]?["alignment"]?.GetValue<string>());
    }
}
