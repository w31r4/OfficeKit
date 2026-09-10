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
    public void BulletStartAtLifecyclePreservesSource(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"level":2,"alignment":"thaiDistributed","indent":7,"hanging":3.5,
            "spaceBefore":0,"spaceAfter":9,"lineSpacingMultiplier":1.25,
            "bullet":{"type":"number","format":"decimal","startAt":3,"fontFamily":"Georgia","sizePercent":1.5,"color":"#112233"},
            "defaultText":{"bold":true,"shadow":{"color":"#112233","scaleY":-0.5}}},
            "runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
            {"style":{"bullet":{"type":"number","scheme":"thaiNumParenBoth","startAt":7}},"runs":[{"text":"Neighbor"}]},
            {"style":{"bullet":{"type":"character","character":"•"}},"runs":[{"text":"Character neighbor"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        AssertBulletStartAt(authored, kind, 3);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            var properties = paragraphs[0].ParagraphProperties!;
            var number = properties.GetFirstChild<A.AutoNumberedBullet>()!;
            number.SetAttribute(new OpenXmlAttribute("startAt", "", "+0003"));
            number.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
            paragraphs[1].ParagraphProperties!.GetFirstChild<A.AutoNumberedBullet>()!
                .SetAttribute(new OpenXmlAttribute("startAt", "", "+0007"));
            properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
            properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        foreach (var start in new[] { 1, 8, 32767 })
        {
            var request = Project(source); FirstBullet(request)["startAt"] = start;
            var requestJson = request.ToJsonString();
            var changed = Compile(request, source).File.ToByteArray();
            Assert.Equal(requestJson, request.ToJsonString());
            AssertBulletStartAt(changed, kind, start);
            AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.bullet.startAt");
            var remove = Project(changed); FirstBullet(remove).Remove("startAt");
            var removed = Compile(remove, changed).File.ToByteArray();
            AssertBulletStartAt(removed, kind, null);
            AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph.bullet.startAt");
            var restore = Project(removed); FirstBullet(restore)["startAt"] = start;
            var restored = Compile(restore, removed).File.ToByteArray();
            AssertBulletStartAt(restored, kind, start);
            AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.bullet.startAt");
        }
        var unrelated = Project(source);
        FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); FirstBullet(denied)["startAt"] = 1;
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.bullet.startAt"));
        Assert.Empty(Compile(denied, source, success: false).File);
        foreach (var invalid in new[] { "null", "false", "0", "-1", "32768", "1.5", "2147483648", "\"2\"", "{}" })
        {
            var request = Project(source); FirstBullet(request)["startAt"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        // Number formatting and removal of direct font/color/size are all
        // modeled now. A combined request must retain exactly that PPJ state.
        var restyle = Project(source);
        FirstTextParagraph(restyle)["style"]!["bullet"] = JsonNode.Parse("""{"type":"number","scheme":"romanUcPeriod","startAt":1}""");
        var restyled = Compile(restyle, source).File.ToByteArray();
        var observed = FirstBullet(Project(restyled));
        Assert.Equal(3, observed.Count);
        Assert.Equal("number", observed["type"]!.GetValue<string>());
        Assert.Equal("romanUcPeriod", observed["scheme"]!.GetValue<string>());
        Assert.Equal(1, observed["startAt"]!.GetValue<int>());
        AssertBulletSize(restyled, kind, "{}");
        Assert.Equal(restyled, Compile(Project(restyled), restyled).File.ToByteArray());
        foreach (var replacement in new[] { """{"type":"character","character":"*"}""", """{"type":"none"}""" })
        {
            var request = Project(source); FirstTextParagraph(request)["style"]!["bullet"] = JsonNode.Parse(replacement);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        Assert.Equal(original, source);
    }

    [Fact]
    public void BulletStartAtRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"bullet":{"type":"number","scheme":"arabicPeriod","startAt":3},
            "defaultText":{"bold":true}},"runs":[{"text":"Retain source numbering"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "0", "-1", "32768", "invalid", "1.5", "1e0", "2147483648", "", "unknown-scheme", "duplicate" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var number = properties.GetFirstChild<A.AutoNumberedBullet>()!;
                if (token == "duplicate") properties.Append(number.CloneNode(true));
                else number.SetAttribute(new OpenXmlAttribute(token == "unknown-scheme" ? "type" : "startAt", "", token));
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["bullet"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["bullet"] = JsonNode.Parse("""{"type":"number","scheme":"arabicPeriod","startAt":1}""");
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var changed = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(changed))["style"]?["bullet"]);
            }
        }
    }

    private static JsonObject FirstBullet(JsonObject program) => FirstTextParagraph(program)["style"]!["bullet"]!.AsObject();

    private static void AssertBulletStartAt(byte[] file, string kind, int? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var number = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.AutoNumberedBullet>()!;
        Assert.Equal("arabicPeriod", number.Type!.InnerText);
        Assert.Equal(expected, number.StartAt?.Value);
        var bullet = FirstBullet(Project(file));
        Assert.Equal("number", bullet["type"]!.GetValue<string>());
        Assert.Equal("arabicPeriod", bullet["scheme"]!.GetValue<string>());
        Assert.Equal(expected, bullet["startAt"]?.GetValue<int>());
    }
}
