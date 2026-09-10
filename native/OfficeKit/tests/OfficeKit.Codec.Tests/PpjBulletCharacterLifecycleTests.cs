using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void BulletCharacterLifecyclePreservesSource(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"level":2,"alignment":"left","indent":7,"hanging":3.5,
            "spaceBefore":0,"spaceAfter":9,"lineSpacingMultiplier":1.25,"defaultText":{"bold":true},
            "bullet":{"type":"character","character":"•","fontFamily":"Georgia","size":12,"color":"#112233"}},
            "runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
            {"style":{"bullet":{"type":"character","character":"◆"}},"runs":[{"text":"Neighbor"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            foreach (var paragraph in paragraphs)
                paragraph.ParagraphProperties!.GetFirstChild<A.CharacterBullet>()!
                    .SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
            var properties = paragraphs[0].ParagraphProperties!;
            properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
            properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        foreach (var character in new[] { "★", "😀", "&" })
        {
            var authorVariant = program.DeepClone().AsObject(); FirstBullet(authorVariant)["character"] = character;
            AssertBulletCharacter(PptxCodecTests.RemoveEmbeddedPpj(Compile(authorVariant).File.ToByteArray()), kind, character);
            var request = Project(source); FirstBullet(request)["character"] = character;
            var requestJson = request.ToJsonString();
            var changed = Compile(request, source).File.ToByteArray();
            Assert.Equal(requestJson, request.ToJsonString());
            AssertBulletCharacter(changed, kind, character);
            AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.bullet.character");
            var restore = Project(changed); FirstBullet(restore)["character"] = "•";
            var restored = Compile(restore, changed).File.ToByteArray();
            AssertBulletCharacter(restored, kind, "•");
            AssertOnlyBodyPropertyChanged(changed, restored, kind, "paragraph.bullet.character");
        }
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); FirstBullet(denied)["character"] = "★";
        RemoveNumberingAuthority(denied, "character");
        Assert.Empty(Compile(denied, source, success: false).File);
        foreach (var invalid in new[] { "null", "false", "0", "\"\"", "\"ab\"", "\"e\\u0301\"", "\"\\u0000\"", "\"\\uFFFF\"", "{}" })
        {
            var request = Project(source); FirstBullet(request)["character"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        var missing = Project(source); FirstBullet(missing).Remove("character");
        Assert.Empty(Compile(missing, source, success: false).File);
        foreach (var marker in new[] { """{"type":"none"}""", """{"type":"number","scheme":"arabicPeriod"}""" })
        {
            var request = Project(source); FirstTextParagraph(request)["style"]!["bullet"] = JsonNode.Parse(marker);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        Assert.Equal(original, source);
    }

    [Fact]
    public void BulletCharacterNativeValidationRejectsInvalidScalars()
    {
        foreach (var character in new[] { "•", "😀", "&", "\uFFFD", "\t", "\n", "\r" })
            PptxBulletCodec.Validate(new PresentationTextParagraph { BulletCharacter = character });
        foreach (var character in new[] { "", "ab", "e\u0301", "\uD800", "\uDC00", "\0", "\u000B", "\uFFFE", "\uFFFF" })
            Assert.Throws<CodecException>(() => PptxBulletCodec.Validate(new PresentationTextParagraph { BulletCharacter = character }));
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"bullet":{"type":"character","character":"REPLACE-ME"}},"runs":[{"text":"Scalar validation"}]}]}
            """);
        foreach (var escaped in new[] { "\\uD800", "\\uDC00" })
        {
            // Send the original JSON escape. A JsonNode string round-trip
            // could replace malformed UTF-16 before it reaches the codec.
            var raw = program.ToJsonString().Replace("REPLACE-ME", escaped, StringComparison.Ordinal);
            Assert.Empty(Invoke(new CodecRequest { Operation = CodecOperation.CompilePpjToPptx,
                PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(raw) } }, success: false).File);
        }
    }

    [Fact]
    public void BulletCharacterRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"bullet":{"type":"character","character":"•"},"defaultText":{"bold":true}},
            "runs":[{"text":"Retain source marker"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "", "ab", "e\u0301", "missing", "duplicate" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var marker = properties.GetFirstChild<A.CharacterBullet>()!;
                if (token == "duplicate") properties.Append(marker.CloneNode(true));
                else if (token == "missing") marker.Char = null;
                else marker.Char = token;
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["bullet"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["bullet"] = JsonNode.Parse("""{"type":"character","character":"★"}""");
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

    private static void AssertBulletCharacter(byte[] file, string kind, string expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var character = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.CharacterBullet>();
        Assert.Equal(expected, character!.Char!.Value);
        var bullet = FirstBullet(Project(file));
        Assert.Equal("character", bullet["type"]!.GetValue<string>());
        Assert.Equal(expected, bullet["character"]!.GetValue<string>());
    }
}
