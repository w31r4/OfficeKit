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
    public void BulletFontLifecyclePreservesSource(string kind)
    {
        foreach (var marker in new[] { """{"type":"character","character":"•"}""",
            """{"type":"number","scheme":"arabicPeriod","startAt":3}""",
            """{"type":"picture","uri":"https://example.com/marker.png"}""" })
        {
            var program = BulletFontProgram(kind); FirstTextParagraph(program)["style"]!["bullet"] = JsonNode.Parse(marker);
            var bullet = FirstBullet(program); bullet["fontFamily"] = "Georgia"; bullet["size"] = 12; bullet["color"] = "#112233";
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertBulletFont(authored, kind, "Georgia");
            var followAuthor = program.DeepClone().AsObject(); SetBulletFont(followAuthor, true);
            AssertBulletFont(PptxCodecTests.RemoveEmbeddedPpj(Compile(followAuthor).File.ToByteArray()), kind, true);
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
                properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
                properties.GetFirstChild<A.BulletSizePoints>()!.SetAttribute(new OpenXmlAttribute("val", "", "+001200"));
                properties.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val = "aabbcc";
                properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                if (properties.GetFirstChild<A.PictureBullet>() is { } picture)
                {
                    // Exercise an imported embedded marker as well as the authored external URI.
                    var slide = document.PresentationPart!.SlideParts.Single();
                    var image = slide.AddImagePart(ImagePartType.Png);
                    using var imageStream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j3ioAAAAASUVORK5CYII="));
                    image.FeedData(imageStream);
                    var blip = picture.GetFirstChild<A.Blip>()!;
                    blip.Link = null; blip.Embed = slide.GetIdOfPart(image);
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            // Each transition starts with candidate bytes and their fresh native projection.
            var before = source;
            foreach (var choice in new object?[] { "宋体 & Symbols", true, null, true, "Georgia", null, "Georgia" })
            {
                var request = Project(before); SetBulletFont(request, choice); var json = request.ToJsonString();
                var changed = Compile(request, before).File.ToByteArray();
                Assert.Equal(json, request.ToJsonString());
                AssertBulletFont(changed, kind, choice);
                AssertOnlyBodyPropertyChanged(before, changed, kind, "paragraph.bullet.font");
                Assert.Equal(changed, Compile(Project(changed), changed).File.ToByteArray());
                before = changed;
            }
            // Test removal directly against the original bytes as well.
            var remove = Project(source); SetBulletFont(remove, null);
            var removed = Compile(remove, source).File.ToByteArray();
            AssertBulletFont(removed, kind, null);
            AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.bullet.font");
            foreach (var field in new[] { "fontFamily", "fontFollowText" })
            {
                var denied = Project(source); SetBulletFont(denied, true); RemoveNumberingAuthority(denied, field);
                Assert.Empty(Compile(denied, source, success: false).File);
            }
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void BulletFontValidationAndNativeFollowTextProjection()
    {
        var program = BulletFontProgram("text");
        foreach (var family in new[] { "Georgia", "宋体 & Symbols", new string('a', 255), string.Concat(Enumerable.Repeat("😀", 255)) })
        {
            SetBulletFont(program, family);
            AssertBulletFont(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", family);
        }
        foreach (var invalid in new[] { "null", "false", "0", "\"\"", "\"   \"", "\"\\u0000\"", "\"\\uFFFF\"", "{}", JsonValue.Create(new string('a', 256))!.ToJsonString() })
        {
            SetBulletFont(program, null); FirstBullet(program)["fontFamily"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
        foreach (var invalid in new[] { "false", "null", "0", "\"true\"" })
        {
            SetBulletFont(program, null); FirstBullet(program)["fontFollowText"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
        SetBulletFont(program, "Georgia"); FirstBullet(program)["fontFollowText"] = true;
        Assert.Empty(Compile(program, success: false).File);
        foreach (var invalid in new[] { "\uD800", "\uDC00", "\0", "\uFFFF" })
            Assert.Throws<CodecException>(() => PptxBulletStyleCodec.Validate(new PresentationTextParagraph { BulletFontFamily = invalid }));
        SetBulletFont(program, "REPLACE-ME");
        foreach (var escaped in new[] { "\\uD800", "\\uDC00" })
            Assert.Empty(Invoke(new CodecRequest { Operation = CodecOperation.CompilePpjToPptx,
                PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString().Replace("REPLACE-ME", escaped, StringComparison.Ordinal)) } }, success: false).File);

        SetBulletFont(program, null);
        using var stream = new MemoryStream(); stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
        using (var document = PresentationDocument.Open(stream, true))
            Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.AddChild(new A.BulletFontText(), true);
        var source = stream.ToArray(); AssertBulletFont(source, "text", true);
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        var changed = Compile(unrelated, source).File.ToByteArray();
        AssertBulletFont(changed, "text", true);
        AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraphDefault.bold");
    }

    [Fact]
    public void BulletFontRetainsUnmodeledSource()
    {
        var program = BulletFontProgram("text"); SetBulletFont(program, "Georgia");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "empty", "missing", "duplicate", "extra-attribute", "foreign-typeface", "follow-extra" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var font = properties.GetFirstChild<A.BulletFont>()!;
                if (token == "empty") font.Typeface = " ";
                else if (token == "missing") font.Typeface = null;
                else if (token == "duplicate") properties.Append(new A.BulletFontText());
                else if (token == "foreign-typeface") { font.Typeface = null; font.SetAttribute(new OpenXmlAttribute("future", "typeface", "urn:officekit:test", "Georgia")); }
                else if (token == "follow-extra") { font.Remove(); var follow = new A.BulletFontText(); follow.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "yes")); properties.AddChild(follow, true); }
                else font.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "yes"));
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstBullet(projected)["fontFamily"]); Assert.Null(FirstBullet(projected)["fontFollowText"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            foreach (var choice in new object[] { "Arial", true })
            {
                var request = Project(source); SetBulletFont(request, choice);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                AssertOnlyBodyPropertyChanged(source, Compile(request, source).File.ToByteArray(), "text", "paragraphDefault.bold");
            }
        }
    }

    private static JsonObject BulletFontProgram(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"level":2,"alignment":"left","indent":7,"hanging":3.5,"spaceBefore":0,"spaceAfter":9,
            "defaultText":{"bold":true},"bullet":{"type":"character","character":"•"}},
            "runs":[{"text":"Keep styled runs","style":{"italic":true,"color":"#334455"}}]},
            {"style":{"bullet":{"type":"character","character":"◆","fontFamily":"Arial"}},"runs":[{"text":"Neighbor"}]}]}
            """);
        return program;
    }

    private static void SetBulletFont(JsonObject program, object? choice)
    {
        var bullet = FirstBullet(program); bullet.Remove("fontFamily"); bullet.Remove("fontFollowText");
        if (choice is string family) bullet["fontFamily"] = family;
        else if (choice is true) bullet["fontFollowText"] = true;
    }

    private static void AssertBulletFont(byte[] bytes, string kind, object? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
        var bullet = FirstBullet(Project(bytes));
        Assert.Equal(expected as string, properties.GetFirstChild<A.BulletFont>()?.Typeface?.Value);
        Assert.Equal(expected is true, properties.GetFirstChild<A.BulletFontText>() is not null);
        Assert.Equal(expected as string, bullet["fontFamily"]?.GetValue<string>());
        Assert.Equal(expected is true ? true : (bool?)null, bullet["fontFollowText"]?.GetValue<bool>());
    }
}
