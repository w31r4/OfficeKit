using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void BulletSizeLifecyclePreservesSource(string kind)
    {
        foreach (var marker in new[] { """{"type":"character","character":"•"}""",
            """{"type":"number","scheme":"arabicPeriod","startAt":3}""",
            """{"type":"picture","uri":"https://example.com/marker.png"}""" })
        {
            var program = BulletFontProgram(kind); FirstTextParagraph(program)["style"]!["bullet"] = JsonNode.Parse(marker);
            var bullet = FirstBullet(program); bullet["fontFamily"] = "Georgia"; bullet["size"] = 12;
            bullet["color"] = JsonNode.Parse("""{"rgb":"#112233","alpha":0.12345}""");
            using var stream = new MemoryStream(); stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
                properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
                properties.GetFirstChild<A.BulletSizePoints>()!.SetAttribute(new OpenXmlAttribute("val", "", "+001200"));
                var rgb = properties.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!;
                rgb.Val = "aabbcc"; rgb.GetFirstChild<A.Alpha>()!.SetAttribute(new OpenXmlAttribute("val", "", "+012345"));
                properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                if (properties.GetFirstChild<A.PictureBullet>() is { } picture)
                {
                    var slide = document.PresentationPart!.SlideParts.Single();
                    var image = slide.AddImagePart(ImagePartType.Png);
                    using var imageStream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j3ioAAAAASUVORK5CYII="));
                    image.FeedData(imageStream);
                    var blip = picture.GetFirstChild<A.Blip>()!; blip.Link = null; blip.Embed = slide.GetIdOfPart(image);
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            var before = source;
            foreach (var choice in new[] { """{"size":14.567}""", """{"sizePercent":1.234567}""",
                """{"sizePercent":1}""", """{"sizeFollowText":true}""", "{}", """{"sizeFollowText":true}""",
                """{"size":12}""", "{}", """{"sizePercent":0.25}""" })
            {
                var request = Project(before); SetBulletSize(request, choice); var json = request.ToJsonString();
                var changed = Compile(request, before).File.ToByteArray();
                AssertBulletSize(changed, kind, choice);
                AssertOnlyBodyPropertyChanged(before, changed, kind, "paragraph.bullet.size");
                Assert.Equal(changed, Compile(Project(changed), changed).File.ToByteArray());
                Assert.Equal(json, request.ToJsonString()); before = changed;
            }
            var remove = Project(source); SetBulletSize(remove, "{}");
            var removed = Compile(remove, source).File.ToByteArray();
            AssertBulletSize(removed, kind, "{}");
            AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.bullet.size");
            foreach (var choice in new[] { """{"sizePercent":1}""", """{"sizeFollowText":true}""", "{}" })
                foreach (var field in new[] { "size" }.Concat(JsonNode.Parse(choice)!.AsObject().Select(p => p.Key)))
                {
                    var denied = Project(source); SetBulletSize(denied, choice); RemoveNumberingAuthority(denied, field);
                    Assert.Empty(Compile(denied, source, success: false).File);
                }
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void BulletSizeBoundsAndPrecision()
    {
        var program = BulletFontProgram("text");
        foreach (var choice in new[] { "{}", """{"size":1}""", """{"size":768}""", """{"size":12.34567}""",
            """{"sizePercent":0.25}""", """{"sizePercent":4}""", """{"sizePercent":1.234567}""", """{"sizeFollowText":true}""" })
        {
            SetBulletSize(program, choice);
            AssertBulletSize(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", choice);
        }
        foreach (var invalid in new[] { """{"size":0.999}""", """{"size":768.001}""", """{"sizePercent":0.24999}""",
            """{"sizePercent":4.00001}""", """{"size":null}""", """{"sizePercent":"1"}""", """{"sizeFollowText":false}""",
            """{"sizeFollowText":null}""", """{"sizeFollowText":"true"}""", """{"size":12,"sizePercent":1}""",
            """{"size":12,"sizeFollowText":true}""", """{"sizePercent":1,"sizeFollowText":true}""",
            """{"size":12,"sizePercent":1,"sizeFollowText":true}""" })
        {
            SetBulletSize(program, invalid); Assert.Empty(Compile(program, success: false).File);
        }
        SetBulletSize(program, "{}");
        using var stream = new MemoryStream(); stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
        using (var document = PresentationDocument.Open(stream, true))
            Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.AddChild(new A.BulletSizeText(), true);
        var source = stream.ToArray(); AssertBulletSize(source, "text", """{"sizeFollowText":true}""");
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
    }

    [Fact]
    public void BulletSizeRetainsUnmodeledSource()
    {
        var program = BulletFontProgram("text"); SetBulletSize(program, """{"size":12}""");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var xml in new[] { "<a:buSzPts/>", "<a:buSzPts val=\"oops\"/>", "<a:buSzPts val=\"999999999999999\"/>",
            "<a:buSzPts val=\"99\"/>", "<a:buSzPts val=\"76801\"/>", "<a:buSzPct val=\"24999\"/>", "<a:buSzPct val=\"400001\"/>",
            "<a:buSzPts val=\"1200\" extra=\"keep\"/>", "<a:buSzTx extra=\"keep\"/>",
            "<a:buSzPts val=\"1200\"/><a:buSzTx/>", "<a:buSzPts future:val=\"1200\" xmlns:future=\"urn:officekit:test\"/>" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                properties.GetFirstChild<A.BulletSizePoints>()!.Remove(); properties.InnerXml += xml;
            }
            var source = stream.ToArray(); var projected = Project(source);
            foreach (var field in new[] { "size", "sizePercent", "sizeFollowText" }) Assert.Null(FirstBullet(projected)[field]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            foreach (var choice in new[] { """{"size":14}""", """{"sizePercent":1}""", """{"sizeFollowText":true}""" })
            {
                var request = Project(source); SetBulletSize(request, choice);
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

    private static void SetBulletSize(JsonObject program, string choice)
    {
        var bullet = FirstBullet(program);
        foreach (var field in new[] { "size", "sizePercent", "sizeFollowText" }) bullet.Remove(field);
        foreach (var property in JsonNode.Parse(choice)!.AsObject()) bullet[property.Key] = property.Value?.DeepClone();
    }

    private static void AssertBulletSize(byte[] bytes, string kind, string choice)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
        var expected = JsonNode.Parse(choice)!.AsObject(); var bullet = FirstBullet(Project(bytes));
        var points = expected["size"] is { } pt ? (int?)Math.Round(pt.GetValue<double>() * 100) : null;
        var percent = expected["sizePercent"] is { } pct ? (int?)Math.Round(pct.GetValue<double>() * 100000) : null;
        Assert.Equal(points, properties.GetFirstChild<A.BulletSizePoints>()?.Val?.Value);
        Assert.Equal(percent, properties.GetFirstChild<A.BulletSizePercentage>()?.Val?.Value);
        Assert.Equal(expected.ContainsKey("sizeFollowText"), properties.GetFirstChild<A.BulletSizeText>() is not null);
        Assert.Equal(points / 100d, bullet["size"]?.GetValue<double>());
        Assert.Equal(percent / 100000d, bullet["sizePercent"]?.GetValue<double>());
        Assert.Equal(expected["sizeFollowText"]?.GetValue<bool>(), bullet["sizeFollowText"]?.GetValue<bool>());
    }
}
