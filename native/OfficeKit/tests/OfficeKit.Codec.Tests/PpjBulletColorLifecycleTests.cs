using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    private static readonly (string? Input, string? Projected, uint? Alpha)[] BulletColorChoices =
    [
        ("\"#005500\"", "\"#005500\"", null),
        ("""{"token":"accent3","alpha":0.12345}""", """{"token":"accent3","alpha":0.12345}""", 12345),
        ("true", null, null),
        (null, null, null),
        ("true", null, null),
        ("""{"rgb":"#112233","alpha":0.12345}""", """{"rgb":"#112233","alpha":0.12345}""", 12345),
        ("\"#112233FF\"", "\"#112233FF\"", 100000),
        (null, null, null),
        ("\"#112233\"", "\"#112233\"", null),
    ];

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void BulletColorLifecyclePreservesSource(string kind)
    {
        foreach (var marker in new[] { """{"type":"character","character":"•"}""",
            """{"type":"number","scheme":"arabicPeriod","startAt":3}""",
            """{"type":"picture","uri":"https://example.com/marker.png"}""" })
        {
            var program = BulletFontProgram(kind); FirstTextParagraph(program)["style"]!["bullet"] = JsonNode.Parse(marker);
            var bullet = FirstBullet(program); bullet["color"] = "#112233"; bullet["fontFamily"] = "Georgia"; bullet["size"] = 12;
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
                properties.SetAttribute(new OpenXmlAttribute("marL", "", "+0088900"));
                properties.GetFirstChild<A.BulletSizePoints>()!.SetAttribute(new OpenXmlAttribute("val", "", "+001200"));
                properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                if (properties.GetFirstChild<A.PictureBullet>() is { } picture)
                {
                    var slide = document.PresentationPart!.SlideParts.Single(); var image = slide.AddImagePart(ImagePartType.Png);
                    using var imageStream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j3ioAAAAASUVORK5CYII="));
                    image.FeedData(imageStream);
                    var blip = picture.GetFirstChild<A.Blip>()!; blip.Link = null; blip.Embed = slide.GetIdOfPart(image);
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            var before = source;
            foreach (var choice in BulletColorChoices)
            {
                var authorVariant = program.DeepClone().AsObject(); SetBulletColor(authorVariant, choice.Input);
                AssertBulletColor(PptxCodecTests.RemoveEmbeddedPpj(Compile(authorVariant).File.ToByteArray()), kind, choice.Projected, choice.Alpha, choice.Input == "true");
                var request = Project(before); SetBulletColor(request, choice.Input); var json = request.ToJsonString();
                var changed = Compile(request, before).File.ToByteArray();
                Assert.Equal(json, request.ToJsonString());
                AssertBulletColor(changed, kind, choice.Projected, choice.Alpha, choice.Input == "true");
                AssertOnlyBodyPropertyChanged(before, changed, kind, "paragraph.bullet.color");
                Assert.Equal(changed, Compile(Project(changed), changed).File.ToByteArray());
                before = changed;
            }
            var remove = Project(source); SetBulletColor(remove, null);
            var removed = Compile(remove, source).File.ToByteArray();
            AssertBulletColor(removed, kind, null, null);
            AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.bullet.color");
            foreach (var field in new[] { "color", "colorFollowText" })
            {
                var denied = Project(source); SetBulletColor(denied, "true"); RemoveNumberingAuthority(denied, field);
                Assert.Empty(Compile(denied, source, success: false).File);
            }
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void BulletColorPrecisionThemeGrammarAndValidation()
    {
        var program = BulletFontProgram("text");
        SetBulletColor(program, "\"#112233\"");
        using var stream = new MemoryStream(); stream.Write(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()));
        using (var document = PresentationDocument.Open(stream, true))
        {
            var rgb = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.BulletColor>()!.GetFirstChild<A.RgbColorModelHex>()!;
            rgb.Val = "aabbcc"; var alpha = new A.Alpha(); alpha.SetAttribute(new OpenXmlAttribute("val", "", "+012345")); rgb.Append(alpha);
        }
        var source = stream.ToArray();
        AssertBulletColor(source, "text", """{"rgb":"#AABBCC","alpha":0.12345}""", 12345);
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), "text", "paragraphDefault.bold");
        SetBulletColor(program, FirstBullet(Project(source))["color"]!.ToJsonString());
        AssertBulletColor(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", """{"rgb":"#AABBCC","alpha":0.12345}""", 12345);
        // Tint/shade resolves a declared literal color; it does not evaluate inherited theme paint.
        var grammar = JsonNode.Parse("""{"paint":{"kind":"color","value":"#335577"}}""")!;
        program["design"]!["grammar"]!["tokens"] = grammar.DeepClone();
        foreach (var (input, expected, alpha) in new (string, string, uint?)[] {
            ("""{"rgb":"#112233","alpha":0}""", "\"#11223300\"", 0),
            ("""{"rgb":"#112233","alpha":0.00001}""", """{"rgb":"#112233","alpha":0.00001}""", 1),
            ("""{"rgb":"#112233","alpha":1}""", "\"#112233FF\"", 100000),
            ("""{"rgb":"#112233"}""", "\"#112233\"", null),
            ("""{"token":"accent1"}""", """{"token":"accent1"}""", null),
            ("""{"token":"accent1","alpha":0}""", """{"token":"accent1","alpha":0}""", 0),
            ("""{"token":"accent1","alpha":1}""", """{"token":"accent1","alpha":1}""", 100000),
            ("""{"token":"paint","tint":1,"shade":0.5}""", "\"#808080\"", null) })
        {
            SetBulletColor(program, input);
            AssertBulletColor(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", expected, alpha);
            var request = Project(source); request["design"]!["grammar"]!["tokens"] = grammar.DeepClone(); SetBulletColor(request, input);
            var changed = Compile(request, source).File.ToByteArray();
            AssertBulletColor(changed, "text", expected, alpha);
            AssertOnlyBodyPropertyChanged(source, changed, "text", "paragraph.bullet.color");
        }
        foreach (var request in new[] { program.DeepClone().AsObject(), Project(source) })
        {
            request["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent3":{"kind":"color","value":"#ABCDEF80"}}""");
            SetBulletColor(request, """{"token":"accent3"}""");
            var sourceBytes = request["pages"]![0]!["elements"]![0]!["nativeRef"] is null ? null : source;
            AssertBulletColor(PptxCodecTests.RemoveEmbeddedPpj(Compile(request, sourceBytes).File.ToByteArray()), "text", "\"#ABCDEF80\"", 50196);
            request["design"]!["grammar"]!["tokens"]!["accent3"]!["kind"] = "string";
            var invalidToken = PpjProgramValidator.Validate(System.Text.Encoding.UTF8.GetBytes(request.ToJsonString()));
            Assert.False(invalidToken.IsValid);
            Assert.Contains(invalidToken.Diagnostics, diagnostic => diagnostic.Code == "ppj.grammar.tokenKind");
            Assert.Empty(Compile(request, sourceBytes, success: false).File);
        }
        foreach (var invalid in new[] { "null", "false", "0", "\"bad\"", "{}", """{"rgb":"#11223380"}""",
            """{"rgb":"#112233","alpha":-1}""", """{"rgb":"#112233","alpha":1.1}""", """{"rgb":"#112233","token":"accent1"}""", """{"token":"unknown"}""" })
        {
            SetBulletColor(program, invalid); Assert.Empty(Compile(program, success: false).File);
        }
        SetBulletColor(program, "\"#112233\""); FirstBullet(program)["colorFollowText"] = true;
        Assert.Empty(Compile(program, success: false).File);
        foreach (var invalid in new[] { "false", "null", "0", "\"true\"" })
        {
            SetBulletColor(program, null); FirstBullet(program)["colorFollowText"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void BulletColorRetainsUnmodeledSource()
    {
        var program = BulletFontProgram("text"); SetBulletColor(program, "\"#112233\"");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var token in new[] { "unknown-scheme", "bad-alpha", "overflow-alpha", "negative-alpha", "duplicate", "transformed", "foreign-val", "missing-color", "follow-extra", "invalid-rgb" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var color = properties.GetFirstChild<A.BulletColor>()!; var rgb = color.GetFirstChild<A.RgbColorModelHex>()!;
                if (token == "unknown-scheme") { color.RemoveAllChildren(); var scheme = new A.SchemeColor(); scheme.SetAttribute(new OpenXmlAttribute("val", "", "future")); color.Append(scheme); }
                else if (token.EndsWith("-alpha", StringComparison.Ordinal)) { var alpha = new A.Alpha(); alpha.SetAttribute(new OpenXmlAttribute("val", "", token == "bad-alpha" ? "bad" : token == "overflow-alpha" ? "2147483648" : "-1")); rgb.Append(alpha); }
                else if (token == "duplicate") properties.Append(new A.BulletColorText());
                else if (token == "transformed") rgb.Append(new A.Tint { Val = 20000 });
                else if (token == "foreign-val") { rgb.Val = null; rgb.SetAttribute(new OpenXmlAttribute("future", "val", "urn:officekit:test", "112233")); }
                else if (token == "missing-color") color.RemoveAllChildren();
                else if (token == "invalid-rgb") rgb.Val = "#112233";
                else { color.Remove(); var follow = new A.BulletColorText(); follow.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "yes")); properties.AddChild(follow, true); }
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstBullet(projected)["color"]); Assert.Null(FirstBullet(projected)["colorFollowText"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            foreach (var choice in new[] { "\"#AABBCC\"", "true" })
            {
                var request = Project(source); SetBulletColor(request, choice);
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

    private static void SetBulletColor(JsonObject program, string? choice)
    {
        var bullet = FirstBullet(program); bullet.Remove("color"); bullet.Remove("colorFollowText");
        if (choice == "true") bullet["colorFollowText"] = true;
        else if (choice is not null) bullet["color"] = JsonNode.Parse(choice);
    }

    private static void AssertBulletColor(byte[] bytes, string kind, string? expected, uint? alpha, bool follow = false)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var properties = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!;
        var color = properties.GetFirstChild<A.BulletColor>(); var expectedNode = expected is null ? null : JsonNode.Parse(expected);
        var expectedRgb = expectedNode is JsonValue ? expectedNode.GetValue<string>().TrimStart('#')[..6] : expectedNode?["rgb"]?.GetValue<string>().TrimStart('#');
        Assert.Equal(expectedRgb, color?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant());
        Assert.Equal(expectedNode is JsonObject ? expectedNode["token"]?.GetValue<string>() : null, color?.GetFirstChild<A.SchemeColor>()?.Val?.InnerText);
        Assert.Equal(alpha is null ? (int?)null : (int)alpha, color?.Descendants<A.Alpha>().SingleOrDefault()?.Val?.Value);
        Assert.Equal(follow, properties.GetFirstChild<A.BulletColorText>() is not null);
        var bullet = FirstBullet(Project(bytes));
        Assert.True(JsonNode.DeepEquals(expectedNode, bullet["color"]), bullet["color"]?.ToJsonString());
        Assert.Equal(follow ? true : (bool?)null, bullet["colorFollowText"]?.GetValue<bool>());
    }
}
