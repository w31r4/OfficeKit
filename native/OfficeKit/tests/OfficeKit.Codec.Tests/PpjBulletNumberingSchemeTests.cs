using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    private static readonly (string Alias, string Scheme)[] NumberingAliases =
    [
        ("decimal", "arabicPeriod"), ("lower-alpha", "alphaLcPeriod"), ("upper-alpha", "alphaUcPeriod"),
        ("lower-roman", "romanLcPeriod"), ("upper-roman", "romanUcPeriod"),
    ];

    [Fact]
    public void BulletNumberingSchemeCatalogAuthorsAndProjects()
    {
        var schemes = """
            alphaLcParenBoth alphaLcParenR alphaLcPeriod alphaUcParenBoth alphaUcParenR alphaUcPeriod
            arabic1Minus arabic2Minus arabicDbPeriod arabicDbPlain arabicParenBoth arabicParenR arabicPeriod arabicPlain
            circleNumDbPlain circleNumWdBlackPlain circleNumWdWhitePlain ea1ChsPeriod ea1ChsPlain ea1ChtPeriod ea1ChtPlain
            ea1JpnChsDbPeriod ea1JpnKorPeriod ea1JpnKorPlain hebrew2Minus hindiAlpha1Period hindiAlphaPeriod
            hindiNumParenR hindiNumPeriod romanLcParenBoth romanLcParenR romanLcPeriod romanUcParenBoth romanUcParenR romanUcPeriod
            thaiAlphaParenBoth thaiAlphaParenR thaiAlphaPeriod thaiNumParenBoth thaiNumParenR thaiNumPeriod
            """.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(41, schemes.Length);
        var cases = schemes.Select(s => (Field: "scheme", Value: s, Scheme: s))
            .Concat(NumberingAliases.Select(a => (Field: "format", Value: a.Alias, Scheme: a.Scheme))).ToArray();
        var paragraphs = new JsonArray();
        foreach (var (field, value, _) in cases)
            paragraphs.Add(new JsonObject
            {
                ["style"] = new JsonObject { ["bullet"] = new JsonObject { ["type"] = "number", [field] = value, ["startAt"] = 3 } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Catalog entry" }),
            });
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject { ["paragraphs"] = paragraphs };
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var document = PresentationDocument.Open(new MemoryStream(source), false);
        Assert.Equal(cases.Select(c => c.Scheme), Owner(document, "text").Descendants<A.AutoNumberedBullet>().Select(b => b.Type!.InnerText));
        Assert.All(Owner(document, "text").Descendants<A.AutoNumberedBullet>(), b => Assert.Equal(3, b.StartAt!.Value));
        var projected = Project(source);
        var bullets = projected["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]!.AsArray().Select(p => p!["style"]!["bullet"]!).ToArray();
        Assert.Equal(cases.Select(c => c.Scheme), bullets.Select(b => b["scheme"]!.GetValue<string>()));
        Assert.All(bullets, b => { Assert.Null(b["format"]); Assert.Equal(3, b["startAt"]!.GetValue<int>()); });
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
    }

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void BulletNumberingSchemeEditsPreserveSource(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"level":2,"alignment":"thaiDistributed","indent":7,"hanging":3.5,
            "spaceBefore":0,"spaceAfter":9,"lineSpacingMultiplier":1.25,"defaultText":{"bold":true},
            "bullet":{"type":"number","scheme":"arabicPeriod","startAt":3,"fontFamily":"Georgia","sizePercent":1.5,"color":"#112233"}},
            "runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
            {"style":{"bullet":{"type":"number","scheme":"thaiNumParenBoth","startAt":7}},"runs":[{"text":"Neighbor"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
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
        var variants = new[] { (Field: "scheme", Value: "thaiNumParenBoth", Scheme: "thaiNumParenBoth"),
            (Field: "scheme", Value: "romanUcParenBoth", Scheme: "romanUcParenBoth") }
            .Concat(NumberingAliases.Select(a => (Field: "format", Value: a.Alias, Scheme: a.Scheme)));
        foreach (var (field, value, scheme) in variants)
        {
            var request = Project(source); SetNumbering(FirstBullet(request), field, value);
            var requestJson = request.ToJsonString();
            var changed = Compile(request, source).File.ToByteArray();
            Assert.Equal(requestJson, request.ToJsonString());
            AssertNumberingScheme(changed, kind, scheme, 3);
            AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.bullet.scheme");
            if (scheme == "arabicPeriod") Assert.Equal(source, changed);
            var restore = Project(changed); SetNumbering(FirstBullet(restore), "scheme", "arabicPeriod");
            var restored = Compile(restore, changed).File.ToByteArray();
            AssertNumberingScheme(restored, kind, "arabicPeriod", 3);
            AssertOnlyBodyPropertyChanged(changed, restored, kind, "paragraph.bullet.scheme");
            var denied = Project(source); SetNumbering(FirstBullet(denied), field, value); RemoveNumberingAuthority(denied, field);
            Assert.Empty(Compile(denied, source, success: false).File);
        }
        var withoutStart = Project(source); FirstBullet(withoutStart).Remove("startAt");
        var absentSource = Compile(withoutStart, source).File.ToByteArray();
        var absentRequest = Project(absentSource); SetNumbering(FirstBullet(absentRequest), "scheme", "romanUcPeriod");
        var absentChanged = Compile(absentRequest, absentSource).File.ToByteArray();
        AssertNumberingScheme(absentChanged, kind, "romanUcPeriod", null);
        AssertOnlyBodyPropertyChanged(absentSource, absentChanged, kind, "paragraph.bullet.scheme");
        var combined = Project(source); SetNumbering(FirstBullet(combined), "scheme", "romanUcPeriod"); FirstBullet(combined)["startAt"] = 1;
        var both = Compile(combined, source).File.ToByteArray();
        AssertNumberingScheme(both, kind, "romanUcPeriod", 1);
        AssertOnlyBodyPropertyChanged(source, both, kind, "paragraph.bullet.schemeAndStartAt");
        foreach (var field in new[] { "scheme", "startAt" })
        {
            var denied = combined.DeepClone().AsObject(); RemoveNumberingAuthority(denied, field);
            Assert.Empty(Compile(denied, source, success: false).File);
        }
        foreach (var invalid in new[] { "null", "false", "0", "\"\"", "\"futureNumber\"", "\"decimal\"", "{}" })
        {
            var request = Project(source); FirstBullet(request)["scheme"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        foreach (var selection in new[] { "missing", "both", "invalid-alias" })
        {
            var request = Project(source); var bullet = FirstBullet(request);
            if (selection == "both") bullet["format"] = "decimal";
            else { bullet.Remove("scheme"); if (selection == "invalid-alias") bullet["format"] = "future"; }
            Assert.Empty(Compile(request, source, success: false).File);
        }
        Assert.Equal(original, source);
    }

    private static void SetNumbering(JsonObject bullet, string field, string value)
    {
        bullet.Remove(field == "scheme" ? "format" : "scheme"); bullet[field] = value;
    }

    private static void RemoveNumberingAuthority(JsonObject program, string field)
    {
        var fields = program["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.bullet." + field));
    }

    private static void AssertNumberingScheme(byte[] file, string kind, string scheme, int? startAt)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var number = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties!.GetFirstChild<A.AutoNumberedBullet>()!;
        Assert.Equal(scheme, number.Type!.InnerText); Assert.Equal(startAt, number.StartAt?.Value);
        var bullet = FirstBullet(Project(file));
        Assert.Equal(scheme, bullet["scheme"]!.GetValue<string>()); Assert.Null(bullet["format"]);
        Assert.Equal(startAt, bullet["startAt"]?.GetValue<int>());
    }
}
