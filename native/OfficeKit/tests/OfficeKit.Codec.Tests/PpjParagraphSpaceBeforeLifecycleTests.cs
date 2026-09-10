using System.Globalization;
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
    public void ParagraphSpaceBeforeLifecyclePreservesSource(string kind)
    {
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{"spaceBefore":7.25},"runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
                {"style":{"spaceBeforeMultiplier":0.12345},"runs":[{"text":"Neighbor"}]}]}
                """);
            if (mixed)
            {
                var style = FirstTextParagraph(program)["style"]!;
                style["spaceAfter"] = 9; style["lineSpacingMultiplier"] = 1.25;
                style["defaultText"] = JsonNode.Parse("""{"italic":true,"shadow":{"color":"#112233","scaleY":-0.5}}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
                var neighbor = paragraphs[1].ParagraphProperties!.GetFirstChild<A.SpaceBefore>()!.FirstChild!;
                neighbor.SetAttribute(new OpenXmlAttribute("val", "", "+012345"));
                if (mixed)
                {
                    var properties = paragraphs[0].ParagraphProperties!;
                    properties.GetFirstChild<A.SpaceAfter>()!.FirstChild!.SetAttribute(new OpenXmlAttribute("val", "", "+0900"));
                    properties.GetFirstChild<A.LineSpacing>()!.FirstChild!.SetAttribute(new OpenXmlAttribute("val", "", "+0125000"));
                    properties.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                    properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            foreach (var (field, value, expected) in new[]
            {
                ("spaceBefore", 0d, 0d), ("spaceBefore", 1.125, 1.12), ("spaceBefore", 1584d, 1584d),
                ("spaceBeforeMultiplier", 0d, 0d), ("spaceBeforeMultiplier", 1.234567, 1.23457), ("spaceBeforeMultiplier", 132d, 132d),
            })
            {
                var request = Project(source); SetBefore(request, field, value);
                var changed = Compile(request, source).File.ToByteArray();
                AssertBefore(changed, kind, field, expected);
                AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph.spaceBefore");
                var remove = Project(changed); SetBefore(remove, null, 0);
                var removed = Compile(remove, changed).File.ToByteArray();
                AssertBefore(removed, kind, null, 0);
                AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph.spaceBefore");
                var restore = Project(removed); SetBefore(restore, field, value);
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertBefore(restored, kind, field, expected);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph.spaceBefore");
                if (field == "spaceBeforeMultiplier")
                {
                    var switchUnit = Project(changed); SetBefore(switchUnit, "spaceBefore", 3);
                    var switched = Compile(switchUnit, changed).File.ToByteArray();
                    AssertBefore(switched, kind, "spaceBefore", 3);
                    AssertOnlyBodyPropertyChanged(changed, switched, kind, "paragraph.spaceBefore");
                }
            }
            if (!mixed)
            {
                var request = Project(source); FirstTextParagraph(request).Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertBefore(removed, kind, null, 0);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.spaceBefore");
            }
            foreach (var deniedField in new[] { "spaceBefore", "spaceBeforeMultiplier" })
            {
                var request = Project(source); SetBefore(request, "spaceBeforeMultiplier", 0.5);
                var fields = request["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                    .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
                fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style." + deniedField));
                Assert.Empty(Compile(request, source, success: false).File);
            }
            foreach (var invalid in new[] { """{"spaceBefore":null}""", """{"spaceBefore":false}""", """{"spaceBefore":-0.0001}""",
                """{"spaceBefore":1584.0001}""", """{"spaceBeforeMultiplier":-0.000001}""", """{"spaceBeforeMultiplier":132.000001}""",
                """{"spaceBefore":{"token":"spacing"}}""", """{"spaceBefore":1,"spaceBeforeMultiplier":1}""" })
            {
                var request = Project(source); SetBefore(request, null, 0);
                foreach (var (key, value) in JsonNode.Parse(invalid)!.AsObject()) FirstTextParagraph(request)["style"]![key] = value?.DeepClone();
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void ParagraphSpaceBeforeStylePrecedenceSelectsUnit()
    {
        var program = Program("text"); var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = JsonNode.Parse("""{"spaceBeforeMultiplier":0.5}""");
        element["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{"spaceBefore":1.125},"runs":[{"text":"Direct points"}]}]}""");
        AssertBefore(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "spaceBefore", 1.12);
        FirstTextParagraph(program).Remove("style");
        element["text"]!["style"] = JsonNode.Parse("""{"paragraph":{"spaceBefore":3}}""");
        AssertBefore(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "spaceBefore", 3);
        FirstTextParagraph(program)["style"] = JsonNode.Parse("""{"spaceBeforeMultiplier":0.75}""");
        AssertBefore(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", "spaceBeforeMultiplier", 0.75);
    }

    [Fact]
    public void ParagraphSpaceBeforeRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"spaceBefore":7,"defaultText":{"bold":true}},"runs":[{"text":"Retain source spacing"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var profile in new[] { "duplicate-slot", "duplicate-child", "mixed-units", "missing-value", "negative", "large-points", "large-multiplier", "invalid", "attribute", "child-attribute", "nested-child" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var before = properties.GetFirstChild<A.SpaceBefore>()!; var points = before.GetFirstChild<A.SpacingPoints>()!;
                if (profile == "duplicate-slot") properties.Append(before.CloneNode(true));
                else if (profile == "duplicate-child") before.Append(points.CloneNode(true));
                else if (profile == "mixed-units") before.Append(new A.SpacingPercent { Val = 50000 });
                else if (profile == "missing-value") points.Val = null;
                else if (profile == "negative") points.Val = -1;
                else if (profile == "large-points") points.Val = 158401;
                else if (profile == "large-multiplier") { points.Remove(); before.Append(new A.SpacingPercent { Val = 13_200_001 }); }
                else if (profile == "invalid") points.SetAttribute(new OpenXmlAttribute("val", "", "invalid"));
                else if (profile == "attribute") before.SetAttribute(new OpenXmlAttribute("future", "", "retained"));
                else if (profile == "child-attribute") points.SetAttribute(new OpenXmlAttribute("future", "", "retained"));
                else before.InnerXml = """<a:spcPts xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="700"><future:content xmlns:future="urn:officekit:test"/></a:spcPts>""";
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.True(FirstTextParagraph(projected)["style"]!["spaceBefore"] is null, profile);
            Assert.True(FirstTextParagraph(projected)["style"]!["spaceBeforeMultiplier"] is null, profile);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetBefore(projected, "spaceBefore", 3);
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var candidate = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["spaceBefore"]);
                Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["spaceBeforeMultiplier"]);
            }
        }
    }

    private static void SetBefore(JsonObject program, string? field, double value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        var style = paragraph["style"]!.AsObject(); style.Remove("spaceBefore"); style.Remove("spaceBeforeMultiplier");
        if (field is not null) style[field] = value;
    }

    private static void AssertBefore(byte[] file, string kind, string? field, double value)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var native = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.GetFirstChild<A.SpaceBefore>();
        var style = FirstTextParagraph(Project(file))["style"];
        if (field is null) { Assert.Null(native); Assert.Null(style?["spaceBefore"]); Assert.Null(style?["spaceBeforeMultiplier"]); return; }
        Assert.NotNull(native);
        var child = Assert.Single(native.ChildElements);
        Assert.Equal(field == "spaceBefore" ? "spcPts" : "spcPct", child.LocalName);
        Assert.Equal(value * (field == "spaceBefore" ? 100 : 100000), double.Parse(child.GetAttribute("val", "").Value, CultureInfo.InvariantCulture), 6);
        Assert.Equal(value, style![field]!.GetValue<double>(), 6);
        Assert.Null(style[field == "spaceBefore" ? "spaceBeforeMultiplier" : "spaceBefore"]);
    }
}
