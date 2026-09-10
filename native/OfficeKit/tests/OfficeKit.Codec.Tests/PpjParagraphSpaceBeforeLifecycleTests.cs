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
    public void ParagraphSpaceBeforeLifecyclePreservesSource(string kind) => CheckParagraphSpacingLifecycle(kind, "spaceBefore");

    private static void CheckParagraphSpacingLifecycle(string kind, string slot)
    {
        var otherSlots = new[] { "spaceBefore", "spaceAfter", "lineSpacing" }.Where(value => value != slot).ToArray();
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{},"runs":[{"text":"Keep runs","style":{"bold":true,"color":"#334455"}}]},
                {"style":{},"runs":[{"text":"Neighbor"}]}]}
                """);
            SetSpacing(slot, program, slot, 7.25);
            program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]![slot + "Multiplier"] = 0.12345;
            if (mixed)
            {
                var style = FirstTextParagraph(program)["style"]!;
                style[otherSlots[0]] = 9; style[otherSlots[1] + "Multiplier"] = 1.25;
                style["defaultText"] = JsonNode.Parse("""{"italic":true,"shadow":{"color":"#112233","scaleY":-0.5}}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
                var neighbor = NativeSpacing(paragraphs[1].ParagraphProperties, slot)!.FirstChild!;
                neighbor.SetAttribute(new OpenXmlAttribute("val", "", "+012345"));
                if (mixed)
                {
                    var properties = paragraphs[0].ParagraphProperties!;
                    NativeSpacing(properties, otherSlots[0])!.FirstChild!.SetAttribute(new OpenXmlAttribute("val", "", "+0900"));
                    NativeSpacing(properties, otherSlots[1])!.FirstChild!.SetAttribute(new OpenXmlAttribute("val", "", "+0125000"));
                    properties.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                    properties.InnerXml += """<future:content xmlns:future="urn:officekit:test" value="retain"/>""";
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            foreach (var (field, value, expected) in new[]
            {
                (slot, slot == "lineSpacing" ? 0.006 : 0d, slot == "lineSpacing" ? 0.01 : 0d),
                (slot, 1.125, 1.12), (slot, 1584d, 1584d),
                (slot + "Multiplier", slot == "lineSpacing" ? 0.000006 : 0d, slot == "lineSpacing" ? 0.00001 : 0d),
                (slot + "Multiplier", 1.234567, 1.23457), (slot + "Multiplier", 132d, 132d),
            })
            {
                var request = Project(source); SetSpacing(slot, request, field, value);
                var changed = Compile(request, source).File.ToByteArray();
                AssertSpacing(slot, changed, kind, field, expected);
                AssertOnlyBodyPropertyChanged(source, changed, kind, "paragraph." + slot);
                var remove = Project(changed); SetSpacing(slot, remove, null, 0);
                var removed = Compile(remove, changed).File.ToByteArray();
                AssertSpacing(slot, removed, kind, null, 0);
                AssertOnlyBodyPropertyChanged(changed, removed, kind, "paragraph." + slot);
                var restore = Project(removed); SetSpacing(slot, restore, field, value);
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertSpacing(slot, restored, kind, field, expected);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraph." + slot);
                if (field == slot + "Multiplier")
                {
                    var switchUnit = Project(changed); SetSpacing(slot, switchUnit, slot, 3);
                    var switched = Compile(switchUnit, changed).File.ToByteArray();
                    AssertSpacing(slot, switched, kind, slot, 3);
                    AssertOnlyBodyPropertyChanged(changed, switched, kind, "paragraph." + slot);
                }
            }
            if (!mixed)
            {
                var request = Project(source); FirstTextParagraph(request).Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertSpacing(slot, removed, kind, null, 0);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph." + slot);
            }
            foreach (var deniedField in new[] { slot, slot + "Multiplier" })
            {
                var request = Project(source); SetSpacing(slot, request, slot + "Multiplier", 0.5);
                var fields = request["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                    .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
                fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style." + deniedField));
                Assert.Empty(Compile(request, source, success: false).File);
            }
            foreach (var invalid in new[] { """{"spaceBefore":null}""", """{"spaceBefore":false}""", """{"spaceBefore":-0.0001}""",
                """{"spaceBefore":1584.0001}""", """{"spaceBeforeMultiplier":-0.000001}""", """{"spaceBeforeMultiplier":132.000001}""",
                """{"spaceBefore":{"token":"spacing"}}""", """{"spaceBefore":1,"spaceBeforeMultiplier":1}""" }
                .Concat(slot == "lineSpacing" ? new[] { """{"spaceBefore":0}""", """{"spaceBeforeMultiplier":0}""",
                    """{"spaceBefore":0.005}""", """{"spaceBeforeMultiplier":0.000005}""" } : Array.Empty<string>()))
            {
                var request = Project(source); SetSpacing(slot, request, null, 0);
                foreach (var (key, value) in JsonNode.Parse(invalid)!.AsObject()) FirstTextParagraph(request)["style"]![key.Replace("spaceBefore", slot, StringComparison.Ordinal)] = value?.DeepClone();
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }
    }

    [Fact]
    public void ParagraphSpaceBeforeStylePrecedenceSelectsUnit() => CheckParagraphSpacingPrecedence("spaceBefore");

    private static void CheckParagraphSpacingPrecedence(string slot)
    {
        var program = Program("text"); var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { [slot + "Multiplier"] = 0.5 };
        element["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{},"runs":[{"text":"Direct points"}]}]}""");
        SetSpacing(slot, program, slot, 1.125);
        AssertSpacing(slot, PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", slot, 1.12);
        FirstTextParagraph(program).Remove("style");
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { [slot] = 3 } };
        AssertSpacing(slot, PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", slot, 3);
        FirstTextParagraph(program)["style"] = new JsonObject { [slot + "Multiplier"] = 0.75 };
        AssertSpacing(slot, PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", slot + "Multiplier", 0.75);
    }

    [Fact]
    public void ParagraphSpaceBeforeRetainsUnmodeledSource() => CheckParagraphSpacingUnmodeledSource("spaceBefore");

    private static void CheckParagraphSpacingUnmodeledSource(string slot)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"bold":true}},"runs":[{"text":"Retain source spacing"}]}]}
            """);
        SetSpacing(slot, program, slot, 7);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var profile in new[] { "duplicate-slot", "duplicate-child", "mixed-units", "missing-value", "negative", "large-points", "large-multiplier", "invalid", "attribute", "child-attribute", "nested-child" }
            .Concat(slot == "lineSpacing" ? new[] { "zero-points", "zero-multiplier" } : Array.Empty<string>()))
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                var before = NativeSpacing(properties, slot)!; var points = before.GetFirstChild<A.SpacingPoints>()!;
                if (profile == "duplicate-slot") properties.Append(before.CloneNode(true));
                else if (profile == "duplicate-child") before.Append(points.CloneNode(true));
                else if (profile == "mixed-units") before.Append(new A.SpacingPercent { Val = 50000 });
                else if (profile == "missing-value") points.Val = null;
                else if (profile == "negative") points.Val = -1;
                else if (profile == "zero-points") points.Val = 0;
                else if (profile == "zero-multiplier") { points.Remove(); before.Append(new A.SpacingPercent { Val = 0 }); }
                else if (profile == "large-points") points.Val = 158401;
                else if (profile == "large-multiplier") { points.Remove(); before.Append(new A.SpacingPercent { Val = 13_200_001 }); }
                else if (profile == "invalid") points.SetAttribute(new OpenXmlAttribute("val", "", "invalid"));
                else if (profile == "attribute") before.SetAttribute(new OpenXmlAttribute("future", "", "retained"));
                else if (profile == "child-attribute") points.SetAttribute(new OpenXmlAttribute("future", "", "retained"));
                else before.InnerXml = """<a:spcPts xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="700"><future:content xmlns:future="urn:officekit:test"/></a:spcPts>""";
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.True(FirstTextParagraph(projected)["style"]![slot] is null, profile);
            Assert.True(FirstTextParagraph(projected)["style"]![slot + "Multiplier"] is null, profile);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetSpacing(slot, projected, slot, 3);
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var candidate = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(candidate))["style"]?[slot]);
                Assert.Null(FirstTextParagraph(Project(candidate))["style"]?[slot + "Multiplier"]);
            }
        }
    }

    private static A.TextSpacingType? NativeSpacing(A.TextParagraphPropertiesType? properties, string slot) => slot switch
    {
        "spaceBefore" => properties?.GetFirstChild<A.SpaceBefore>(),
        "spaceAfter" => properties?.GetFirstChild<A.SpaceAfter>(),
        "lineSpacing" => properties?.GetFirstChild<A.LineSpacing>(),
        _ => throw new ArgumentOutOfRangeException(nameof(slot)),
    };

    private static void SetSpacing(string slot, JsonObject program, string? field, double value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        var style = paragraph["style"]!.AsObject(); style.Remove(slot); style.Remove(slot + "Multiplier");
        if (field is not null) style[field] = value;
    }

    private static void AssertSpacing(string slot, byte[] file, string kind, string? field, double value)
    {
        using var document = PresentationDocument.Open(new MemoryStream(file), false);
        var native = NativeSpacing(Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties, slot);
        var style = FirstTextParagraph(Project(file))["style"];
        if (field is null) { Assert.Null(native); Assert.Null(style?[slot]); Assert.Null(style?[slot + "Multiplier"]); return; }
        Assert.NotNull(native);
        var child = Assert.Single(native.ChildElements);
        Assert.Equal(field == slot ? "spcPts" : "spcPct", child.LocalName);
        Assert.Equal(value * (field == slot ? 100 : 100000), double.Parse(child.GetAttribute("val", "").Value, CultureInfo.InvariantCulture), 6);
        Assert.Equal(value, style![field]!.GetValue<double>(), 6);
        Assert.Null(style[field == slot ? slot + "Multiplier" : slot]);
    }
}
