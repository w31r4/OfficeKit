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
    public void ParagraphDefaultSoftEdgeLifecyclePreservesSource(string kind)
    {
        foreach (var mixed in new[] { false, true })
        {
            var program = Program(kind);
            program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
                {"paragraphs":[{"style":{"defaultText":{"softEdge":{"radius":2.25}}},
                "runs":[{"text":"Retain run effect","style":{"softEdge":{"radius":3}}}]},
                {"style":{"defaultText":{"softEdge":{"radius":1.234567}}},"runs":[{"text":"Neighbor"}]}]}
                """);
            if (mixed)
            {
                var defaults = FirstTextParagraph(program)["style"]!["defaultText"]!;
                defaults["bold"] = true;
                defaults["glow"] = JsonNode.Parse("""{"color":"#445566","radius":2}""");
                defaults["innerShadow"] = JsonNode.Parse("""{"color":"#112233"}""");
                defaults["shadow"] = JsonNode.Parse("""{"color":{"token":"accent2"},"scaleY":-0.5}""");
                defaults["reflection"] = JsonNode.Parse("""{"startPosition":0.2,"scaleY":-0.5}""");
            }
            var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var defaults = Owner(document, kind).Descendants<A.DefaultRunProperties>().First();
                defaults.Dirty = false;
                if (mixed)
                {
                    var list = defaults.GetFirstChild<A.EffectList>()!;
                    list.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
                    list.InnerXml += """<future:effect xmlns:future="urn:officekit:test" value="retain"/>""";
                }
            }
            var source = stream.ToArray(); var original = source.ToArray();
            Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
            AssertSoftEdge(source, 2.25);
            foreach (var (radius, expected) in new[] { (0d, 0d), (0.0001, 0.000079), (0.005, 0.005039), (0.015, 0.014961), (1000d, 1000d) })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]!["softEdge"]!["radius"] = radius;
                var candidate = Compile(request, source).File.ToByteArray();
                AssertSoftEdge(candidate, expected);
                AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraphDefault.softEdge");
                if (radius == 0)
                {
                    var removeZero = Project(candidate);
                    FirstTextParagraph(removeZero)["style"]!["defaultText"]!.AsObject().Remove("softEdge");
                    var removed = Compile(removeZero, candidate).File.ToByteArray();
                    AssertSoftEdge(removed, null);
                    AssertOnlyBodyPropertyChanged(candidate, removed, kind, "paragraphDefault.softEdge");
                }
            }
            foreach (var mode in mixed ? new[] { "field" } : new[] { "field", "defaultText", "style" })
            {
                var request = Project(source); var paragraph = FirstTextParagraph(request);
                if (mode == "field") paragraph["style"]!["defaultText"]!.AsObject().Remove("softEdge");
                else if (mode == "defaultText") paragraph["style"]!.AsObject().Remove("defaultText");
                else paragraph.Remove("style");
                var removed = Compile(request, source).File.ToByteArray();
                AssertSoftEdge(removed, null);
                AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraphDefault.softEdge");
                var restore = Project(removed); var restoredParagraph = FirstTextParagraph(restore);
                restoredParagraph["style"] ??= new JsonObject();
                restoredParagraph["style"]!["defaultText"] ??= new JsonObject();
                restoredParagraph["style"]!["defaultText"]!["softEdge"] = JsonNode.Parse("""{"radius":2.25}""");
                var restored = Compile(restore, removed).File.ToByteArray();
                AssertSoftEdge(restored, 2.25);
                AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraphDefault.softEdge");
            }
            var denied = Project(source);
            FirstTextParagraph(denied)["style"]!["defaultText"]!["softEdge"]!["radius"] = 0;
            var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
                .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
            fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.defaultText.softEdge"));
            Assert.Empty(Compile(denied, source, success: false).File);
            foreach (var invalid in new[] { "null", "false", "{}", """{"radius":-0.000001}""", """{"radius":1000.000001}""",
                """{"radius":"2"}""", """{"radius":{"token":"softness"}}""", """{"radius":2,"future":true}""" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]!["softEdge"] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            Assert.Equal(original, source);
        }

        void AssertSoftEdge(byte[] file, double? expected)
        {
            JsonNode? value = expected is { } radius ? new JsonObject { ["radius"] = radius } : null;
            AssertDefaultEffectEquals(value, ParagraphDefaultScalar(file, kind, "softEdge"));
            AssertDefaultEffectEquals(value, FirstTextParagraph(Project(file))["style"]?["defaultText"]?["softEdge"]);
        }
    }

    [Fact]
    public void ParagraphDefaultSoftEdgeRetainsUnmodeledSource()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"bold":true,"softEdge":{"radius":2}}},"runs":[{"text":"Retain source soft edge"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var profile in new[] { "missing-radius", "negative-radius", "large-radius", "invalid-radius", "attribute", "child", "duplicate-effect", "duplicate-list", "dag" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var defaults = Owner(document, "text").Descendants<A.DefaultRunProperties>().First();
                var list = defaults.GetFirstChild<A.EffectList>()!;
                var softEdge = list.GetFirstChild<A.SoftEdge>()!;
                if (profile == "missing-radius") softEdge.Radius = null;
                else if (profile == "negative-radius") softEdge.SetAttribute(new OpenXmlAttribute("rad", "", "-1"));
                else if (profile == "large-radius") softEdge.Radius = 12_700_001;
                else if (profile == "invalid-radius") softEdge.SetAttribute(new OpenXmlAttribute("rad", "", "invalid"));
                else if (profile == "attribute") softEdge.SetAttribute(new OpenXmlAttribute("future", "", "retained"));
                else if (profile == "child") list.InnerXml = """<a:softEdge xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rad="25400"><future:content xmlns:future="urn:officekit:test"/></a:softEdge>""";
                else if (profile == "duplicate-effect") list.Append(softEdge.CloneNode(true));
                else if (profile == "duplicate-list") defaults.Append(list.CloneNode(true));
                else defaults.Append(new A.EffectDag());
            }
            var source = stream.ToArray(); var projected = Project(source);
            Assert.True(FirstTextParagraph(projected)["style"]!["defaultText"]!["softEdge"] is null,
                profile + ": " + FirstTextParagraph(projected)["style"]!["defaultText"]!.ToJsonString());
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            FirstTextParagraph(projected)["style"]!["defaultText"]!["softEdge"] = JsonNode.Parse("""{"radius":3}""");
            Assert.Empty(Compile(projected, source, success: false).File);
            foreach (var remove in new[] { false, true })
            {
                var request = Project(source);
                if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
                else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
                var candidate = Compile(request, source).File.ToByteArray();
                AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
                Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["defaultText"]?["softEdge"]);
            }
        }
    }
}
