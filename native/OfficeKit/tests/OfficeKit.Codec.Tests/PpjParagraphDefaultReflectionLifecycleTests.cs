using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    private const string DefaultReflectionJson = """
        {"blur":2,"distance":3,"angle":45,"startOpacity":0.4,"endOpacity":0.1,
        "startPosition":0.2,"endPosition":0.8,"fadeAngle":90,"scaleX":-1,"scaleY":0.5,
        "skewX":-12,"skewY":8,"alignment":"ctr","rotateWithShape":true}
        """;

    private static (string Input, string Expected)[] DefaultReflectionCases() => new[]
    {
        (DefaultReflectionJson, DefaultReflectionJson),
        ("{}", "{}"),
        ("""{"blur":1000,"distance":100000,"angle":-90,"startOpacity":1,"endOpacity":0,"startPosition":0,"endPosition":1}""",
         """{"blur":1000,"distance":100000,"angle":270,"startOpacity":1,"endOpacity":0,"startPosition":0,"endPosition":1}"""),
        ("""{"scaleX":-21474.83648,"scaleY":21474.83647}""", """{"scaleX":-21474.83648,"scaleY":21474.83647}"""),
        ("""{"blur":0.0001,"distance":1.23456,"angle":359.999999,"fadeAngle":-360,"startOpacity":{"token":"fade"},"endOpacity":0.123456,"startPosition":0.234567,"endPosition":0.765432,"scaleX":1.234567,"scaleY":-0.000004,"skewX":-1.123456,"skewY":1.123456,"rotateWithShape":false}""",
         """{"blur":0.000079,"distance":1.234567,"angle":0,"fadeAngle":0,"startOpacity":0.25,"endOpacity":0.12346,"startPosition":0.23457,"endPosition":0.76543,"scaleX":1.23457,"scaleY":0,"skewX":-1.12345,"skewY":1.12345,"rotateWithShape":false}"""),
    };

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultReflectionOptionalFieldsPreserveMixedSource(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"bold":true,
            "glow":{"color":"#445566","radius":2},"innerShadow":{"color":"#112233"},
            "shadow":{"color":"#334455","blur":2,"distance":3,"angle":90},"softEdge":{"radius":2}}},
            "runs":[{"text":"Preserve reflection state","style":{"reflection":{"blur":3,"distance":4,"angle":90,"startOpacity":0.4,"endOpacity":0.1}}}]},
            {"style":{"defaultText":{"reflection":{"startPosition":0.12345,"scaleY":-0.5,"rotateWithShape":false}}},"runs":[{"text":"Neighbor"}]}]}
            """);
        FirstTextParagraph(program)["style"]!["defaultText"]!["reflection"] = JsonNode.Parse(DefaultReflectionJson);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var list = Owner(document, kind).Descendants<A.DefaultRunProperties>().First().GetFirstChild<A.EffectList>()!;
            list.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
            list.InnerXml += """<future:effect xmlns:future="urn:officekit:test" value="retain"/>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var baseline = Project(source);
        Assert.Equal(source, Compile(baseline, source).File.ToByteArray());
        AssertDefaultEffectEquals(JsonNode.Parse(DefaultReflectionJson), ParagraphDefaultScalar(source, kind, "reflection"));
        var before = FirstTextParagraph(baseline)["style"]!["defaultText"]!;
        foreach (var field in JsonNode.Parse(DefaultReflectionJson)!.AsObject().Select(pair => pair.Key).Append("reflection"))
        {
            var request = Project(source);
            var defaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
            if (field == "reflection") defaults.Remove(field); else defaults["reflection"]!.AsObject().Remove(field);
            var expected = defaults["reflection"]?.DeepClone();
            var candidate = Compile(request, source).File.ToByteArray();
            var after = FirstTextParagraph(Project(candidate))["style"]!["defaultText"]!;
            AssertDefaultEffectEquals(expected, ParagraphDefaultScalar(candidate, kind, "reflection"));
            AssertDefaultEffectEquals(expected, after["reflection"]);
            foreach (var sibling in new[] { "glow", "innerShadow", "shadow", "softEdge", "bold" })
                Assert.True(JsonNode.DeepEquals(before[sibling], after[sibling]), sibling);
            AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraphDefault.reflection");

            var restore = Project(candidate);
            if (field == "reflection") FirstTextParagraph(restore)["style"]!["defaultText"]!["reflection"] = expected = new JsonObject();
            else
            {
                JsonNode zero = field == "alignment" ? JsonValue.Create("tl")! : field == "rotateWithShape" ? JsonValue.Create(false)! : JsonValue.Create(0d)!;
                FirstTextParagraph(restore)["style"]!["defaultText"]!["reflection"]![field] = zero.DeepClone();
                expected![field] = zero;
            }
            var restored = Compile(restore, candidate).File.ToByteArray();
            AssertDefaultEffectEquals(expected, ParagraphDefaultScalar(restored, kind, "reflection"));
            AssertDefaultEffectEquals(expected, FirstTextParagraph(Project(restored))["style"]!["defaultText"]!["reflection"]);
            AssertOnlyBodyPropertyChanged(candidate, restored, kind, "paragraphDefault.reflection");
        }

        foreach (var invalid in new[] { "null", "false", """{"blur":-1}""", """{"blur":1000.001}""", """{"distance":100000.001}""",
            """{"angle":360.001}""", """{"fadeAngle":-360.001}""", """{"startPosition":-0.1}""", """{"endPosition":1.1}""",
            """{"scaleX":-21474.83649}""", """{"scaleY":21474.83648}""", """{"skewX":90}""", """{"skewY":-89.999999}""",
            """{"alignment":"middle"}""", """{"rotateWithShape":0}""", """{"startOpacity":-0.1}""", """{"endOpacity":1.1}""",
            """{"startOpacity":{"token":"missing"}}""", """{"endOpacity":{"token":"paint"}}""", """{"future":true}""" })
        {
            var request = Project(source);
            request["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"paint":{"kind":"color","value":"#112233"}}""");
            FirstTextParagraph(request)["style"]!["defaultText"]!["reflection"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        var invalidRun = program.DeepClone().AsObject();
        FirstTextParagraph(invalidRun)["runs"]![0]!["style"]!["reflection"] = new JsonObject();
        Assert.Empty(Compile(invalidRun, success: false).File);
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("duplicate-reflection")]
    [InlineData("duplicate-list")]
    [InlineData("dag")]
    [InlineData("blur")]
    [InlineData("skew")]
    [InlineData("alignment")]
    [InlineData("attribute")]
    [InlineData("child")]
    public void ParagraphDefaultReflectionRetainsUnmodeledSource(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{"defaultText":{"bold":true,"reflection":{}}},"runs":[{"text":"Retain source reflection"}]}]}""");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(document, "text").Descendants<A.DefaultRunProperties>().First();
            var list = defaults.GetFirstChild<A.EffectList>()!;
            var reflection = list.GetFirstChild<A.Reflection>()!;
            if (profile == "duplicate-reflection") list.Append(reflection.CloneNode(true));
            else if (profile == "duplicate-list") defaults.Append(list.CloneNode(true));
            else if (profile == "dag") defaults.Append(new A.EffectDag());
            else if (profile == "blur") reflection.BlurRadius = -1;
            else if (profile == "skew") reflection.HorizontalSkew = 5_400_000;
            else if (profile == "alignment") reflection.SetAttribute(new OpenXmlAttribute("algn", "", "unknown"));
            else if (profile == "attribute") reflection.SetAttribute(new OpenXmlAttribute("future", "", "retain"));
            else list.InnerXml = """<a:reflection xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><future:content xmlns:future="urn:officekit:test" value="retain"/></a:reflection>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]!["defaultText"]!["reflection"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["reflection"] = JsonNode.Parse(DefaultReflectionJson);
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var request = Project(source);
            if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
            else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(request, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["defaultText"]?["reflection"]);
        }
        Assert.Equal(original, source);
    }

    private static JsonNode? NativeDefaultReflection(A.DefaultRunProperties? defaults)
    {
        var native = defaults?.GetFirstChild<A.EffectList>()?.GetFirstChild<A.Reflection>();
        if (native is null) return null;
        var output = new JsonObject();
        foreach (var (field, attribute, scale) in new[]
        {
            ("blur", "blurRad", 12700d), ("distance", "dist", 12700d), ("angle", "dir", 60000d),
            ("startOpacity", "stA", 100000d), ("endOpacity", "endA", 100000d),
            ("startPosition", "stPos", 100000d), ("endPosition", "endPos", 100000d),
            ("fadeAngle", "fadeDir", 60000d), ("scaleX", "sx", 100000d), ("scaleY", "sy", 100000d),
            ("skewX", "kx", 60000d), ("skewY", "ky", 60000d),
        })
        {
            var raw = native.GetAttributes().FirstOrDefault(item => item.LocalName == attribute && item.NamespaceUri.Length == 0).Value;
            if (string.IsNullOrEmpty(raw)) continue;
            var value = double.Parse(raw, CultureInfo.InvariantCulture) / scale;
            output[field] = scale == 12700 ? Math.Round(value, 6, MidpointRounding.AwayFromZero) : value;
        }
        if (native.Alignment is { } alignment) output["alignment"] = alignment.InnerText;
        if (native.RotateWithShape is { } rotate) output["rotateWithShape"] = rotate.Value;
        return output;
    }
}
