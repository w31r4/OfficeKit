using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    private const string DefaultShadowJson = """{"color":"#112233"}""";
    private const string FullDefaultShadowJson = """
        {"color":"#112233","blur":2,"distance":3,"angle":45,"opacity":0.5,
        "scaleX":-1,"scaleY":0.5,"skewX":-12,"skewY":8,"alignment":"ctr","rotateWithShape":true}
        """;

    private static (string Input, string Expected)[] DefaultShadowCases() => new[]
    {
        ("""{"color":"#FF0000","blur":0,"distance":0,"angle":0,"opacity":0,"scaleX":0,"scaleY":0,"skewX":0,"skewY":0,"alignment":"tl","rotateWithShape":false}""",
         """{"color":"#FF0000","blur":0,"distance":0,"angle":0,"opacity":0,"scaleX":0,"scaleY":0,"skewX":0,"skewY":0,"alignment":"tl","rotateWithShape":false}"""),
        ("""{"color":"#000000FF","blur":1000,"distance":100000,"angle":-90}""",
         """{"color":"#000000","blur":1000,"distance":100000,"angle":270,"opacity":1}"""),
        ("""{"color":{"token":"accent1","tint":0.5},"blur":0.0001,"distance":1.23456,"angle":359.999999,"opacity":{"token":"fade"},"scaleX":1.234567,"scaleY":-0.000004,"skewX":-1.123456,"skewY":1.123456}""",
         """{"color":"#808080","blur":0.000079,"distance":1.234567,"angle":0,"opacity":0.25,"scaleX":1.23457,"scaleY":0,"skewX":-1.12345,"skewY":1.12345}"""),
        ("""{"color":"#11223380"}""", """{"color":"#112233","opacity":0.50196}"""),
        ("""{"color":"#112233","scaleX":-21474.83648,"scaleY":21474.83647}""", """{"color":"#112233","scaleX":-21474.83648,"scaleY":21474.83647}"""),
    };

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultShadowOptionalFieldsPreserveMixedSource(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"bold":true,
            "glow":{"color":"#445566","radius":2},"innerShadow":{"color":"#112233"},
            "reflection":{"startPosition":0.2,"scaleY":-0.5},"softEdge":{"radius":2}}},
            "runs":[{"text":"Preserve shadow state","style":{"shadow":{"color":"#778899","blur":3,"distance":4,"angle":90}}}]},
            {"style":{"defaultText":{"shadow":{"color":{"token":"accent2"},"opacity":0.12345,"scaleX":-0.5,"rotateWithShape":false}}},"runs":[{"text":"Neighbor"}]}]}
            """);
        FirstTextParagraph(program)["style"]!["defaultText"]!["shadow"] = JsonNode.Parse(FullDefaultShadowJson);
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
        AssertDefaultEffectEquals(JsonNode.Parse(FullDefaultShadowJson), ParagraphDefaultScalar(source, kind, "shadow"));
        var before = FirstTextParagraph(baseline)["style"]!["defaultText"]!;
        foreach (var field in JsonNode.Parse(FullDefaultShadowJson)!.AsObject().Select(pair => pair.Key).Where(field => field != "color"))
        {
            var request = Project(source);
            var shadow = FirstTextParagraph(request)["style"]!["defaultText"]!["shadow"]!.AsObject();
            shadow.Remove(field); var expected = shadow.DeepClone();
            var candidate = Compile(request, source).File.ToByteArray();
            var after = FirstTextParagraph(Project(candidate))["style"]!["defaultText"]!;
            AssertDefaultEffectEquals(expected, ParagraphDefaultScalar(candidate, kind, "shadow"));
            AssertDefaultEffectEquals(expected, after["shadow"]);
            foreach (var sibling in new[] { "glow", "innerShadow", "reflection", "softEdge", "bold" })
                Assert.True(JsonNode.DeepEquals(before[sibling], after[sibling]), sibling);
            AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraphDefault.shadow");

            var restore = Project(candidate);
            JsonNode zero = field == "alignment" ? JsonValue.Create("tl")! : field == "rotateWithShape" ? JsonValue.Create(false)! : JsonValue.Create(0d)!;
            FirstTextParagraph(restore)["style"]!["defaultText"]!["shadow"]![field] = zero.DeepClone(); expected[field] = zero;
            var restored = Compile(restore, candidate).File.ToByteArray();
            AssertDefaultEffectEquals(expected, ParagraphDefaultScalar(restored, kind, "shadow"));
            AssertDefaultEffectEquals(expected, FirstTextParagraph(Project(restored))["style"]!["defaultText"]!["shadow"]);
            AssertOnlyBodyPropertyChanged(candidate, restored, kind, "paragraphDefault.shadow");
        }
        foreach (var invalid in new[] { "null", "false", "{}", """{"color":"#112233","blur":-1}""", """{"color":"#112233","blur":1000.001}""",
            """{"color":"#112233","distance":100000.001}""", """{"color":"#112233","angle":360.001}""",
            """{"color":"#112233","scaleX":-21474.83649}""", """{"color":"#112233","scaleY":21474.83648}""",
            """{"color":"#112233","skewX":90}""", """{"color":"#112233","skewY":-89.999999}""",
            """{"color":"#112233","alignment":"middle"}""", """{"color":"#112233","rotateWithShape":0}""",
            """{"color":"#112233","opacity":-0.1}""", """{"color":"#112233","opacity":1.1}""",
            """{"color":{"token":"missing"}}""", """{"color":{"token":"fade"}}""", """{"color":"#112233","opacity":{"token":"paint"}}""",
            """{"color":"#112233","future":true}""" })
        {
            var request = Project(source);
            request["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"paint":{"kind":"color","value":"#112233"},"fade":{"kind":"opacity","value":0.5}}""");
            FirstTextParagraph(request)["style"]!["defaultText"]!["shadow"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(request, source, success: false).File);
        }
        var invalidRun = program.DeepClone().AsObject();
        FirstTextParagraph(invalidRun)["runs"]![0]!["style"]!["shadow"] = JsonNode.Parse(DefaultShadowJson);
        Assert.Empty(Compile(invalidRun, success: false).File);
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("duplicate-shadow")]
    [InlineData("duplicate-list")]
    [InlineData("dag")]
    [InlineData("blur")]
    [InlineData("large-distance")]
    [InlineData("skew")]
    [InlineData("scale-invalid")]
    [InlineData("scale-overflow")]
    [InlineData("scale-y-invalid")]
    [InlineData("scale-y-overflow")]
    [InlineData("attribute")]
    [InlineData("color")]
    [InlineData("alpha-child")]
    public void ParagraphDefaultShadowRetainsUnmodeledSource(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""{"paragraphs":[{"style":{"defaultText":{"bold":true,"shadow":{"color":"#112233"}}},"runs":[{"text":"Retain source shadow"}]}]}""");
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(document, "text").Descendants<A.DefaultRunProperties>().First();
            var list = defaults.GetFirstChild<A.EffectList>()!;
            var shadow = list.GetFirstChild<A.OuterShadow>()!;
            if (profile == "duplicate-shadow") list.Append(shadow.CloneNode(true));
            else if (profile == "duplicate-list") defaults.Append(list.CloneNode(true));
            else if (profile == "dag") defaults.Append(new A.EffectDag());
            else if (profile == "blur") shadow.BlurRadius = -1;
            else if (profile == "large-distance") shadow.Distance = 1_270_000_001;
            else if (profile == "skew") shadow.HorizontalSkew = 5_400_000;
            else if (profile == "scale-invalid") shadow.SetAttribute(new OpenXmlAttribute("sx", "", "not-an-int"));
            else if (profile == "scale-overflow") shadow.SetAttribute(new OpenXmlAttribute("sx", "", "2147483648"));
            else if (profile == "scale-y-invalid") shadow.SetAttribute(new OpenXmlAttribute("sy", "", "not-an-int"));
            else if (profile == "scale-y-overflow") shadow.SetAttribute(new OpenXmlAttribute("sy", "", "2147483648"));
            else if (profile == "attribute") shadow.SetAttribute(new OpenXmlAttribute("future", "", "retain"));
            else if (profile == "color") shadow.GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Tint { Val = 50000 });
            else shadow.GetFirstChild<A.RgbColorModelHex>()!.InnerXml = """<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="50000"><future:content xmlns:future="urn:officekit:test"/></a:alpha>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]!["defaultText"]!["shadow"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["shadow"] = JsonNode.Parse(DefaultShadowJson);
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var request = Project(source);
            if (remove) FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("bold");
            else FirstTextParagraph(request)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(request, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["defaultText"]?["shadow"]);
        }
        Assert.Equal(original, source);
    }

    private static JsonNode? NativeDefaultShadow(A.DefaultRunProperties? defaults)
    {
        var native = defaults?.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>();
        if (native is null) return null;
        var color = native.FirstChild!;
        var output = new JsonObject
        {
            ["color"] = color is A.SchemeColor scheme ? new JsonObject { ["token"] = scheme.Val!.InnerText }
                : JsonValue.Create("#" + ((A.RgbColorModelHex)color).Val!.InnerText),
        };
        if (color.GetFirstChild<A.Alpha>() is { } alpha) output["opacity"] = alpha.Val!.Value / 100000d;
        foreach (var (field, attribute, scale) in new[]
        {
            ("blur", "blurRad", 12700d), ("distance", "dist", 12700d), ("angle", "dir", 60000d),
            ("scaleX", "sx", 100000d), ("scaleY", "sy", 100000d),
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
