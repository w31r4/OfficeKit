using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text", false, "bold")]
    [InlineData("text", false, "italic")]
    [InlineData("text", true, "bold")]
    [InlineData("text", true, "italic")]
    [InlineData("shape", false, "bold")]
    [InlineData("shape", false, "italic")]
    [InlineData("shape", true, "bold")]
    [InlineData("shape", true, "italic")]
    [InlineData("text", false, "size")]
    [InlineData("text", true, "size")]
    [InlineData("shape", false, "size")]
    [InlineData("shape", true, "size")]
    [InlineData("text", false, "fontFamily")]
    [InlineData("text", true, "fontFamily")]
    [InlineData("shape", false, "fontFamily")]
    [InlineData("shape", true, "fontFamily")]
    [InlineData("text", false, "fontFamilyEastAsia")]
    [InlineData("text", true, "fontFamilyEastAsia")]
    [InlineData("shape", false, "fontFamilyEastAsia")]
    [InlineData("shape", true, "fontFamilyEastAsia")]
    [InlineData("text", false, "fontFamilyComplexScript")]
    [InlineData("text", true, "fontFamilyComplexScript")]
    [InlineData("shape", false, "fontFamilyComplexScript")]
    [InlineData("shape", true, "fontFamilyComplexScript")]
    [InlineData("text", false, "language")]
    [InlineData("text", true, "language")]
    [InlineData("shape", false, "language")]
    [InlineData("shape", true, "language")]
    [InlineData("text", false, "kerning")]
    [InlineData("text", true, "kerning")]
    [InlineData("shape", false, "kerning")]
    [InlineData("shape", true, "kerning")]
    [InlineData("text", false, "letterSpacing")]
    [InlineData("text", true, "letterSpacing")]
    [InlineData("shape", false, "letterSpacing")]
    [InlineData("shape", true, "letterSpacing")]
    [InlineData("text", false, "baseline")]
    [InlineData("text", true, "baseline")]
    [InlineData("shape", false, "baseline")]
    [InlineData("shape", true, "baseline")]
    [InlineData("text", false, "capitalization")]
    [InlineData("text", true, "capitalization")]
    [InlineData("shape", false, "capitalization")]
    [InlineData("shape", true, "capitalization")]
    [InlineData("text", false, "strike")]
    [InlineData("text", true, "strike")]
    [InlineData("shape", false, "strike")]
    [InlineData("shape", true, "strike")]
    [InlineData("text", false, "underline")]
    [InlineData("text", true, "underline")]
    [InlineData("shape", false, "underline")]
    [InlineData("shape", true, "underline")]
    [InlineData("text", false, "highlight")]
    [InlineData("text", true, "highlight")]
    [InlineData("shape", false, "highlight")]
    [InlineData("shape", true, "highlight")]
    [InlineData("text", false, "color")]
    [InlineData("text", true, "color")]
    [InlineData("shape", false, "color")]
    [InlineData("shape", true, "color")]
    [InlineData("text", false, "gradient")]
    [InlineData("text", true, "gradient")]
    [InlineData("shape", false, "gradient")]
    [InlineData("shape", true, "gradient")]
    public void ParagraphDefaultScalarPresencePreservesOtherState(string kind, bool otherDefaults, string field)
    {
        var program = Program(kind);
        var isFont = field is "fontFamily" or "fontFamilyEastAsia" or "fontFamilyComplexScript";
        JsonNode? initial = field switch
        {
            "size" => JsonValue.Create(18.25),
            "kerning" => JsonValue.Create(12.25),
            "letterSpacing" => JsonValue.Create(-1.25),
            "baseline" => JsonValue.Create(-12.5),
            "capitalization" => JsonValue.Create("all"),
            "strike" => JsonValue.Create("sngStrike"),
            "underline" => JsonValue.Create("sng"),
            "highlight" => JsonValue.Create("#FFFF00"),
            "color" => JsonValue.Create("#224466"),
            "gradient" => JsonNode.Parse(DefaultGradientJson),
            "language" => JsonValue.Create("en-US"),
            "fontFamily" => JsonValue.Create("Arial"),
            "fontFamilyEastAsia" => JsonValue.Create("SimSun"),
            "fontFamilyComplexScript" => JsonValue.Create("Arial"),
            _ => JsonValue.Create(true),
        };
        var initialJson = initial!.ToJsonString();
        var defaults = new JsonObject { [field] = initial };
        if (otherDefaults)
        {
            defaults[field == "bold" ? "italic" : "bold"] = true; if (field != "size") defaults["size"] = 18; defaults["fontFamily"] = "Arial";
            if (isFont)
            {
                defaults["fontFamilyEastAsia"] = "SimSun";
                defaults["fontFamilyComplexScript"] = "Arial";
            }
            defaults["softEdge"] = new JsonObject { ["radius"] = 2 };
        }
        var paragraphStyle = new JsonObject { ["defaultText"] = defaults };
        if (otherDefaults) { paragraphStyle["alignment"] = "center"; paragraphStyle["spaceBefore"] = 7; }
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(
                new JsonObject { ["style"] = paragraphStyle, ["runs"] = new JsonArray(
                    new JsonObject { ["text"] = "Retain ", ["style"] = new JsonObject { ["bold"] = false, ["fontFamily"] = "Courier New", ["language"] = "de-DE", ["kerning"] = 8, ["letterSpacing"] = 2, ["baseline"] = 5, ["capitalization"] = "none", ["strike"] = "noStrike", ["underline"] = "none", ["highlight"] = "#00FF00", ["color"] = "#33445580" } },
                    new JsonObject { ["text"] = "these runs", ["style"] = new JsonObject { ["italic"] = false } }) },
                new JsonObject { ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["bold"] = false } },
                    ["runs"] = new JsonArray(new JsonObject { ["text"] = "Second paragraph" }) })
        };
        if (field == "gradient")
            program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["defaultText"]!["gradient"] = initial.DeepClone();
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var nativeDefaults = Owner(doc, kind).Descendants<A.DefaultRunProperties>().First();
            nativeDefaults.Dirty = false;
            if (field == "gradient")
            {
                var neighbor = Owner(doc, kind).Descendants<A.DefaultRunProperties>().Skip(1).First().GetFirstChild<A.GradientFill>()!;
                neighbor.Append(new A.TileRectangle());
                neighbor.GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>().First()
                    .GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Alpha { Val = 12345 });
            }
            if (field == "language") nativeDefaults.SetAttribute(new OpenXmlAttribute("altLang", "", "ja-JP"));
            // Authored fontFamily also supplies an East Asian fallback. Make
            // this source-only fixture genuinely Latin-only for wrapper removal.
            if (field == "fontFamily" && !otherDefaults)
                nativeDefaults.GetFirstChild<A.EastAsianFont>()?.Remove();
            if (otherDefaults) nativeDefaults.SetAttribute(new OpenXmlAttribute(field == "bold" ? "i" : "b", "", "true"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        if (field == "gradient") Assert.True(JsonNode.DeepEquals(initial, ParagraphDefaultScalar(source, kind, field)));
        else Assert.Equal(initialJson, ParagraphDefaultScalar(source, kind, field)?.ToJsonString());
        foreach (var mode in otherDefaults ? new[] { "field" } : new[] { "field", "defaultText", "style" })
        {
            var request = Project(source); var paragraph = FirstTextParagraph(request);
            if (mode == "field") paragraph["style"]!["defaultText"]!.AsObject().Remove(field);
            else if (mode == "defaultText") paragraph["style"]!.AsObject().Remove("defaultText");
            else paragraph.Remove("style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(ParagraphDefaultScalar(deleted, kind, field));
            Assert.Null(FirstTextParagraph(Project(deleted))["style"]?["defaultText"]?[field]);
            AssertOnlyBodyPropertyChanged(source, deleted, kind, "paragraphDefault." + field);
            foreach (var value in field switch
            {
                "gradient" => DefaultGradientCases().Select(c => c.Input).ToArray(),
                "size" => new[] { "1", "18.25", "18.256", "18.125", "768" },
                "kerning" => new[] { "0", "0.001", "0.01", "12.25", "12.256", "12.125", "768" },
                "letterSpacing" => new[] { "-768", "-1.256", "-1.125", "-0.001", "0", "0.01", "1.125", "1.256", "768" },
                "baseline" => new[] { "-400", "-12.3456", "-12.3445", "-0.0001", "0", "0.001", "12.3445", "12.3456", "400" },
                "color" => new[] { "\"#000000\"", "\"#ff00ff\"", "\"#11223300\"", "\"#11223380\"", "\"#112233ff\"", """{"token":"mark","tint":1,"shade":0.5}""" },
                "highlight" => new[] { "\"#000000\"", "\"#ff00ff\"", "\"#ffffffff\"", """{"token":"mark","tint":1,"shade":0.5}""" },
                "underline" => new[] { "none", "single", "double", "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong", "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl" }
                    .Select(underline => JsonValue.Create(underline)!.ToJsonString()).ToArray(),
                "strike" => new[] { "false", "true", "\"noStrike\"", "\"sngStrike\"", "\"dblStrike\"" },
                "capitalization" => new[] { "none", "small", "all" }.Select(capitalization => JsonValue.Create(capitalization)!.ToJsonString()).ToArray(),
                "language" => new[] { "fr-FR", "zh-Hant-TW", "EN-us", "en-" + string.Join("-", Enumerable.Repeat("abcdefgh", 6)) + "-abcdef" }
                    .Select(language => JsonValue.Create(language)!.ToJsonString()).ToArray(),
                "fontFamilyComplexScript" => new[] { "Amiri", "+mn-cs", new string('F', 255) }
                    .Select(font => JsonValue.Create(font)!.ToJsonString()).ToArray(),
                "fontFamilyEastAsia" => new[] { "MS Gothic", "+mn-ea", new string('F', 255) }
                    .Select(font => JsonValue.Create(font)!.ToJsonString()).ToArray(),
                "fontFamily" => new[] { "Georgia", "+mn-lt", new string('F', 255) }
                    .Select(font => JsonValue.Create(font)!.ToJsonString()).ToArray(),
                _ => new[] { "false", "true" },
            })
            {
                var restore = Project(deleted); var restoredParagraph = FirstTextParagraph(restore);
                if (field is "highlight" or "color" or "gradient")
                    restore["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"mark":{"kind":"color","value":"#000000"},"accent1":{"kind":"color","value":"#000000"}}""");
                restoredParagraph["style"] ??= new JsonObject();
                restoredParagraph["style"]!["defaultText"] ??= new JsonObject();
                restoredParagraph["style"]!["defaultText"]![field] = JsonNode.Parse(value);
                var restored = Compile(restore, deleted).File.ToByteArray();
                var precision = field == "baseline" ? 1000d : 100d;
                var expected = field is "size" or "kerning" or "letterSpacing" or "baseline"
                    ? JsonValue.Create(checked((int)Math.Round(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * precision)) / precision)!.ToJsonString()
                    : field == "strike" && value is "false" or "true"
                        ? JsonValue.Create(value == "true" ? "sngStrike" : "noStrike")!.ToJsonString() : value;
                if (field == "underline")
                    expected = JsonValue.Create(JsonNode.Parse(value)!.GetValue<string>() switch
                    { "single" => "sng", "double" => "dbl", var token => token })!.ToJsonString();
                if (field == "highlight")
                    expected = JsonValue.Create(value.StartsWith("{", StringComparison.Ordinal) ? "#808080"
                        : JsonNode.Parse(value)!.GetValue<string>()[..7].ToUpperInvariant())!.ToJsonString();
                if (field == "color")
                    expected = JsonValue.Create(value.StartsWith("{", StringComparison.Ordinal) ? "#808080"
                        : JsonNode.Parse(value)!.GetValue<string>().ToUpperInvariant())!.ToJsonString();
                if (field == "gradient")
                {
                    expected = DefaultGradientCases().Single(c => c.Input == value).Expected;
                    Assert.True(JsonNode.DeepEquals(JsonNode.Parse(expected), ParagraphDefaultScalar(restored, kind, field)));
                }
                else Assert.Equal(expected, ParagraphDefaultScalar(restored, kind, field)?.ToJsonString());
                var expectedProjection = field == "underline"
                    ? JsonValue.Create(JsonNode.Parse(expected)!.GetValue<string>() switch
                    { "sng" => "single", "dbl" => "double", var token => token })!.ToJsonString() : expected;
                if (field == "gradient")
                    Assert.True(JsonNode.DeepEquals(JsonNode.Parse(expectedProjection), FirstTextParagraph(Project(restored))["style"]!["defaultText"]![field]));
                else Assert.Equal(expectedProjection, FirstTextParagraph(Project(restored))["style"]!["defaultText"]![field]!.ToJsonString());
                AssertOnlyBodyPropertyChanged(deleted, restored, kind, "paragraphDefault." + field);
            }
        }
        if (field == "strike")
        {
            var cancel = Project(source);
            FirstTextParagraph(cancel)["style"]!["defaultText"]![field] = false;
            var canceled = Compile(cancel, source).File.ToByteArray();
            Assert.Equal("noStrike", ParagraphDefaultScalar(canceled, kind, field)!.GetValue<string>());
            Assert.Equal("noStrike", FirstTextParagraph(Project(canceled))["style"]!["defaultText"]![field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(source, canceled, kind, "paragraphDefault.strike");
            var removeCanceled = Project(canceled);
            FirstTextParagraph(removeCanceled)["style"]!["defaultText"]!.AsObject().Remove(field);
            var omitted = Compile(removeCanceled, canceled).File.ToByteArray();
            Assert.Null(ParagraphDefaultScalar(omitted, kind, field));
            Assert.Null(FirstTextParagraph(Project(omitted))["style"]?["defaultText"]?[field]);
            AssertOnlyBodyPropertyChanged(canceled, omitted, kind, "paragraphDefault.strike");
        }
        if (field == "underline")
        {
            var cancel = Project(source);
            FirstTextParagraph(cancel)["style"]!["defaultText"]![field] = "none";
            var canceled = Compile(cancel, source).File.ToByteArray();
            Assert.Equal("none", ParagraphDefaultScalar(canceled, kind, field)!.GetValue<string>());
            Assert.Equal("none", FirstTextParagraph(Project(canceled))["style"]!["defaultText"]![field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(source, canceled, kind, "paragraphDefault.underline");
            var removeCanceled = Project(canceled);
            FirstTextParagraph(removeCanceled)["style"]!["defaultText"]!.AsObject().Remove(field);
            var omitted = Compile(removeCanceled, canceled).File.ToByteArray();
            Assert.Null(ParagraphDefaultScalar(omitted, kind, field));
            Assert.Null(FirstTextParagraph(Project(omitted))["style"]?["defaultText"]?[field]);
            AssertOnlyBodyPropertyChanged(canceled, omitted, kind, "paragraphDefault.underline");
        }
        var denied = Project(source);
        FirstTextParagraph(denied)["style"]!["defaultText"]![field] = isFont ? JsonValue.Create("Georgia")
            : field == "gradient" ? JsonNode.Parse(DefaultGradientCases().First().Input)
            : field == "color" ? JsonValue.Create("#FF0000")
            : field == "highlight" ? JsonValue.Create("#112233")
            : field == "underline" ? JsonValue.Create("double")
            : field == "strike" ? JsonValue.Create("dblStrike")
            : field == "capitalization" ? JsonValue.Create("small")
            : field == "language" ? JsonValue.Create("fr-FR")
            : field is "size" or "kerning" or "letterSpacing" or "baseline" ? JsonValue.Create(24) : JsonValue.Create(false);
        var capability = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!;
        var fields = capability["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.defaultText." + field));
        Assert.Empty(Compile(denied, source, success: false).File);
        var unsupported = Project(source);
        FirstTextParagraph(unsupported)["style"]!["defaultText"]!["glow"] = JsonNode.Parse("""{"radius":2,"color":"#112233"}""");
        Assert.Empty(Compile(unsupported, source, success: false).File);
        if (field == "size")
            foreach (var invalid in new[] { 0d, -1d, 768.01, 0.001, 0.01, 0.99 })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]!["size"] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "kerning")
            foreach (var invalid in new[] { -0.001, -1, 768.001, 769 })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "letterSpacing")
            foreach (var invalid in new[] { -769, -768.001, 768.001, 769 })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "baseline")
            foreach (var invalid in new[] { -401, -400.0001, 400.0001, 401 })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "capitalization")
            foreach (var invalid in new[] { "", "ALL", "normal", " small", "small " })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "strike")
            foreach (var invalid in new[] { "", "SngStrike", "single", " noStrike", "true" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "underline")
            foreach (var invalid in new[] { "", "Sng", "sngStrike", " single", "true" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "underline")
            foreach (var invalid in new JsonNode?[] { JsonValue.Create(true), JsonValue.Create(1), new JsonObject(), null })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid?.DeepClone();
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "highlight")
            foreach (var invalid in new[] { "\"#FFFFFF80\"", "\"#FFFFFF00\"", "\"none\"", "false", "null", """{"token":"missing"}""" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "color")
            foreach (var invalid in new[] { "\"#GGGGGG\"", "\"none\"", "false", "null", """{"token":"missing"}""" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "color")
            foreach (var retainColor in new[] { true, false })
            {
                var request = Project(source);
                var gradientDefaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
                if (!retainColor)
                {
                    gradientDefaults.Remove("color");
                    RemoveDefaultFieldAuthority(request, "gradient");
                }
                gradientDefaults["gradient"] = JsonNode.Parse("""{"kind":"linear","angle":0,"stops":[{"offset":0,"color":"#000000"},{"offset":1,"color":"#FFFFFF"}]}""");
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (field == "gradient")
        {
            foreach (var invalid in new[] {
                """{"stops":[]}""", """{"kind":"radial","angle":0,"stops":[{"offset":0,"color":"#112233"},{"offset":1,"color":"#FFFFFF"}]}""",
                """{"angle":361,"stops":[{"offset":0,"color":"#112233"},{"offset":1,"color":"#FFFFFF"}]}""",
                """{"stops":[{"offset":0.8,"color":"#112233"},{"offset":0.2,"color":"#FFFFFF"}]}""",
                """{"stops":[{"offset":0,"color":"#112233","opacity":1.01},{"offset":1,"color":"#FFFFFF"}]}""",
                """{"stops":[{"offset":0,"color":{"token":"missing"}},{"offset":1,"color":"#FFFFFF"}]}""", "false", "null" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = JsonNode.Parse(invalid);
                Assert.Empty(Compile(request, source, success: false).File);
            }
            var conflict = Project(source);
            FirstTextParagraph(conflict)["style"]!["defaultText"]!["color"] = "#112233";
            Assert.Empty(Compile(conflict, source, success: false).File);
        }
        if (field == "language")
            foreach (var invalid in new[] { "", "en_US", " en-US", "en-US ", "a", "en-" + string.Join("-", Enumerable.Repeat("abcdefgh", 6)) + "-abcdefg" })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        if (isFont)
            foreach (var invalid in new[] { "", "   ", new string('F', 256) })
            {
                var request = Project(source);
                FirstTextParagraph(request)["style"]!["defaultText"]![field] = invalid;
                Assert.Empty(Compile(request, source, success: false).File);
            }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("fontFamily", false)]
    [InlineData("fontFamily", true)]
    [InlineData("fontFamilyEastAsia", false)]
    [InlineData("fontFamilyEastAsia", true)]
    [InlineData("fontFamilyComplexScript", false)]
    [InlineData("fontFamilyComplexScript", true)]
    public void ParagraphDefaultFontRejectsUnmodeledNode(string field, bool childContent)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { [field] = "Arial", ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain native font" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, "text").Descendants<A.DefaultRunProperties>().First();
            var tag = field switch { "fontFamily" => "latin", "fontFamilyEastAsia" => "ea", _ => "cs" };
            if (childContent)
            {
                defaults.InnerXml = $"<a:{tag} xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" typeface=\"Arial\"><a:extLst/></a:{tag}>";
            }
            else defaults.ChildElements.Single(child => child.LocalName == tag)
                .SetAttribute(new OpenXmlAttribute("pitchFamily", "", "34"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?[field]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]![field] = "Georgia";
        Assert.Empty(Compile(projected, source, success: false).File);
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultLanguagePreservesUnmodeledNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["language"] = "en-US", ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain native language" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First().Language = "en_US";
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["language"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["language"] = "fr-FR";
        Assert.Empty(Compile(projected, source, success: false).File);
        var unrelated = Project(source);
        FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        var candidate = Compile(unrelated, source).File.ToByteArray();
        Assert.Equal("en_US", ParagraphDefaultScalar(candidate, "text", "language")!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultKerningPreservesUnmodeledNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["kerning"] = 12, ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain native kerning" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First().Kerning = 76801;
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["kerning"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["kerning"] = 12;
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal(768.01, ParagraphDefaultScalar(candidate, "text", "kerning")!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultLetterSpacingPreservesUnmodeledNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["letterSpacing"] = 12, ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain native spacing" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First().Spacing = 76801;
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["letterSpacing"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["letterSpacing"] = 12;
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal(768.01, ParagraphDefaultScalar(candidate, "text", "letterSpacing")!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultBaselinePreservesUnmodeledNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["baseline"] = 12, ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain native baseline" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First().Baseline = 400001;
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["baseline"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["baseline"] = 12;
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal(400.001, ParagraphDefaultScalar(candidate, "text", "baseline")!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultCapitalizationRetainsInvalidNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["capitalization"] = "all", ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain Mixed Case" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First()
                .SetAttribute(new OpenXmlAttribute("cap", "", "futureCaps"));
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["capitalization"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["capitalization"] = "small";
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal("futureCaps", ParagraphDefaultScalar(candidate, "text", "capitalization")!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal("futureCaps", ParagraphDefaultScalar(source, "text", "capitalization")!.GetValue<string>());
        Assert.Equal(original, source);
    }

    [Fact]
    public void ParagraphDefaultStrikeRetainsInvalidNativeValue()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["strike"] = "sngStrike", ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain Mixed Case" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
            Owner(doc, "text").Descendants<A.DefaultRunProperties>().First()
                .SetAttribute(new OpenXmlAttribute("strike", "", "futureStrike"));
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["strike"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["strike"] = "dblStrike";
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal("futureStrike", ParagraphDefaultScalar(candidate, "text", "strike")!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal("futureStrike", ParagraphDefaultScalar(source, "text", "strike")!.GetValue<string>());
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("effect")]
    [InlineData("effect-only")]
    public void ParagraphDefaultUnderlineRetainsUnmodeledNativeContent(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = new JsonObject
        {
            ["paragraphs"] = new JsonArray(new JsonObject
            {
                ["style"] = new JsonObject { ["defaultText"] = new JsonObject { ["underline"] = "sng", ["bold"] = true } },
                ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain Mixed Case" }),
            }),
        };
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, "text").Descendants<A.DefaultRunProperties>().First();
            if (profile == "unknown") defaults.SetAttribute(new OpenXmlAttribute("u", "", "futureUnderline"));
            else
            {
                defaults.Append(new A.UnderlineFillText());
                if (profile == "effect-only") defaults.Underline = null;
            }
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["underline"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["underline"] = "dbl";
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var unrelated = Project(source);
            var defaults = FirstTextParagraph(unrelated)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            Assert.Equal(ParagraphDefaultScalar(source, "text", "underline")?.ToJsonString(),
                ParagraphDefaultScalar(candidate, "text", "underline")?.ToJsonString());
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultHighlightRetainsThemeAndNeighborParagraph(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"highlight":"#FFFF00","bold":true}},"runs":[{"text":"First"}]},
            {"style":{"defaultText":{"highlight":"#FF0000"}},"runs":[{"text":"Second"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, kind).Descendants<A.DefaultRunProperties>().ToArray();
            for (var i = 0; i < 2; i++)
            {
                defaults[i].GetFirstChild<A.Highlight>()!.Remove();
                defaults[i].AddChild(new A.Highlight(new A.SchemeColor
                { Val = i == 0 ? A.SchemeColorValues.Accent1 : A.SchemeColorValues.Accent2 }), true);
            }
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Equal("accent1", FirstTextParagraph(projected)["style"]!["defaultText"]!["highlight"]!["token"]!.GetValue<string>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var input in new[] { """{"token":"accent3"}""", "\"#005500\"", """{"token":"accent1","tint":1,"shade":0.5}""" })
        {
            var request = Project(source);
            FirstTextParagraph(request)["style"]!["defaultText"]!["highlight"] = JsonNode.Parse(input);
            var candidate = Compile(request, source).File.ToByteArray();
            var expected = input.Contains("tint", StringComparison.Ordinal) ? "\"#808080\"" : input;
            Assert.Equal(expected, ParagraphDefaultScalar(candidate, kind, "highlight")!.ToJsonString());
            Assert.Equal(expected, FirstTextParagraph(Project(candidate))["style"]!["defaultText"]!["highlight"]!.ToJsonString());
            AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraphDefault.highlight");
        }
        var remove = Project(source); FirstTextParagraph(remove)["style"]!["defaultText"]!.AsObject().Remove("highlight");
        var removed = Compile(remove, source).File.ToByteArray();
        Assert.Null(ParagraphDefaultScalar(removed, kind, "highlight"));
        Assert.Null(FirstTextParagraph(Project(removed))["style"]?["defaultText"]?["highlight"]);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraphDefault.highlight");
        var restore = Project(removed);
        FirstTextParagraph(restore)["style"]!["defaultText"]!["highlight"] = JsonNode.Parse("""{"token":"accent1"}""");
        var restored = Compile(restore, removed).File.ToByteArray();
        Assert.Equal("""{"token":"accent1"}""", ParagraphDefaultScalar(restored, kind, "highlight")!.ToJsonString());
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraphDefault.highlight");
        var shadowed = Project(source);
        shadowed["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent3":{"kind":"color","value":"#abcdef"}}""");
        FirstTextParagraph(shadowed)["style"]!["defaultText"]!["highlight"] = JsonNode.Parse("""{"token":"accent3"}""");
        var shadowedBytes = Compile(shadowed, source).File.ToByteArray();
        Assert.Equal("#ABCDEF", ParagraphDefaultScalar(shadowedBytes, kind, "highlight")!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(source, shadowedBytes, kind, "paragraphDefault.highlight");
        shadowed["design"]!["grammar"]!["tokens"]!["accent3"]!["value"] = "#ABCDEF80";
        Assert.Empty(Compile(shadowed, source, success: false).File);
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultColorRetainsThemeAndNeighborParagraph(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"color":"#FFFF00","bold":true}},"runs":[{"text":"First"}]},
            {"style":{"defaultText":{"color":"#FF0000"}},"runs":[{"text":"Second"}]},
            {"style":{"defaultText":{"color":"#112233"}},"runs":[{"text":"Third"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, kind).Descendants<A.DefaultRunProperties>().ToArray();
            for (var i = 0; i < 2; i++)
            {
                defaults[i].GetFirstChild<A.SolidFill>()!.Remove();
                defaults[i].AddChild(new A.SolidFill(new A.SchemeColor
                { Val = i == 0 ? A.SchemeColorValues.Accent1 : A.SchemeColorValues.Accent2 }), true);
            }
            defaults[1].GetFirstChild<A.SolidFill>()!.GetFirstChild<A.SchemeColor>()!.Append(new A.Alpha { Val = 42000 });
            defaults[2].GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Alpha { Val = 42000 });
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Equal("accent1", FirstTextParagraph(projected)["style"]!["defaultText"]!["color"]!["token"]!.GetValue<string>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        foreach (var input in new[] { """{"token":"accent3"}""", """{"token":"accent3","alpha":0}""", """{"token":"accent3","alpha":1}""", """{"token":"accent3","alpha":0.12345}""", "\"#005500\"", """{"token":"accent1","tint":1,"shade":0.5}""" })
        {
            var request = Project(source);
            FirstTextParagraph(request)["style"]!["defaultText"]!["color"] = JsonNode.Parse(input);
            var candidate = Compile(request, source).File.ToByteArray();
            var expected = input.Contains("tint", StringComparison.Ordinal) ? "\"#808080\"" : input;
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse(expected), ParagraphDefaultScalar(candidate, kind, "color")));
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse(expected), FirstTextParagraph(Project(candidate))["style"]!["defaultText"]!["color"]));
            AssertOnlyBodyPropertyChanged(source, candidate, kind, "paragraphDefault.color");
        }
        var remove = Project(source); FirstTextParagraph(remove)["style"]!["defaultText"]!.AsObject().Remove("color");
        var removed = Compile(remove, source).File.ToByteArray();
        Assert.Null(ParagraphDefaultScalar(removed, kind, "color"));
        Assert.Null(FirstTextParagraph(Project(removed))["style"]?["defaultText"]?["color"]);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraphDefault.color");
        var restore = Project(removed);
        FirstTextParagraph(restore)["style"]!["defaultText"]!["color"] = JsonNode.Parse("""{"token":"accent1"}""");
        var restored = Compile(restore, removed).File.ToByteArray();
        Assert.Equal("""{"token":"accent1"}""", ParagraphDefaultScalar(restored, kind, "color")!.ToJsonString());
        AssertOnlyBodyPropertyChanged(removed, restored, kind, "paragraphDefault.color");
        var shadowed = Project(source);
        shadowed["design"]!["grammar"]!["tokens"] = JsonNode.Parse("""{"accent3":{"kind":"color","value":"#abcdef"}}""");
        FirstTextParagraph(shadowed)["style"]!["defaultText"]!["color"] = JsonNode.Parse("""{"token":"accent3"}""");
        var shadowedBytes = Compile(shadowed, source).File.ToByteArray();
        Assert.Equal("#ABCDEF", ParagraphDefaultScalar(shadowedBytes, kind, "color")!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(source, shadowedBytes, kind, "paragraphDefault.color");
        shadowed["design"]!["grammar"]!["tokens"]!["accent3"]!["value"] = "#ABCDEF80";
        var translucent = Compile(shadowed, source).File.ToByteArray();
        Assert.Equal("#ABCDEF80", ParagraphDefaultScalar(translucent, kind, "color")!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(source, translucent, kind, "paragraphDefault.color");
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("alpha")]
    [InlineData("duplicate")]
    [InlineData("hidden-content")]
    [InlineData("malformed-text")]
    public void ParagraphDefaultHighlightRetainsUnmodeledNativeContent(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"highlight":"#FFFF00","bold":true}},"runs":[{"text":"Retain source text"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, "text").Descendants<A.DefaultRunProperties>().First();
            var highlight = defaults.GetFirstChild<A.Highlight>()!;
            if (profile == "alpha") highlight.GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Alpha { Val = 50000 });
            else if (profile == "duplicate") defaults.Append(highlight.CloneNode(true));
            else if (profile == "hidden-content")
                highlight.InnerXml = """<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FFFF00"><future:color xmlns:future="urn:officekit:test" value="retain"/></a:srgbClr>""";
            else highlight.InnerXml = """<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FFFF00">retain</a:srgbClr>""";
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Null(FirstTextParagraph(projected)["style"]?["defaultText"]?["highlight"]);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["highlight"] = "#112233";
        Assert.Empty(Compile(projected, source, success: false).File);
        foreach (var remove in new[] { false, true })
        {
            var request = Project(source); var defaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            if (profile == "malformed-text")
            {
                // The SDK drops illegal character data inside this color node.
                // Keep binding rejection instead of producing a lossy candidate.
                Assert.Empty(Compile(request, source, success: false).File);
                continue;
            }
            var candidate = Compile(request, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("luminance")]
    [InlineData("duplicate")]
    [InlineData("no-fill")]
    public void ParagraphDefaultColorRetainsSourceOwnedPaint(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"color":"#FFFF00","bold":true}},"runs":[{"text":"Retain source paint"}]}]}
            """);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var doc = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(doc, "text").Descendants<A.DefaultRunProperties>().First();
            var fill = defaults.GetFirstChild<A.SolidFill>()!;
            if (profile == "duplicate") defaults.Append(fill.CloneNode(true));
            else
            {
                fill.Remove();
                if (profile == "no-fill") defaults.AddChild(new A.NoFill(), true);
                else defaults.AddChild(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 50000 })
                    { Val = A.SchemeColorValues.Accent1 }), true);
            }
        }
        var source = stream.ToArray(); var original = source.ToArray();
        var projected = Project(source);
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        FirstTextParagraph(projected)["style"]!["defaultText"]!["color"] = "#112233";
        Assert.Empty(Compile(projected, source, success: false).File);
        if (profile == "luminance")
            foreach (var wrapper in new[] { false, true })
            {
                var request = Project(source);
                if (wrapper) FirstTextParagraph(request).Remove("style");
                else FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject().Remove("color");
                Assert.Empty(Compile(request, source, success: false).File);
            }
        foreach (var remove in new[] { false, true })
        {
            var request = Project(source); var defaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(request, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphDefaultGradientAndColorSwitchWithBothAuthorities(string kind)
    {
        var program = Program(kind);
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"color":"#112233","bold":true}},"runs":[{"text":"Retain text","style":{"color":"#AABBCC80"}}]}]}
            """);
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        var defaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
        defaults.Remove("color"); defaults["gradient"] = JsonNode.Parse(DefaultGradientJson);
        foreach (var field in new[] { "color", "gradient" })
        {
            var denied = request.DeepClone().AsObject(); RemoveDefaultFieldAuthority(denied, field);
            Assert.Empty(Compile(denied, source, success: false).File);
        }
        var gradient = Compile(request, source).File.ToByteArray();
        Assert.Null(ParagraphDefaultScalar(gradient, kind, "color"));
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse(DefaultGradientJson), ParagraphDefaultScalar(gradient, kind, "gradient")));
        Assert.True(JsonNode.DeepEquals(JsonNode.Parse(DefaultGradientJson), FirstTextParagraph(Project(gradient))["style"]!["defaultText"]!["gradient"]));
        AssertOnlyBodyPropertyChanged(source, gradient, kind, "paragraphDefault.paint");
        var restore = Project(gradient);
        var restoredDefaults = FirstTextParagraph(restore)["style"]!["defaultText"]!.AsObject();
        restoredDefaults.Remove("gradient"); restoredDefaults["color"] = "#FF000080";
        foreach (var field in new[] { "color", "gradient" })
        {
            var denied = restore.DeepClone().AsObject(); RemoveDefaultFieldAuthority(denied, field);
            Assert.Empty(Compile(denied, gradient, success: false).File);
        }
        var color = Compile(restore, gradient).File.ToByteArray();
        Assert.Null(ParagraphDefaultScalar(color, kind, "gradient"));
        Assert.Equal("#FF000080", ParagraphDefaultScalar(color, kind, "color")!.GetValue<string>());
        Assert.Equal("#FF000080", FirstTextParagraph(Project(color))["style"]!["defaultText"]!["color"]!.GetValue<string>());
        AssertOnlyBodyPropertyChanged(gradient, color, kind, "paragraphDefault.paint");

        using var stream = new MemoryStream(); stream.Write(source);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var fill = Owner(document, kind).Descendants<A.DefaultRunProperties>().First().GetFirstChild<A.SolidFill>()!;
            fill.RemoveAllChildren();
            fill.Append(new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent1 });
        }
        var transformed = stream.ToArray();
        var incompatible = Project(transformed);
        var changed = FirstTextParagraph(incompatible)["style"]!["defaultText"]!.AsObject();
        changed.Remove("color"); changed["gradient"] = JsonNode.Parse(DefaultGradientJson);
        Assert.Empty(Compile(incompatible, transformed, success: false).File);
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("scaled")]
    [InlineData("theme")]
    [InlineData("duplicate")]
    [InlineData("unknown-element")]
    public void ParagraphDefaultGradientRetainsUnmodeledSourcePaint(string profile)
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"defaultText":{"bold":true}},"runs":[{"text":"Retain source text"}]}]}
            """);
        FirstTextParagraph(program)["style"]!["defaultText"]!["gradient"] = JsonNode.Parse(DefaultGradientJson);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var defaults = Owner(document, "text").Descendants<A.DefaultRunProperties>().First();
            var fill = defaults.GetFirstChild<A.GradientFill>()!;
            if (profile == "scaled") fill.GetFirstChild<A.LinearGradientFill>()!.Scaled = true;
            else if (profile == "duplicate") defaults.Append(fill.CloneNode(true));
            else
            {
                var stop = fill.GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>().First();
                if (profile == "theme")
                {
                    stop.RemoveAllChildren();
                    stop.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
                }
                else stop.GetFirstChild<A.RgbColorModelHex>()!.InnerXml =
                    """<future:color xmlns:future="urn:officekit:test" value="retain"/>""";
            }
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        foreach (var field in new[] { "gradient", "color" })
        {
            var replace = Project(source);
            FirstTextParagraph(replace)["style"]!["defaultText"]![field] = field == "gradient" ? JsonNode.Parse(DefaultGradientJson) : JsonValue.Create("#FF0000");
            Assert.Empty(Compile(replace, source, success: false).File);
        }
        foreach (var remove in new[] { false, true })
        {
            var request = Project(source);
            var defaults = FirstTextParagraph(request)["style"]!["defaultText"]!.AsObject();
            if (remove) defaults.Remove("bold"); else defaults["bold"] = false;
            var candidate = Compile(request, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
        }
        Assert.Equal(original, source);
    }

    private const string DefaultGradientJson = """{"kind":"linear","angle":30,"stops":[{"offset":0,"color":"#112233"},{"offset":1,"color":"#AABBCC","opacity":1}]}""";

    private static (string Input, string Expected)[] DefaultGradientCases() => new[]
    {
        ("""{"kind":"linear","angle":-90,"stops":[{"offset":0,"color":"#000000","opacity":0},{"offset":0,"color":"#FF0000","opacity":1},{"offset":1,"color":"#FFFFFF"}]}""",
         """{"kind":"linear","angle":270,"stops":[{"offset":0,"color":"#000000","opacity":0},{"offset":0,"color":"#FF0000","opacity":1},{"offset":1,"color":"#FFFFFF"}]}"""),
        ("""{"kind":"radial","stops":[{"offset":0,"color":"#00000080"},{"offset":0.5,"color":{"token":"accent1"}},{"offset":1,"color":{"token":"mark","tint":1,"shade":0.5}}]}""",
         """{"kind":"radial","stops":[{"offset":0,"color":"#000000","opacity":0.50196},{"offset":0.5,"color":"#000000"},{"offset":1,"color":"#808080"}]}"""),
        ("""{"angle":359.999999,"stops":[{"offset":0.123456,"color":"#112233"},{"offset":1,"color":"#AABBCC","opacity":0.123456}]}""",
         """{"kind":"linear","angle":0,"stops":[{"offset":0.12346,"color":"#112233"},{"offset":1,"color":"#AABBCC","opacity":0.12346}]}"""),
    };

    private static void RemoveDefaultFieldAuthority(JsonObject program, string field)
    {
        var fields = program["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.defaultText." + field));
    }

    private static JsonNode? NativeDefaultGradient(A.DefaultRunProperties? defaults)
    {
        var gradient = defaults?.GetFirstChild<A.GradientFill>();
        if (gradient is null) return null;
        var linear = gradient.GetFirstChild<A.LinearGradientFill>();
        var stops = new JsonArray();
        foreach (var stop in gradient.GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>())
        {
            var color = stop.GetFirstChild<A.RgbColorModelHex>()!;
            var value = new JsonObject { ["offset"] = stop.Position!.Value / 100000d, ["color"] = "#" + color.Val!.Value };
            if (color.GetFirstChild<A.Alpha>() is { } alpha) value["opacity"] = alpha.Val!.Value / 100000d;
            stops.Add(value);
        }
        var output = new JsonObject { ["kind"] = linear is null ? "radial" : "linear", ["stops"] = stops };
        if (linear is not null) output["angle"] = linear.Angle!.Value / 60000d;
        return output;
    }

    private static JsonNode? NativeDefaultColor(A.DefaultRunProperties? defaults)
    {
        var color = defaults?.GetFirstChild<A.SolidFill>()?.FirstChild;
        if (color is null) return null;
        var alpha = color.GetFirstChild<A.Alpha>()?.Val?.Value;
        if (color is A.SchemeColor scheme)
        {
            var output = new JsonObject { ["token"] = scheme.Val!.InnerText };
            if (alpha is { } value) output["alpha"] = value / 100000d;
            return output;
        }
        var rgb = "#" + ((A.RgbColorModelHex)color).Val!.InnerText.ToUpperInvariant();
        if (alpha is { } opacity) rgb += ((int)Math.Round(opacity / 100000d * 255)).ToString("X2");
        return JsonValue.Create(rgb);
    }

    private static JsonObject FirstTextParagraph(JsonObject program) =>
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!.AsObject();

    private static JsonNode? ParagraphDefaultScalar(byte[] bytes, string kind, string field)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var paragraph = Owner(document, kind).Descendants<A.Paragraph>().First();
        var defaults = paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
        return field switch
        {
            "bold" => defaults?.Bold is { } bold ? JsonValue.Create(bold.Value) : null,
            "italic" => defaults?.Italic is { } italic ? JsonValue.Create(italic.Value) : null,
            "size" => defaults?.FontSize is { } size ? JsonValue.Create(size.Value / 100d) : null,
            "language" => defaults?.Language is { } language ? JsonValue.Create(language.Value) : null,
            "kerning" => defaults?.Kerning is { } kerning ? JsonValue.Create(kerning.Value / 100d) : null,
            "letterSpacing" => defaults?.Spacing is { } spacing ? JsonValue.Create(spacing.Value / 100d) : null,
            "baseline" => defaults?.Baseline is { } baseline ? JsonValue.Create(baseline.Value / 1000d) : null,
            "capitalization" => defaults?.Capital is { } caps ? JsonValue.Create(caps.InnerText) : null,
            "strike" => defaults?.Strike is { } strike ? JsonValue.Create(strike.InnerText) : null,
            "underline" => defaults?.Underline is { } underline ? JsonValue.Create(underline.InnerText) : null,
            "color" => NativeDefaultColor(defaults),
            "gradient" => NativeDefaultGradient(defaults),
            "highlight" => defaults?.GetFirstChild<A.Highlight>() is { } highlight
                ? highlight.GetFirstChild<A.SchemeColor>() is { } scheme
                    ? new JsonObject { ["token"] = scheme.Val!.InnerText }
                    : JsonValue.Create("#" + highlight.GetFirstChild<A.RgbColorModelHex>()!.Val!.InnerText.ToUpperInvariant()) : null,
            "fontFamily" => defaults?.GetFirstChild<A.LatinFont>()?.Typeface is { } family ? JsonValue.Create(family.Value) : null,
            "fontFamilyEastAsia" => defaults?.GetFirstChild<A.EastAsianFont>()?.Typeface is { } eastAsian ? JsonValue.Create(eastAsian.Value) : null,
            "fontFamilyComplexScript" => defaults?.GetFirstChild<A.ComplexScriptFont>()?.Typeface is { } complexScript ? JsonValue.Create(complexScript.Value) : null,
            _ => throw new ArgumentOutOfRangeException(nameof(field)),
        };
    }

    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void TextWarpPresetAndGuidesCanBeRemovedAndRestored(string kind)
    {
        var program = Program(kind);
        Style(program, kind)["textWarpPreset"] = "textArchUp";
        Style(program, kind)["textWarpAdjustments"] = WarpGuides();
        Style(program, kind)["fromWordArt"] = false;
        Style(program, kind)["flatTextZ"] = 0;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var invalid = Project(source); Style(invalid, kind).Remove("textWarpPreset");
        Assert.Empty(Compile(invalid, source, success: false).File);
        foreach (var empty in new[] { false, true })
        {
            var request = Project(source);
            if (empty) Style(request, kind)["textWarpAdjustments"] = new JsonArray();
            else Style(request, kind).Remove("textWarpAdjustments");
            var cleared = Compile(request, source).File.ToByteArray();
            var warp = Assert.Single(Body(cleared, kind).Elements<A.PresetTextWarp>());
            Assert.Equal("textArchUp", warp.Preset!.InnerText);
            Assert.Empty(warp.ChildElements);
            Assert.False(Style(Project(cleared), kind).ContainsKey("textWarpAdjustments"));
            AssertOnlyBodyPropertyChanged(source, cleared, kind, "textWarpAdjustments");
            var restore = Project(cleared); Style(restore, kind)["textWarpAdjustments"] = WarpGuides();
            var restored = Compile(restore, cleared).File.ToByteArray();
            AssertWarpGuides(restored, kind);
            AssertOnlyBodyPropertyChanged(cleared, restored, kind, "textWarpAdjustments");
        }
        var remove = Project(source);
        Style(remove, kind).Remove("textWarpPreset"); Style(remove, kind).Remove("textWarpAdjustments");
        var absent = Compile(remove, source).File.ToByteArray();
        Assert.Empty(Body(absent, kind).Elements<A.PresetTextWarp>());
        Assert.False(Style(Project(absent), kind).ContainsKey("textWarpPreset"));
        AssertOnlyBodyPropertyChanged(source, absent, kind, "textWarpPreset");
        foreach (var preset in new[] { "textNoShape", "textArchUp", "textWave1" })
        {
            var restore = Project(absent); Style(restore, kind)["textWarpPreset"] = preset;
            Style(restore, kind)["textWarpAdjustments"] = WarpGuides();
            var restored = Compile(restore, absent).File.ToByteArray();
            Assert.Equal(preset, Style(Project(restored), kind)["textWarpPreset"]!.GetValue<string>());
            AssertWarpGuides(restored, kind);
            AssertOnlyBodyPropertyChanged(absent, restored, kind, "textWarpPreset");
            var again = Project(restored);
            Style(again, kind).Remove("textWarpPreset"); Style(again, kind).Remove("textWarpAdjustments");
            Assert.Empty(Body(Compile(again, restored).File.ToByteArray(), kind).Elements<A.PresetTextWarp>());
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void TextWarpOnlyAndCombinedBodyStylesCanBeRemoved(string kind)
    {
        foreach (var combined in new[] { false, true })
        {
            var program = Program(kind);
            if (!combined) Style(program, kind).Clear();
            Style(program, kind)["textWarpPreset"] = "textArchUp";
            Style(program, kind)["textWarpAdjustments"] = WarpGuides();
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            Style(request, kind).Parent!.AsObject().Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Empty(Body(deleted, kind).Elements<A.PresetTextWarp>());
            Assert.Null(Body(deleted, kind).UpRight); Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).Wrap); Assert.Null(Body(deleted, kind).LeftInset);
            if (!combined) AssertOnlyBodyPropertyChanged(source, deleted, kind, "textWarpPreset");
            var restore = Project(deleted); var element = restore["pages"]![0]!["elements"]![0]!;
            var style = new JsonObject { ["textWarpPreset"] = "textNoShape", ["textWarpAdjustments"] = WarpGuides() };
            if (kind == "shape") element["textStyle"] = style;
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = style,
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal("textNoShape", Style(Project(restored), kind)["textWarpPreset"]!.GetValue<string>());
            AssertWarpGuides(restored, kind);
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "textWarpPreset");
        }
    }

    private static JsonArray WarpGuides() => new(
        new JsonObject { ["name"] = "adj", ["value"] = 0 },
        new JsonObject { ["name"] = "adj2", ["value"] = int.MinValue },
        new JsonObject { ["name"] = "adj3", ["value"] = int.MaxValue });

    private static void AssertWarpGuides(byte[] source, string kind)
    {
        var warp = Assert.Single(Body(source, kind).Elements<A.PresetTextWarp>());
        Assert.Equal(new[] { "adj", "adj2", "adj3" }, warp.Descendants<A.ShapeGuide>().Select(g => g.Name!.Value));
        Assert.Equal(new[] { "val 0", "val -2147483648", "val 2147483647" }, warp.Descendants<A.ShapeGuide>().Select(g => g.Formula!.Value));
        Assert.True(JsonNode.DeepEquals(WarpGuides(), Style(Project(source), kind)["textWarpAdjustments"]));
    }

    [Fact]
    public void TextWarpDeletionWireAndNativeGuards()
    {
        var command = new PresentationTextBodyProperties { NoTextWarpPreset = true };
        Assert.Equal(new byte[] { 240, 2, 1 }, command.ToByteArray());
        var remove = PresentationTextBodyProperties.Parser.ParseFrom(command.ToByteArray());
        Assert.True(remove.NoTextWarpPreset); Assert.False(remove.HasTextWarpPreset);
        var body = new PresentationTextBody { BodyProperties = remove };
        PptxBodyPropertiesCodec.Validate(body);
        foreach (var invalid in new[] {
            new PresentationTextBodyProperties { NoTextWarpPreset = false },
            new PresentationTextBodyProperties { NoTextWarpPreset = true, TextWarpPreset = "textNoShape" },
            new PresentationTextBodyProperties { NoTextWarpPreset = true, TextWarpAdjustments = { new PresentationTextWarpAdjustment { Name = "adj", Value = 0 } } } })
        {
            Assert.Throws<CodecException>(() => PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = invalid }));
            Assert.False(PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(invalid));
        }
        foreach (var xml in new[] {
            """<a:prstTxWarp prst="textArchUp"/><a:prstTxWarp prst="textArchUp"/>""",
            """<a:prstTxWarp prst="unknown"/>""",
            """<a:prstTxWarp prst="textArchUp" unknown="1"/>""",
            """<a:prstTxWarp prst="textArchUp"><a:avLst/><a:avLst/></a:prstTxWarp>""",
            """<a:prstTxWarp prst="textArchUp"><a:avLst><a:gd name="adj" fmla="sum 1 2 3"/></a:avLst></a:prstTxWarp>""",
            """<a:prstTxWarp prst="textArchUp"><a:avLst><a:gd name="adj" fmla="val 0"><a:extLst/></a:gd></a:avLst></a:prstTxWarp>""" })
        {
            var properties = new A.BodyProperties("""<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">""" + xml + "</a:bodyPr>");
            var native = new DocumentFormat.OpenXml.Presentation.TextBody(properties);
            Assert.True(Record.Exception(() => PptxBodyPropertiesCodec.Apply(native, body)) is CodecException, xml);
            Assert.NotEmpty(native.Descendants<A.PresetTextWarp>());
        }
    }

    [Fact]
    public void FlatTextDepthDeletionWireAndNativeGuards()
    {
        var remove = new PresentationTextBodyProperties { NoFlatTextZ = true };
        Assert.Equal(new byte[] { 232, 2, 1 }, remove.ToByteArray());
        remove = PresentationTextBodyProperties.Parser.ParseFrom(remove.ToByteArray());
        Assert.True(remove.HasNoFlatTextZ); Assert.True(remove.NoFlatTextZ); Assert.False(remove.HasFlatTextZ);
        Assert.Equal(new byte[] { 184, 2, 0 }, new PresentationTextBodyProperties { FlatTextZ = 0 }.ToByteArray());
        var body = new PresentationTextBody { BodyProperties = remove };
        PptxBodyPropertiesCodec.Validate(body);
        foreach (var invalid in new[] {
            new PresentationTextBodyProperties { NoFlatTextZ = false },
            new PresentationTextBodyProperties { NoFlatTextZ = true, FlatTextZ = 0 },
            new PresentationTextBodyProperties { NoFlatTextZ = true, FlatTextZ = -1 } })
        {
            Assert.Throws<CodecException>(() => PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = invalid }));
            Assert.False(PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(invalid));
        }
        var canonical = new DocumentFormat.OpenXml.Presentation.TextBody(new A.BodyProperties(new A.FlatText { Z = 0 }));
        PptxBodyPropertiesCodec.Apply(canonical, body);
        Assert.Empty(canonical.GetFirstChild<A.BodyProperties>()!.ChildElements);
        foreach (var xml in new[] {
            """<a:flatTx z="0"/><a:flatTx z="1"/>""",
            """<a:flatTx/>""",
            """<a:flatTx z="01"/>""",
            """<a:flatTx z="2147483648"/>""",
            """<a:flatTx z="0" unknown="1"/>""",
            """<a:flatTx z="0"><a:extLst/></a:flatTx>""" })
        {
            var properties = new A.BodyProperties("""<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">""" + xml + "</a:bodyPr>");
            var native = new DocumentFormat.OpenXml.Presentation.TextBody(properties);
            Assert.True(Record.Exception(() => PptxBodyPropertiesCodec.Apply(native, body)) is CodecException, xml);
            Assert.NotEmpty(native.Descendants<A.FlatText>());
        }
    }

    [Theory]
    [InlineData(32, 40)]
    [InlineData(33, 41)]
    [InlineData(34, 42)]
    [InlineData(35, 43)]
    [InlineData(36, 44)]
    public void OptionalBooleanDeletionWireIntentIsUnambiguous(int setterNumber, int deleteNumber)
    {
        var setter = PresentationTextBodyProperties.Descriptor.FindFieldByNumber(setterNumber).Accessor;
        var deleter = PresentationTextBodyProperties.Descriptor.FindFieldByNumber(deleteNumber).Accessor;
        foreach (var value in new[] { false, true })
        {
            var properties = new PresentationTextBodyProperties();
            setter.SetValue(properties, value);
            Assert.Equal(new byte[] { (byte)(0x80 | ((setterNumber << 3) & 0x7f)), 0x02, value ? (byte)1 : (byte)0 }, properties.ToByteArray());
            PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = properties });
        }
        var command = new PresentationTextBodyProperties(); deleter.SetValue(command, true);
        var remove = PresentationTextBodyProperties.Parser.ParseFrom(command.ToByteArray());
        Assert.True(deleter.HasValue(remove)); Assert.Equal(true, deleter.GetValue(remove)); Assert.False(setter.HasValue(remove));
        PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = remove });
        Assert.True(PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(remove));
        foreach (bool? setterValue in new bool?[] { null, false, true })
        {
            var invalid = new PresentationTextBodyProperties();
            deleter.SetValue(invalid, setterValue is not null);
            if (setterValue is { } value) setter.SetValue(invalid, value);
            Assert.Throws<CodecException>(() => PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = invalid }));
            Assert.False(PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(invalid));
        }
    }

    [Theory]
    [InlineData("shape", "columnDirection")]
    [InlineData("shape", "verticalText")]
    [InlineData("shape", "wrap")]
    [InlineData("shape", "horizontalOverflow")]
    [InlineData("shape", "verticalOverflow")]
    [InlineData("shape", "verticalAlignment")]
    [InlineData("text", "columnDirection")]
    [InlineData("text", "verticalText")]
    [InlineData("text", "wrap")]
    [InlineData("text", "horizontalOverflow")]
    [InlineData("text", "verticalOverflow")]
    [InlineData("text", "verticalAlignment")]
    [InlineData("master", "columnDirection")]
    [InlineData("master", "verticalText")]
    [InlineData("master", "wrap")]
    [InlineData("master", "horizontalOverflow")]
    [InlineData("master", "verticalOverflow")]
    [InlineData("master", "verticalAlignment")]
    [InlineData("layout", "columnDirection")]
    [InlineData("layout", "verticalText")]
    [InlineData("layout", "wrap")]
    [InlineData("layout", "horizontalOverflow")]
    [InlineData("layout", "verticalOverflow")]
    [InlineData("layout", "verticalAlignment")]
    [InlineData("table", "columnDirection")]
    [InlineData("table", "verticalText")]
    [InlineData("table", "wrap")]
    [InlineData("table", "horizontalOverflow")]
    [InlineData("table", "verticalOverflow")]
    [InlineData("table", "verticalAlignment")]
    public void EnumBodyPropertySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind); Style(program, kind)[field] = BodyPropertyValues(field).Last();
        if (field == "horizontalOverflow") Style(program, kind)["verticalOverflow"] = "ellipsis";
        if (field == "verticalOverflow") Style(program, kind)["horizontalOverflow"] = "clip";
        if (field == "verticalAlignment") Style(program, kind)["anchorCenter"] = true;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(BodyPropertyValues(field).Last(), Style(request, kind)[field]!.GetValue<string>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        AssertBodyPropertyValue(source, kind, field, BodyPropertyValues(field).Last());
        Style(request, kind).Remove(field);
        var deleted = Compile(request, source).File.ToByteArray();
        AssertBodyPropertyValue(deleted, kind, field, null);
        Assert.False(Style(Project(deleted), kind).ContainsKey(field));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
        foreach (var value in BodyPropertyValues(field))
        {
            var restore = Project(deleted);
            Style(restore, kind)[field] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            AssertBodyPropertyValue(restored, kind, field, value);
            Assert.Equal(value, Style(Project(restored), kind)[field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove(field);
            AssertBodyPropertyValue(Compile(removeAgain, restored).File.ToByteArray(), kind, field, null);
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape", "columnDirection")]
    [InlineData("shape", "verticalText")]
    [InlineData("shape", "wrap")]
    [InlineData("shape", "horizontalOverflow")]
    [InlineData("shape", "verticalOverflow")]
    [InlineData("shape", "verticalAlignment")]
    [InlineData("table", "columnDirection")]
    [InlineData("table", "verticalText")]
    [InlineData("table", "wrap")]
    [InlineData("table", "horizontalOverflow")]
    [InlineData("table", "verticalOverflow")]
    [InlineData("table", "verticalAlignment")]
    public void EnumBodyPropertyOnlyAndCombinedRemovableStylesPreservePresence(string kind, string field)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style[field] = BodyPropertyValues(field).Last();
            if (withOtherProperties)
            {
                style["upright"] = false; style["rotation"] = 12;
                foreach (var sibling in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" }.Where(key => key != field))
                    style[sibling] = BodyPropertyValues(sibling).First();
            }
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            AssertBodyPropertyValue(deleted, kind, field, null);
            Assert.Null(Body(deleted, kind).UpRight);
            Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).Vertical);
            Assert.Null(Body(deleted, kind).RightToLeftColumns);
            Assert.Null(Body(deleted, kind).Wrap);
            Assert.Null(Body(deleted, kind).HorizontalOverflow);
            Assert.Null(Body(deleted, kind).VerticalOverflow);
            Assert.Null(Body(deleted, kind).Anchor);
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { [field] = BodyPropertyValues(field).First() };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { [field] = BodyPropertyValues(field).First() },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            AssertBodyPropertyValue(restored, kind, field, BodyPropertyValues(field).First());
            Assert.Equal(BodyPropertyValues(field).First(), Style(Project(restored), kind)[field]!.GetValue<string>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
        }
    }

    [Theory]
    [InlineData("shape", "columnGap")]
    [InlineData("shape", "columns")]
    [InlineData("shape", "flatTextZ")]
    [InlineData("text", "columnGap")]
    [InlineData("text", "columns")]
    [InlineData("text", "flatTextZ")]
    [InlineData("master", "columnGap")]
    [InlineData("master", "columns")]
    [InlineData("master", "flatTextZ")]
    [InlineData("layout", "columnGap")]
    [InlineData("layout", "columns")]
    [InlineData("layout", "flatTextZ")]
    [InlineData("table", "columnGap")]
    [InlineData("table", "columns")]
    [InlineData("table", "flatTextZ")]
    public void NumericBodyPropertySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind); var style = Style(program, kind);
        style["columnGap"] = 12.5; style["columns"] = 3; style["columnDirection"] = "right-to-left";
        if (field == "flatTextZ") { style[field] = -42; style["textWarpPreset"] = "textArchUp"; }
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(field == "flatTextZ" ? -42 : field == "columns" ? 3 : 12.5, Style(request, kind)[field]!.GetValue<double>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        Assert.Equal(field == "flatTextZ" ? -42 : field == "columns" ? 3 : 158750, NumericBodyProperty(source, kind, field));
        Style(request, kind).Remove(field);
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(NumericBodyProperty(deleted, kind, field));
        Assert.False(Style(Project(deleted), kind).ContainsKey(field));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
        foreach (var value in field == "flatTextZ" ? new[] { (double)int.MinValue, -42d, 0d, (double)int.MaxValue } : field == "columns" ? new[] { 1d, 3d, 16d } : new[] { 0d, 12.5d, 10000d })
        {
            var restore = Project(deleted);
            Style(restore, kind)[field] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(checked((int)(value * (field == "columnGap" ? 12700 : 1))), NumericBodyProperty(restored, kind, field));
            Assert.Equal(value, Style(Project(restored), kind)[field]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove(field);
            Assert.Null(NumericBodyProperty(Compile(removeAgain, restored).File.ToByteArray(), kind, field));
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape", "columnGap")]
    [InlineData("shape", "columns")]
    [InlineData("shape", "flatTextZ")]
    [InlineData("table", "columnGap")]
    [InlineData("table", "columns")]
    [InlineData("table", "flatTextZ")]
    public void NumericBodyPropertyOnlyAndCombinedRemovableStylesPreservePresence(string kind, string field)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style[field] = field == "flatTextZ" ? -42 : field == "columns" ? 3 : 12.5;
            if (withOtherProperties)
            {
                style["upright"] = false; style["rotation"] = 12;
                style[field == "columns" ? "columnGap" : "columns"] = field == "columns" ? 12.5 : 3;
                foreach (var otherField in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" })
                    style[otherField] = BodyPropertyValues(otherField).First();
            }
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(NumericBodyProperty(deleted, kind, field));
            Assert.Null(Body(deleted, kind).UpRight);
            if (withOtherProperties)
            {
                Assert.Null(Body(deleted, kind).Rotation);
                Assert.Null(Body(deleted, kind).ColumnCount);
                Assert.Null(Body(deleted, kind).ColumnSpacing);
                foreach (var otherField in new[] { "columnDirection", "verticalText", "wrap", "horizontalOverflow", "verticalOverflow", "verticalAlignment" })
                    AssertBodyPropertyValue(deleted, kind, otherField, null);
            }
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { [field] = field == "columns" ? 1 : 0 };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { [field] = field == "columns" ? 1 : 0 },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(field == "columns" ? 1 : 0, NumericBodyProperty(restored, kind, field));
            Assert.Equal(field == "columns" ? 1 : 0, Style(Project(restored), kind)[field]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
        }
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void MarginEdgesAndObjectRemovalPreserveOtherState(string kind)
    {
        var program = Program(kind);
        Style(program, kind)["margins"] = new JsonObject { ["left"] = 7, ["top"] = 8, ["right"] = 9, ["bottom"] = 10 };
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        foreach (var edge in new[] { "left", "top", "right", "bottom" })
        {
            var request = Project(source);
            Style(request, kind)["margins"]!.AsObject().Remove(edge);
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(Margin(deleted, kind, edge));
            Assert.False(Style(Project(deleted), kind)["margins"]!.AsObject().ContainsKey(edge));
            AssertOnlyBodyPropertyChanged(source, deleted, kind, "margins." + edge);
            foreach (var value in new[] { 0d, 12.5d, 10000d })
            {
                var restore = Project(deleted);
                Style(restore, kind)["margins"]![edge] = value;
                var restored = Compile(restore, deleted).File.ToByteArray();
                Assert.Equal(checked((int)(value * 12700)), Margin(restored, kind, edge));
                Assert.Equal(value, Style(Project(restored), kind)["margins"]![edge]!.GetValue<double>());
                AssertOnlyBodyPropertyChanged(deleted, restored, kind, "margins." + edge);
                var again = Project(restored);
                Style(again, kind)["margins"]!.AsObject().Remove(edge);
                Assert.Null(Margin(Compile(again, restored).File.ToByteArray(), kind, edge));
            }
        }
        foreach (var empty in new[] { false, true })
        {
            var request = Project(source);
            if (empty) Style(request, kind)["margins"] = new JsonObject();
            else Style(request, kind).Remove("margins");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.False(Style(Project(deleted), kind).ContainsKey("margins"));
            foreach (var edge in new[] { "left", "top", "right", "bottom" }) Assert.Null(Margin(deleted, kind, edge));
            AssertOnlyBodyPropertyChanged(source, deleted, kind, "margins");
            var restore = Project(deleted);
            Style(restore, kind)["margins"] = new JsonObject { ["left"] = 0 };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(0, Margin(restored, kind, "left"));
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "margins.left");
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void MarginStyleRemovalAndCompactRestoration(string kind)
    {
        foreach (var combined in new[] { false, true })
        {
            var program = Program(kind);
            if (!combined) Style(program, kind).Clear();
            Style(program, kind)["margins"] = new JsonObject { ["left"] = 0, ["top"] = 12.5, ["right"] = 9, ["bottom"] = 10 };
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            foreach (var edge in new[] { "left", "top", "right", "bottom" }) Assert.Null(Margin(deleted, kind, edge));
            Assert.Null(Body(deleted, kind).UpRight);
            Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).Wrap);
            if (!combined) AssertOnlyBodyPropertyChanged(source, deleted, kind, "margins");
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            var style = new JsonObject { ["margins"] = new JsonObject { ["bottom"] = 0 } };
            if (kind == "shape") element["textStyle"] = style;
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = style,
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(0, Margin(restored, kind, "bottom"));
            Assert.Equal(0, Style(Project(restored), kind)["margins"]!["bottom"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "margins.bottom");
        }
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void AutoFitChoicesAndPercentagesPreserveSourceLifecycle(string kind)
    {
        var program = Program(kind);
        Style(program, kind)["autoFit"] = "shrink-text";
        Style(program, kind)["normalAutoFit"] = new JsonObject { ["fontScale"] = 87.125, ["lineSpacingReduction"] = 12.5 };
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var invalid = Project(source);
        Style(invalid, kind).Remove("autoFit");
        Assert.Empty(Compile(invalid, source, success: false).File);
        foreach (var field in new[] { "fontScale", "lineSpacingReduction", "normalAutoFit" })
        {
            var request = Project(source);
            if (field == "normalAutoFit") Style(request, kind).Remove(field);
            else Style(request, kind)["normalAutoFit"]!.AsObject().Remove(field);
            var deleted = Compile(request, source).File.ToByteArray();
            var native = Assert.Single(Body(deleted, kind).Elements<A.NormalAutoFit>());
            if (field != "lineSpacingReduction") Assert.Null(native.FontScale);
            if (field != "fontScale") Assert.Null(native.LineSpaceReduction);
            var fresh = Style(Project(deleted), kind);
            Assert.Equal("shrink-text", fresh["autoFit"]!.GetValue<string>());
            if (field == "normalAutoFit") Assert.False(fresh.ContainsKey(field));
            else Assert.False(fresh["normalAutoFit"]!.AsObject().ContainsKey(field));
            AssertOnlyBodyPropertyChanged(source, deleted, kind, "normalAutoFit." + field);
            foreach (var value in field == "fontScale" ? new[] { 1d, 87.125d, 100d } : new[] { 0d, 12.5d, 13200d })
            {
                var restore = Project(deleted);
                Style(restore, kind)["normalAutoFit"] ??= new JsonObject();
                var restoredField = field == "normalAutoFit" ? "lineSpacingReduction" : field;
                Style(restore, kind)["normalAutoFit"]![restoredField] = value;
                var restored = Compile(restore, deleted).File.ToByteArray();
                var profile = Assert.Single(Body(restored, kind).Elements<A.NormalAutoFit>());
                Assert.Equal(checked((int)(value * 1000)), restoredField == "fontScale" ? profile.FontScale!.Value : profile.LineSpaceReduction!.Value);
                Assert.Equal(value, Style(Project(restored), kind)["normalAutoFit"]![restoredField]!.GetValue<double>());
                AssertOnlyBodyPropertyChanged(deleted, restored, kind, "normalAutoFit." + restoredField);
            }
        }
        var remove = Project(source);
        Style(remove, kind).Remove("autoFit"); Style(remove, kind).Remove("normalAutoFit");
        var absent = Compile(remove, source).File.ToByteArray();
        Assert.Empty(AutoFitChildren(Body(absent, kind)));
        Assert.False(Style(Project(absent), kind).ContainsKey("autoFit"));
        Assert.False(Style(Project(absent), kind).ContainsKey("normalAutoFit"));
        AssertOnlyBodyPropertyChanged(source, absent, kind, "autoFit");
        foreach (var mode in new[] { "none", "shrink-text", "resize-shape" })
        {
            var restore = Project(absent); Style(restore, kind)["autoFit"] = mode;
            var restored = Compile(restore, absent).File.ToByteArray();
            Assert.Equal(mode switch { "none" => "noAutofit", "shrink-text" => "normAutofit", _ => "spAutoFit" }, Assert.Single(AutoFitChildren(Body(restored, kind))).LocalName);
            Assert.Equal(mode, Style(Project(restored), kind)["autoFit"]!.GetValue<string>());
            Assert.Equal(restored, Compile(Project(restored), restored).File.ToByteArray());
            AssertOnlyBodyPropertyChanged(absent, restored, kind, "autoFit");
            var again = Project(restored); Style(again, kind).Remove("autoFit");
            Assert.Empty(AutoFitChildren(Body(Compile(again, restored).File.ToByteArray(), kind)));
            var switchMode = Project(source);
            Style(switchMode, kind)["autoFit"] = mode; Style(switchMode, kind).Remove("normalAutoFit");
            var switched = Compile(switchMode, source).File.ToByteArray();
            Assert.False(Style(Project(switched), kind).ContainsKey("normalAutoFit"));
            AssertOnlyBodyPropertyChanged(source, switched, kind, "autoFit");
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void AutoFitSimpleStyleRemovalAndRestoration(string kind)
    {
        foreach (var combined in new[] { false, true })
        {
            var program = Program(kind);
            if (!combined) Style(program, kind).Clear();
            Style(program, kind)["autoFit"] = "shrink-text";
            Style(program, kind)["normalAutoFit"] = new JsonObject { ["fontScale"] = 100, ["lineSpacingReduction"] = 0 };
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Empty(AutoFitChildren(Body(deleted, kind)));
            Assert.Null(Body(deleted, kind).UpRight); Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).Wrap); Assert.Null(Body(deleted, kind).LeftInset);
            if (!combined) AssertOnlyBodyPropertyChanged(source, deleted, kind, "autoFit");
            var restore = Project(deleted); var element = restore["pages"]![0]!["elements"]![0]!;
            var style = new JsonObject { ["autoFit"] = "shrink-text", ["normalAutoFit"] = new JsonObject { ["fontScale"] = 100, ["lineSpacingReduction"] = 0 } };
            if (kind == "shape") element["textStyle"] = style;
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = style,
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            var profile = Assert.Single(Body(restored, kind).Elements<A.NormalAutoFit>());
            Assert.Equal(100000, profile.FontScale!.Value); Assert.Equal(0, profile.LineSpaceReduction!.Value);
            Assert.Equal(0, Style(Project(restored), kind)["normalAutoFit"]!["lineSpacingReduction"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "autoFit");
        }
    }

    private static IEnumerable<OpenXmlElement> AutoFitChildren(A.BodyProperties body) =>
        body.ChildElements.Where(child => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit);

    private static int? Margin(byte[] bytes, string kind, string edge)
    {
        var body = Body(bytes, kind);
        return edge switch { "left" => body.LeftInset?.Value, "top" => body.TopInset?.Value,
            "right" => body.RightInset?.Value, "bottom" => body.BottomInset?.Value, _ => throw new ArgumentException(edge) };
    }

    private static long? NumericBodyProperty(byte[] bytes, string kind, string field) =>
        field == "flatTextZ" ? Body(bytes, kind).GetFirstChild<A.FlatText>()?.Z?.Value :
        field == "columns" ? Body(bytes, kind).ColumnCount?.Value : Body(bytes, kind).ColumnSpacing?.Value;

    [Theory]
    [InlineData("shape")]
    [InlineData("text")]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("table")]
    public void RotationSourceRemovalAndRestorationPreserveOtherState(string kind)
    {
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(Program(kind)).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(12, Style(request, kind)["rotation"]!.GetValue<double>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        Assert.Equal(12 * 60000, Body(source, kind).Rotation!.Value);
        Style(request, kind).Remove("rotation");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(Body(deleted, kind).Rotation);
        Assert.False(Style(Project(deleted), kind).ContainsKey("rotation"));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, "rotation");
        foreach (var value in new[] { 0, -24 })
        {
            var restore = Project(deleted);
            Style(restore, kind)["rotation"] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(value * 60000, Body(restored, kind).Rotation!.Value);
            Assert.Equal(value, Style(Project(restored), kind)["rotation"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "rotation");
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove("rotation");
            Assert.Null(Body(Compile(removeAgain, restored).File.ToByteArray(), kind).Rotation);
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("table")]
    public void RotationOnlyAndCombinedRemovableStylesPreservePresence(string kind)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style["rotation"] = 12;
            if (withOtherProperties) style["upright"] = false;
            var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            var request = Project(source);
            var owner = (kind == "shape" ? request["pages"]![0]!["elements"]![0] :
                request["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
            owner.Remove(kind == "shape" ? "textStyle" : "style");
            var deleted = Compile(request, source).File.ToByteArray();
            Assert.Null(Body(deleted, kind).Rotation);
            Assert.Null(Body(deleted, kind).UpRight);
            if (!withOtherProperties) AssertOnlyBodyPropertyChanged(source, deleted, kind, "rotation");
            var restore = Project(deleted);
            var element = restore["pages"]![0]!["elements"]![0]!;
            if (kind == "shape") element["textStyle"] = new JsonObject { ["rotation"] = 0 };
            else element["rows"]![0]!["cells"]![0]!["text"] = new JsonObject
            {
                ["style"] = new JsonObject { ["rotation"] = 0 },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }),
            };
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(0, Body(restored, kind).Rotation!.Value);
            Assert.Equal(0, Style(Project(restored), kind)["rotation"]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, "rotation");
        }
    }

    [Theory]
    [InlineData("shape", "upright")]
    [InlineData("shape", "anchorCenter")]
    [InlineData("shape", "forceAntiAlias")]
    [InlineData("shape", "spaceFirstLastParagraph")]
    [InlineData("shape", "compatibleLineSpacing")]
    [InlineData("shape", "fromWordArt")]
    [InlineData("table", "upright")]
    [InlineData("table", "anchorCenter")]
    [InlineData("table", "forceAntiAlias")]
    [InlineData("table", "spaceFirstLastParagraph")]
    [InlineData("table", "compatibleLineSpacing")]
    [InlineData("table", "fromWordArt")]
    public void BooleanBodyStyleCanBeRemovedWithExistingAuthority(string kind, string field)
    {
        var program = Program(kind);
        Style(program, kind).Clear(); Style(program, kind)[field] = true;
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var request = Project(source);
        Assert.Equal(new[] { field }, Style(request, kind).Select(p => p.Key));
        JsonObject StyleOwner(JsonObject input) => (kind == "shape" ? input["pages"]![0]!["elements"]![0] :
            input["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"])!.AsObject();
        StyleOwner(request).Remove(kind == "shape" ? "textStyle" : "style");
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(BooleanBodyProperty(deleted, kind, field));
        var fresh = Project(deleted);
        if (kind == "shape") Assert.False(StyleOwner(fresh).ContainsKey("textStyle"));
        else
        {
            var cell = fresh["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!;
            Assert.Equal("Retain this text", cell["text"]!.GetValue<string>());
            cell["text"] = new JsonObject { ["style"] = new JsonObject { [field] = false },
                ["paragraphs"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["text"] = "Retain this text" }) }) };
            var restored = Compile(fresh, deleted).File.ToByteArray();
            Assert.False(BooleanBodyProperty(restored, kind, field)!.Value);
            Assert.False(Style(Project(restored), kind)[field]!.GetValue<bool>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
        }
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);

        var denied = Project(source);
        StyleOwner(denied).Remove(kind == "shape" ? "textStyle" : "style");
        var capabilities = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray();
        capabilities.Remove(capabilities.Single(c => c!["operation"]!.GetValue<string>() == (kind == "shape" ? "setTextBodyStyle" : "setTableCellStyle")));
        Assert.Empty(Compile(denied, source, success: false).File);

    }

    [Theory]
    [InlineData("shape", "upright")]
    [InlineData("shape", "anchorCenter")]
    [InlineData("shape", "forceAntiAlias")]
    [InlineData("shape", "spaceFirstLastParagraph")]
    [InlineData("shape", "compatibleLineSpacing")]
    [InlineData("shape", "fromWordArt")]
    [InlineData("text", "upright")]
    [InlineData("text", "anchorCenter")]
    [InlineData("text", "forceAntiAlias")]
    [InlineData("text", "spaceFirstLastParagraph")]
    [InlineData("text", "compatibleLineSpacing")]
    [InlineData("text", "fromWordArt")]
    [InlineData("master", "upright")]
    [InlineData("master", "anchorCenter")]
    [InlineData("master", "forceAntiAlias")]
    [InlineData("master", "spaceFirstLastParagraph")]
    [InlineData("master", "compatibleLineSpacing")]
    [InlineData("master", "fromWordArt")]
    [InlineData("layout", "upright")]
    [InlineData("layout", "anchorCenter")]
    [InlineData("layout", "forceAntiAlias")]
    [InlineData("layout", "spaceFirstLastParagraph")]
    [InlineData("layout", "compatibleLineSpacing")]
    [InlineData("layout", "fromWordArt")]
    [InlineData("table", "upright")]
    [InlineData("table", "anchorCenter")]
    [InlineData("table", "forceAntiAlias")]
    [InlineData("table", "spaceFirstLastParagraph")]
    [InlineData("table", "compatibleLineSpacing")]
    [InlineData("table", "fromWordArt")]
    public void BooleanBodySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind);
        Style(program, kind)[field] = true;
        if (field == "fromWordArt") Style(program, kind)["textWarpPreset"] = "textArchUp";
        if (field is "forceAntiAlias" or "spaceFirstLastParagraph" or "compatibleLineSpacing") Style(program, kind)["anchorCenter"] = false;
        if (field is "spaceFirstLastParagraph" or "compatibleLineSpacing")
        {
            var paragraphs = JsonNode.Parse("""[{"style":{"spaceBefore":7,"spaceAfter":9,"lineSpacing":1.25},"runs":[{"text":"Retain this text"}]}]""");
            var owner = Style(program, kind).Parent!.AsObject();
            if (kind == "table") owner["paragraphs"] = paragraphs;
            else owner["text"] = new JsonObject { ["paragraphs"] = paragraphs };
        }
        Style(program, kind)["verticalAlignment"] = "middle";
        var authored = Compile(program);
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored.File.ToByteArray());
        var original = source.ToArray();
        var projected = Project(source);
        Assert.True(Style(projected, kind)[field]!.GetValue<bool>());
        Assert.Equal(source, Compile(projected, source).File.ToByteArray());
        Assert.True(BooleanBodyProperty(source, kind, field)!.Value);
        var request = Project(source);
        Style(request, kind).Remove(field);
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(BooleanBodyProperty(deleted, kind, field));
        Assert.False(Style(Project(deleted), kind).ContainsKey(field));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
        foreach (var value in new[] { false, true })
        {
            var restore = Project(deleted);
            Style(restore, kind)[field] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(value, BooleanBodyProperty(restored, kind, field)!.Value);
            Assert.Equal(value, Style(Project(restored), kind)[field]!.GetValue<bool>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove(field);
            Assert.Null(BooleanBodyProperty(Compile(removeAgain, restored).File.ToByteArray(), kind, field));
        }
        Assert.Equal(original, source);
    }

    private static bool? BooleanBodyProperty(byte[] bytes, string kind, string field) =>
        field switch { "anchorCenter" => Body(bytes, kind).AnchorCenter?.Value,
            "fromWordArt" => Body(bytes, kind).FromWordArt?.Value,
            "compatibleLineSpacing" => Body(bytes, kind).CompatibleLineSpacing?.Value,
            "spaceFirstLastParagraph" => Body(bytes, kind).UseParagraphSpacing?.Value,
            "forceAntiAlias" => Body(bytes, kind).ForceAntiAlias?.Value, _ => Body(bytes, kind).UpRight?.Value };

    private static string[] BodyPropertyValues(string field) => field switch
    {
        "verticalText" => ["horizontal", "vertical", "vertical270"],
        "columnDirection" => ["left-to-right", "right-to-left"],
        "wrap" => ["none", "square"],
        "horizontalOverflow" => ["overflow", "clip"],
        "verticalOverflow" => ["overflow", "ellipsis", "clip"],
        "verticalAlignment" => ["top", "middle", "bottom"],
        _ => throw new ArgumentException(field),
    };

    private static void AssertBodyPropertyValue(byte[] bytes, string kind, string field, string? value)
    {
        var body = Body(bytes, kind);
        if (field == "verticalAlignment")
        {
            if (value is null) Assert.Null(body.Anchor);
            else Assert.Equal(value switch
            {
                "top" => A.TextAnchoringTypeValues.Top,
                "middle" => A.TextAnchoringTypeValues.Center,
                "bottom" => A.TextAnchoringTypeValues.Bottom,
                _ => throw new ArgumentException(value),
            }, body.Anchor!.Value);
        }
        else if (field == "verticalOverflow")
        {
            if (value is null) Assert.Null(body.VerticalOverflow);
            else Assert.Equal(value switch
            {
                "overflow" => A.TextVerticalOverflowValues.Overflow,
                "ellipsis" => A.TextVerticalOverflowValues.Ellipsis,
                "clip" => A.TextVerticalOverflowValues.Clip,
                _ => throw new ArgumentException(value),
            }, body.VerticalOverflow!.Value);
        }
        else if (field == "horizontalOverflow")
        {
            if (value is null) Assert.Null(body.HorizontalOverflow);
            else Assert.Equal(value == "clip" ? A.TextHorizontalOverflowValues.Clip : A.TextHorizontalOverflowValues.Overflow, body.HorizontalOverflow!.Value);
        }
        else if (field == "wrap")
        {
            if (value is null) Assert.Null(body.Wrap);
            else Assert.Equal(value == "square" ? A.TextWrappingValues.Square : A.TextWrappingValues.None, body.Wrap!.Value);
        }
        else if (field == "columnDirection")
        {
            if (value is null) Assert.Null(body.RightToLeftColumns);
            else Assert.Equal(value == "right-to-left", body.RightToLeftColumns!.Value);
        }
        else
        {
            if (value is null) Assert.Null(body.Vertical);
            else Assert.Equal(value switch
            {
                "horizontal" => A.TextVerticalValues.Horizontal,
                "vertical" => A.TextVerticalValues.Vertical,
                "vertical270" => A.TextVerticalValues.Vertical270,
                _ => throw new ArgumentException(value),
            }, body.Vertical!.Value);
        }
    }

    private static JsonObject Program(string kind)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "package.json"))) directory = directory.Parent;
        var program = JsonNode.Parse(File.ReadAllBytes(Path.Combine(directory!.FullName, "examples/ppj/minimum.ppj")))!.AsObject();
        var element = JsonNode.Parse("""
            {"id":"owner","type":"shape","frame":{"x":30,"y":40,"width":200,"height":100},
             "geometry":{"kind":"preset","preset":"rect"},"text":"Retain this text",
             "textStyle":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}}}
            """)!.AsObject();
        if (kind == "table")
        {
            element = JsonNode.Parse("""
                {"id":"owner","type":"table","frame":{"x":30,"y":40,"width":200,"height":100},
                 "columns":[{"id":"column","width":200}],"rows":[{"id":"row","height":100,"cells":[{"id":"cell",
                 "text":{"style":{"upright":true,"rotation":12,"wrap":"square","margins":{"left":7}},
                 "paragraphs":[{"runs":[{"text":"Retain this text"}]}]}}]}]}
                """)!.AsObject();
        }
        else if (kind == "text")
        {
            element["type"] = kind;
            element["style"] = element["textStyle"]!.DeepClone(); element.Remove("textStyle"); element.Remove("geometry");
        }
        program["pages"]![0]!["elements"] = new JsonArray(element);
        if (kind is "master" or "layout")
        {
            program["design"]!["masters"] = JsonNode.Parse("""[{"id":"master","name":"Master","placeholders":[]}]""");
            program["design"]!["layouts"] = JsonNode.Parse("""[{"id":"layout","name":"Layout","master":"master","layoutType":"blank","placeholders":[]}]""");
            var placeholder = new JsonObject { ["id"] = "placeholder", ["name"] = "Owner", ["placeholderType"] = "body", ["index"] = 0,
                ["frame"] = element["frame"]!.DeepClone(), ["text"] = element["text"]!.DeepClone(), ["style"] = element["textStyle"]!.DeepClone() };
            program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"] = new JsonArray(placeholder);
            program["pages"]![0]!["layout"] = "layout";
            program["pages"]![0]!["elements"] = new JsonArray();
        }
        return program;
    }

    private static JsonObject Style(JsonObject program, string kind)
    {
        if (kind is "master" or "layout") return program["design"]![kind == "master" ? "masters" : "layouts"]![0]!["placeholders"]![0]!["style"]!.AsObject();
        var element = program["pages"]![0]!["elements"]![0]!;
        return (kind == "table" ? element["rows"]![0]!["cells"]![0]!["text"]!["style"] : element[kind == "shape" ? "textStyle" : "style"])!.AsObject();
    }
    private static CodecResponse Compile(JsonObject program, byte[]? source = null, bool success = true) => Invoke(new CodecRequest
    {
        Operation = CodecOperation.CompilePpjToPptx, File = source is null ? ByteString.Empty : ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { ProgramJson = ByteString.CopyFromUtf8(program.ToJsonString()) },
    }, success);
    private static JsonObject Project(byte[] source) => JsonNode.Parse(Invoke(new CodecRequest
    {
        Operation = CodecOperation.ProjectPptxToPpj, File = ByteString.CopyFrom(source),
        PresentationProgram = new PresentationProgramRequest { SourceUri = "source.pptx" },
    }).PresentationProgram.ProgramJson.ToByteArray())!.AsObject();
    private static CodecResponse Invoke(CodecRequest request, bool success = true)
    {
        request.ProtocolVersion = CodecProtocol.ProtocolVersion; request.Family = ArtifactFamily.Presentation;
        var bytes = request.ToByteArray(); var response = PpjCodecProtocol.InvokeResponse(ref bytes, null);
        Assert.True(response.Ok == success, string.Join("\n", response.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        return response;
    }
    private static OpenXmlPartRootElement Owner(PresentationDocument document, string kind) => kind switch
    {
        "master" => document.PresentationPart!.SlideMasterParts.Single().SlideMaster,
        "layout" => document.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single().SlideLayout,
        _ => document.PresentationPart!.SlideParts.Single().Slide,
    };
    private static A.BodyProperties Body(byte[] bytes, string kind)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return (A.BodyProperties)Assert.Single(Owner(document, kind).Descendants<A.BodyProperties>()).CloneNode(true);
    }
    private static void AssertOnlyBodyPropertyChanged(byte[] before, byte[] after, string kind, string field = "upright")
    {
        using var left = PresentationDocument.Open(new MemoryStream(before), false);
        using var right = PresentationDocument.Open(new MemoryStream(after), false);
        var oldSlide = Owner(left, kind);
        var newSlide = Owner(right, kind);
        if (field.StartsWith("paragraphDefault."))
        {
            foreach (var owner in new[] { oldSlide, newSlide })
            {
                var paragraphProperties = owner.Descendants<A.Paragraph>().First().ParagraphProperties;
                var defaults = paragraphProperties?.GetFirstChild<A.DefaultRunProperties>();
                if (defaults is not null)
                {
                    if (field == "paragraphDefault.bold") defaults.Bold = null;
                    else if (field == "paragraphDefault.italic") defaults.Italic = null;
                    else if (field == "paragraphDefault.size") defaults.FontSize = null;
                    else if (field == "paragraphDefault.fontFamily") defaults.GetFirstChild<A.LatinFont>()?.Remove();
                    else if (field == "paragraphDefault.fontFamilyEastAsia") defaults.GetFirstChild<A.EastAsianFont>()?.Remove();
                    else if (field == "paragraphDefault.fontFamilyComplexScript") defaults.GetFirstChild<A.ComplexScriptFont>()?.Remove();
                    else if (field == "paragraphDefault.language") defaults.Language = null;
                    else if (field == "paragraphDefault.kerning") defaults.Kerning = null;
                    else if (field == "paragraphDefault.letterSpacing") defaults.Spacing = null;
                    else if (field == "paragraphDefault.baseline") defaults.Baseline = null;
                    else if (field == "paragraphDefault.capitalization") defaults.Capital = null;
                    else if (field == "paragraphDefault.strike") defaults.Strike = null;
                    else if (field == "paragraphDefault.underline") defaults.Underline = null;
                    else if (field == "paragraphDefault.highlight") defaults.GetFirstChild<A.Highlight>()?.Remove();
                    else if (field == "paragraphDefault.color") defaults.GetFirstChild<A.SolidFill>()?.Remove();
                    else if (field == "paragraphDefault.gradient") defaults.GetFirstChild<A.GradientFill>()?.Remove();
                    else if (field == "paragraphDefault.paint")
                    {
                        defaults.GetFirstChild<A.SolidFill>()?.Remove();
                        defaults.GetFirstChild<A.GradientFill>()?.Remove();
                    }
                    if (defaults.GetAttributes().Count == 0 && defaults.ChildElements.Count == 0) defaults.Remove();
                }
                if (paragraphProperties is not null && paragraphProperties.GetAttributes().Count == 0 && paragraphProperties.ChildElements.Count == 0)
                    paragraphProperties.Remove();
            }
        }
        else if (field == "textWarpPreset")
        {
            foreach (var child in oldSlide.Descendants<A.PresetTextWarp>().ToArray()) child.Remove();
            foreach (var child in newSlide.Descendants<A.PresetTextWarp>().ToArray()) child.Remove();
        }
        else if (field == "textWarpAdjustments")
        {
            foreach (var child in oldSlide.Descendants<A.PresetTextWarp>().SelectMany(w => w.Elements<A.AdjustValueList>()).ToArray()) child.Remove();
            foreach (var child in newSlide.Descendants<A.PresetTextWarp>().SelectMany(w => w.Elements<A.AdjustValueList>()).ToArray()) child.Remove();
        }
        else if (field == "autoFit")
        {
            foreach (var child in AutoFitChildren(Assert.Single(oldSlide.Descendants<A.BodyProperties>())).ToArray()) child.Remove();
            foreach (var child in AutoFitChildren(Assert.Single(newSlide.Descendants<A.BodyProperties>())).ToArray()) child.Remove();
        }
        else if (field.StartsWith("normalAutoFit."))
        {
            foreach (var profile in new[] { Assert.Single(oldSlide.Descendants<A.NormalAutoFit>()), Assert.Single(newSlide.Descendants<A.NormalAutoFit>()) })
            {
                if (field != "normalAutoFit.lineSpacingReduction") profile.FontScale = null;
                if (field != "normalAutoFit.fontScale") profile.LineSpaceReduction = null;
            }
        }
        else if (field == "margins" || field.StartsWith("margins."))
        {
            foreach (var edge in new[] { "left", "top", "right", "bottom" }.Where(edge => field == "margins" || field == "margins." + edge))
            {
                var attribute = edge[0] + "Ins";
                Assert.Single(oldSlide.Descendants<A.BodyProperties>()).RemoveAttribute(attribute, "");
                Assert.Single(newSlide.Descendants<A.BodyProperties>()).RemoveAttribute(attribute, "");
            }
        }
        else if (field == "anchorCenter")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).AnchorCenter = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).AnchorCenter = null;
        }
        else if (field == "forceAntiAlias")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).ForceAntiAlias = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).ForceAntiAlias = null;
        }
        else if (field == "spaceFirstLastParagraph")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).UseParagraphSpacing = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).UseParagraphSpacing = null;
        }
        else if (field == "compatibleLineSpacing")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).CompatibleLineSpacing = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).CompatibleLineSpacing = null;
        }
        else if (field == "flatTextZ")
        {
            foreach (var child in oldSlide.Descendants<A.FlatText>().ToArray()) child.Remove();
            foreach (var child in newSlide.Descendants<A.FlatText>().ToArray()) child.Remove();
        }
        else if (field == "fromWordArt")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).FromWordArt = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).FromWordArt = null;
        }
        else if (field == "columnGap")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).ColumnSpacing = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).ColumnSpacing = null;
        }
        else if (field == "columns")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).ColumnCount = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).ColumnCount = null;
        }
        else if (field == "verticalAlignment")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Anchor = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Anchor = null;
        }
        else if (field == "verticalOverflow")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).VerticalOverflow = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).VerticalOverflow = null;
        }
        else if (field == "horizontalOverflow")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).HorizontalOverflow = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).HorizontalOverflow = null;
        }
        else if (field == "wrap")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Wrap = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Wrap = null;
        }
        else if (field == "verticalText")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Vertical = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Vertical = null;
        }
        else if (field == "columnDirection")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).RightToLeftColumns = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).RightToLeftColumns = null;
        }
        else if (field == "rotation")
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).Rotation = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).Rotation = null;
        }
        else
        {
            Assert.Single(oldSlide.Descendants<A.BodyProperties>()).UpRight = null;
            Assert.Single(newSlide.Descendants<A.BodyProperties>()).UpRight = null;
        }
        XElement Content(string xml)
        {
            var node = XElement.Parse(xml);
            foreach (var attribute in node.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).ToArray()) attribute.Remove();
            return node;
        }
        // The table writer may repeat namespace declarations. Compare expanded
        // XML names/values while keeping all non-namespace attributes/content.
        Assert.True(XNode.DeepEquals(Content(oldSlide.OuterXml), Content(newSlide.OuterXml)));
        using var oldZip = new ZipArchive(new MemoryStream(before));
        using var newZip = new ZipArchive(new MemoryStream(after));
        Assert.Equal(oldZip.Entries.Select(e => e.FullName).Order(), newZip.Entries.Select(e => e.FullName).Order());
        var ownerPart = kind switch { "master" => "ppt/slideMasters/slideMaster1.xml", "layout" => "ppt/slideLayouts/slideLayout1.xml", _ => "ppt/slides/slide1.xml" };
        foreach (var entry in oldZip.Entries.Where(e => e.FullName != ownerPart))
        {
            using var a = entry.Open(); using var b = newZip.GetEntry(entry.FullName)!.Open();
            using var aa = new MemoryStream(); using var bb = new MemoryStream(); a.CopyTo(aa); b.CopyTo(bb);
            Assert.Equal(aa.ToArray(), bb.ToArray());
        }
    }
}
