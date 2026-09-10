using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
    public void ParagraphRightIndentLifecyclePreservesSource(string kind)
    {
        var program = RightIndentProgram(kind);
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        AssertRightIndent(authored, kind, 18.25);
        using var stream = new MemoryStream(); stream.Write(authored);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var paragraphs = Owner(document, kind).Descendants<A.Paragraph>().ToArray();
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("marR", "", "+00231775"));
            paragraphs[1].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("marR", "", "+000000"));
            paragraphs[0].ParagraphProperties!.SetAttribute(new OpenXmlAttribute("future", "keep", "urn:officekit:test", "retained"));
        }
        var source = stream.ToArray(); var original = source.ToArray();
        Assert.Equal(source, Compile(Project(source), source).File.ToByteArray());
        var before = source;
        foreach (var value in new double?[] { 0, 25.12345, null, 40, 0 })
        {
            var request = Project(before); SetRightIndent(request, value); var json = request.ToJsonString();
            var candidate = Compile(request, before).File.ToByteArray();
            AssertRightIndent(candidate, kind, value);
            AssertOnlyBodyPropertyChanged(before, candidate, kind, "paragraph.rightIndent");
            Assert.Equal(candidate, Compile(Project(candidate), candidate).File.ToByteArray());
            Assert.Equal(json, request.ToJsonString()); before = candidate;
        }
        var remove = Project(source); SetRightIndent(remove, null);
        var removed = Compile(remove, source).File.ToByteArray();
        AssertRightIndent(removed, kind, null);
        AssertOnlyBodyPropertyChanged(source, removed, kind, "paragraph.rightIndent");
        var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
        AssertOnlyBodyPropertyChanged(source, Compile(unrelated, source).File.ToByteArray(), kind, "paragraphDefault.bold");
        var denied = Project(source); SetRightIndent(denied, 0);
        var fields = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(c => c!["operation"]!.GetValue<string>() == "setTextParagraphStyle")!["fields"]!.AsArray();
        fields.Remove(fields.Single(f => f!.GetValue<string>() == "text.paragraphs[].style.rightIndent"));
        Assert.Empty(Compile(denied, source, success: false).File);
        Assert.Equal(original, source);

        var minimal = Program(kind);
        minimal["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"style":{"rightIndent":0},"runs":[{"text":""}]}]}
            """);
        var onlyRight = PptxCodecTests.RemoveEmbeddedPpj(Compile(minimal).File.ToByteArray());
        AssertRightIndent(onlyRight, kind, 0);
        var removeStyle = Project(onlyRight); FirstTextParagraph(removeStyle).Remove("style");
        AssertRightIndent(Compile(removeStyle, onlyRight).File.ToByteArray(), kind, null);
    }

    [Fact]
    public void ParagraphRightIndentBoundsPrecisionAndPrecedence()
    {
        var program = RightIndentProgram("text");
        foreach (var value in new[] { 0, 0.5 / 12700, 1.5 / 12700, 25.12345, 4032 })
        {
            SetRightIndent(program, value);
            var bytes = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
            AssertRightIndent(bytes, "text", value);
            using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));
        }
        var element = program["pages"]![0]!["elements"]![0]!;
        element["style"]!["paragraph"] = new JsonObject { ["rightIndent"] = 50 };
        SetRightIndent(program, 0);
        AssertRightIndent(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 0);
        SetRightIndent(program, null);
        AssertRightIndent(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 50);
        element["text"]!["style"] = new JsonObject { ["paragraph"] = new JsonObject { ["rightIndent"] = 25.5 } };
        AssertRightIndent(PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray()), "text", 25.5);
        foreach (var invalid in new[] { "null", "false", "\"12\"", "-0.00001", "4032.00001", "{}" })
        {
            FirstTextParagraph(program)["style"]!["rightIndent"] = JsonNode.Parse(invalid);
            Assert.Empty(Compile(program, success: false).File);
        }
    }

    [Fact]
    public void ParagraphRightIndentRetainsUnmodeledSource()
    {
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(RightIndentProgram("text")).File.ToByteArray());
        foreach (var token in new[] { "-1", "51206401", "1.5", "invalid", "9223372036854775808", "" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
                Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!
                    .SetAttribute(new OpenXmlAttribute("marR", "", token));
            var source = stream.ToArray(); var projected = Project(source);
            Assert.Null(FirstTextParagraph(projected)["style"]?["rightIndent"]);
            Assert.Equal(source, Compile(projected, source).File.ToByteArray());
            SetRightIndent(projected, 0);
            Assert.Empty(Compile(projected, source, success: false).File);
            var unrelated = Project(source); FirstTextParagraph(unrelated)["style"]!["defaultText"]!["bold"] = false;
            var candidate = Compile(unrelated, source).File.ToByteArray();
            AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraphDefault.bold");
            Assert.Null(FirstTextParagraph(Project(candidate))["style"]?["rightIndent"]);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphRightIndentNativeRemovalValidatesSource(bool grouped)
    {
        var program = RightIndentProgram("text");
        if (grouped)
        {
            var elements = program["pages"]![0]!["elements"]!.AsArray(); var child = elements[0]!.DeepClone();
            elements[0] = new JsonObject { ["id"] = "right-indent-group", ["type"] = "group",
                ["frame"] = child["frame"]!.DeepClone(), ["elements"] = new JsonArray(child) };
        }
        var authored = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        foreach (var profile in new[] { "modeled", "absent", "unknown" })
        {
            using var stream = new MemoryStream(); stream.Write(authored);
            using (var document = PresentationDocument.Open(stream, true))
            {
                var properties = Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!;
                if (profile == "absent") properties.RightMargin = null;
                if (profile == "unknown") properties.SetAttribute(new OpenXmlAttribute("marR", "", "invalid"));
            }
            var source = stream.ToArray();
            var imported = Call(new CodecRequest { ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ImportPptx, Family = ArtifactFamily.Presentation,
                File = ByteString.CopyFrom(source) });
            Assert.True(imported.Ok);
            var element = imported.Artifact.Presentation.Slides[0].Elements[0];
            if (grouped) element = element.Group.Children[0];
            element.Shape.TextBody.Paragraphs[0].NoMarginRight = true;
            var exported = Call(new CodecRequest { ProtocolVersion = CodecProtocol.ProtocolVersion,
                Operation = CodecOperation.ExportPptx, Family = ArtifactFamily.Presentation,
                Artifact = imported.Artifact });
            if (profile == "unknown")
            {
                Assert.False(exported.Ok); Assert.Empty(exported.File);
                Assert.Contains(exported.Diagnostics, d => d.Code == "unsupported_presentation_edit");
            }
            else
            {
                Assert.True(exported.Ok);
                var candidate = exported.File.ToByteArray();
                using var document = PresentationDocument.Open(new MemoryStream(candidate), false);
                Assert.Null(Owner(document, "text").Descendants<A.Paragraph>().First().ParagraphProperties!.RightMargin);
                var projected = Project(candidate);
                var projectedElement = projected["pages"]![0]!["elements"]![0]!;
                if (grouped) projectedElement = projectedElement["elements"]![0]!;
                Assert.Null(projectedElement["text"]!["paragraphs"]![0]!["style"]?["rightIndent"]);
                AssertOnlyBodyPropertyChanged(source, candidate, "text", "paragraph.rightIndent");
            }
        }
        static CodecResponse Call(CodecRequest request) =>
            CodecResponse.Parser.ParseFrom(CodecProtocol.Invoke(request.ToByteArray()));
    }

    private static JsonObject RightIndentProgram(string kind)
    {
        var program = DirectionProgram(kind);
        var style = FirstTextParagraph(program)["style"]!;
        style["rightIndent"] = 18.25; style["indent"] = 12; style["hanging"] = 4;
        program["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![1]!["style"]!["rightIndent"] = 0;
        return program;
    }

    private static void SetRightIndent(JsonObject program, double? value)
    {
        var paragraph = FirstTextParagraph(program); paragraph["style"] ??= new JsonObject();
        if (value is null) paragraph["style"]!.AsObject().Remove("rightIndent");
        else paragraph["style"]!["rightIndent"] = value.Value;
    }

    private static void AssertRightIndent(byte[] bytes, string kind, double? expected)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var margin = Owner(document, kind).Descendants<A.Paragraph>().First().ParagraphProperties?.RightMargin;
        var projected = FirstTextParagraph(Project(bytes))["style"]?["rightIndent"];
        if (expected is null) { Assert.Null(margin); Assert.Null(projected); return; }
        var emu = (long)Math.Round(expected.Value * 12700);
        Assert.NotNull(margin); Assert.Equal(emu, margin.Value);
        Assert.NotNull(projected); Assert.Equal(emu, (long)Math.Round(projected.GetValue<double>() * 12700));
    }
}
