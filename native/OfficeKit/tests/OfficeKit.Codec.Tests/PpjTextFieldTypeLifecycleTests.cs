using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Fact]
    public void StaticTextFieldTypeLifecyclePatchesOnlyFieldTypeToken()
    {
        var program = Program("text");
        program["pages"]![0]!["elements"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"runs":[{"field":{"id":"{11111111-2222-4333-8444-555555555555}","type":"customStatic","text":"cached"}}]}]}
            """);
        var authored = Compile(program).File.ToByteArray();
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored);
        var originalSource = source.ToArray();

        var projected = Project(source);
        var element = projected["pages"]![0]!["elements"]![0]!.AsObject();
        var field = element["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        var fieldLeaf = element["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "textFieldType")!.AsObject();
        Assert.Equal("customStatic", fieldLeaf["value"]!.GetValue<string>());
        var capability = element["nativeRef"]!["capabilities"]!.AsArray()
            .Single(item => item!["operation"]!.GetValue<string>() == "setTextField");
        Assert.Contains("text.paragraphs[].runs[].field.type", capability!["fields"]!.AsArray().Select(item => item!.GetValue<string>()));

        field["type"] = "customStatic2";
        var edited = Compile(projected, source);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => item.Code + ": " + item.Message)));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        Assert.Equal("customStatic", ReadField(source).Type!.Value);
        var output = edited.File.ToByteArray();
        var outputField = ReadField(output);
        Assert.Equal("customStatic2", outputField.Type!.Value);
        Assert.Equal("{11111111-2222-4333-8444-555555555555}", outputField.Id!.Value);
        Assert.Equal("cached", outputField.Text!.Text);
        using (var document = PresentationDocument.Open(new MemoryStream(output), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var reprojected = Project(output);
        var reprojectedField = reprojected["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!;
        Assert.Equal("customStatic2", reprojectedField!["type"]!.GetValue<string>());
        Assert.Equal("cached", reprojectedField["text"]!.GetValue<string>());

        var denied = Project(source);
        denied["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["type"] = "customStatic2";
        var deniedCapability = denied["pages"]![0]!["elements"]![0]!["nativeRef"]!["capabilities"]!.AsArray()
            .Single(item => item!["operation"]!.GetValue<string>() == "setTextField");
        deniedCapability!["fields"]!.AsArray().Clear();
        Assert.False(Compile(denied, source, success: false).Ok);
        Assert.Equal(originalSource, source);

        var automatic = Project(source);
        automatic["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["type"] = "slidenum";
        Assert.False(Compile(automatic, source, success: false).Ok);

        var identity = Project(source);
        identity["pages"]![0]!["elements"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["id"] = "{22222222-3333-4444-8555-666666666666}";
        Assert.False(Compile(identity, source, success: false).Ok);
        Assert.Equal(originalSource, source);
    }

    private static A.Field ReadField(byte[] bytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return Assert.Single(document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Field>()).CloneNode(true) as A.Field
            ?? throw new InvalidOperationException("Expected a field.");
    }
}
