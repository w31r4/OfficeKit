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

    [Fact]
    public void StaticTableTextFieldTypeLifecyclePatchesOnlyFieldTypeToken()
    {
        var program = Program("table");
        program["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"] = JsonNode.Parse("""
            {"paragraphs":[{"runs":[{"field":{"id":"{11111111-2222-4333-8444-555555555555}","type":"customStatic","text":"cached"}}]}]}
            """);
        var authored = Compile(program).File.ToByteArray();
        var source = PptxCodecTests.RemoveEmbeddedPpj(authored);
        var originalSource = source.ToArray();

        var projected = Project(source);
        var table = projected["pages"]![0]!["elements"]![0]!.AsObject();
        var cell = table["rows"]![0]!["cells"]![0]!.AsObject();
        var field = cell["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        var fieldLeaf = table["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "tableTextFieldType")!.AsObject();
        Assert.Equal("customStatic", fieldLeaf["value"]!.GetValue<string>());
        var capability = table["nativeRef"]!["capabilities"]!.AsArray()
            .Single(item => item!["operation"]!.GetValue<string>() == "setTextField");
        Assert.Contains("table.rows[].cells[].text.paragraphs[].runs[].field.type",
            capability!["fields"]!.AsArray().Select(item => item!.GetValue<string>()));

        field["type"] = "customStatic2";
        var edited = Compile(projected, source);
        Assert.True(edited.Ok, string.Join("\n", edited.Diagnostics.Select(item => item.Code + ": " + item.Message)));
        Assert.Equal(["ppt/slides/slide1.xml"], edited.PresentationProgram.ChangedParts);
        Assert.Equal(originalSource, source);
        var output = edited.File.ToByteArray();
        var outputField = ReadTableField(output);
        Assert.Equal("customStatic2", outputField.Type!.Value);
        Assert.Equal("{11111111-2222-4333-8444-555555555555}", outputField.Id!.Value);
        Assert.Equal("cached", outputField.Text!.Text);
        using (var document = PresentationDocument.Open(new MemoryStream(output), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document));

        var reprojected = Project(output);
        var reprojectedField = reprojected["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!;
        Assert.Equal("customStatic2", reprojectedField!["type"]!.GetValue<string>());
        Assert.Equal("cached", reprojectedField["text"]!.GetValue<string>());

        var automatic = Project(source);
        automatic["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["type"] = "slidenum";
        Assert.False(Compile(automatic, source, success: false).Ok);

        var identity = Project(source);
        identity["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["id"] = "{22222222-3333-4444-8555-666666666666}";
        Assert.False(Compile(identity, source, success: false).Ok);

        var combined = Project(source);
        var combinedField = combined["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        combinedField["type"] = "customStatic2";
        combinedField["text"] = "changed";
        Assert.False(Compile(combined, source, success: false).Ok);

        var stale = Project(source);
        var staleField = stale["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!.AsObject();
        staleField["type"] = "customStatic2";
        stale["pages"]![0]!["elements"]![0]!["nativeRef"]!["leaves"]!.AsArray()
            .Single(leaf => leaf!["kind"]!.GetValue<string>() == "tableTextFieldType")!["expectedHash"] = new string('0', 64);
        Assert.False(Compile(stale, source, success: false).Ok);

        var invalid = Project(source);
        invalid["pages"]![0]!["elements"]![0]!["rows"]![0]!["cells"]![0]!["text"]!["paragraphs"]![0]!["runs"]![0]!["field"]!["type"] = "";
        Assert.False(Compile(invalid, source, success: false).Ok);
        Assert.Equal(originalSource, source);
    }

    private static A.Field ReadField(byte[] bytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        return Assert.Single(document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Field>()).CloneNode(true) as A.Field
            ?? throw new InvalidOperationException("Expected a field.");
    }

    private static A.Field ReadTableField(byte[] bytes)
    {
        using var document = PresentationDocument.Open(new MemoryStream(bytes), false);
        var table = document.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>().Single();
        return Assert.Single(table.Elements<A.TableRow>().Single().Elements<A.TableCell>().First().Descendants<A.Field>())
            .CloneNode(true) as A.Field
            ?? throw new InvalidOperationException("Expected a table field.");
    }
}
