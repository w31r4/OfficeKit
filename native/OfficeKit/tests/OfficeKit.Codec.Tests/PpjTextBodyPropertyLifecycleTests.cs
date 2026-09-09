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
    [Fact]
    public void AnchorCenterDeletionWireIntentIsUnambiguous()
    {
        foreach (var value in new[] { false, true })
        {
            var properties = new PresentationTextBodyProperties { AnchorCenter = value };
            Assert.Equal(new byte[] { 0x80, 0x02, value ? (byte)1 : (byte)0 }, properties.ToByteArray());
            PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = properties });
        }
        var remove = PresentationTextBodyProperties.Parser.ParseFrom(new PresentationTextBodyProperties { NoAnchorCenter = true }.ToByteArray());
        Assert.True(remove.HasNoAnchorCenter); Assert.True(remove.NoAnchorCenter); Assert.False(remove.HasAnchorCenter);
        PptxBodyPropertiesCodec.Validate(new PresentationTextBody { BodyProperties = remove });
        Assert.True(PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(remove));
        foreach (var invalid in new[]
        {
            new PresentationTextBodyProperties { NoAnchorCenter = false },
            new PresentationTextBodyProperties { NoAnchorCenter = true, AnchorCenter = false },
            new PresentationTextBodyProperties { NoAnchorCenter = true, AnchorCenter = true },
        })
        {
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
    [InlineData("text", "columnGap")]
    [InlineData("text", "columns")]
    [InlineData("master", "columnGap")]
    [InlineData("master", "columns")]
    [InlineData("layout", "columnGap")]
    [InlineData("layout", "columns")]
    [InlineData("table", "columnGap")]
    [InlineData("table", "columns")]
    public void NumericColumnPropertySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind); var style = Style(program, kind);
        style["columnGap"] = 12.5; style["columns"] = 3; style["columnDirection"] = "right-to-left";
        var source = PptxCodecTests.RemoveEmbeddedPpj(Compile(program).File.ToByteArray());
        var original = source.ToArray();
        var request = Project(source);
        Assert.Equal(field == "columns" ? 3 : 12.5, Style(request, kind)[field]!.GetValue<double>());
        Assert.Equal(source, Compile(request, source).File.ToByteArray());
        Assert.Equal(field == "columns" ? 3 : 158750, NumericColumnProperty(source, kind, field));
        Style(request, kind).Remove(field);
        var deleted = Compile(request, source).File.ToByteArray();
        Assert.Null(NumericColumnProperty(deleted, kind, field));
        Assert.False(Style(Project(deleted), kind).ContainsKey(field));
        AssertOnlyBodyPropertyChanged(source, deleted, kind, field);
        foreach (var value in field == "columns" ? new[] { 1d, 3d, 16d } : new[] { 0d, 12.5d, 10000d })
        {
            var restore = Project(deleted);
            Style(restore, kind)[field] = value;
            var restored = Compile(restore, deleted).File.ToByteArray();
            Assert.Equal(checked((int)(value * (field == "columns" ? 1 : 12700))), NumericColumnProperty(restored, kind, field));
            Assert.Equal(value, Style(Project(restored), kind)[field]!.GetValue<double>());
            AssertOnlyBodyPropertyChanged(deleted, restored, kind, field);
            var removeAgain = Project(restored);
            Style(removeAgain, kind).Remove(field);
            Assert.Null(NumericColumnProperty(Compile(removeAgain, restored).File.ToByteArray(), kind, field));
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("shape", "columnGap")]
    [InlineData("shape", "columns")]
    [InlineData("table", "columnGap")]
    [InlineData("table", "columns")]
    public void NumericColumnPropertyOnlyAndCombinedRemovableStylesPreservePresence(string kind, string field)
    {
        foreach (var withOtherProperties in new[] { false, true })
        {
            var program = Program(kind); var style = Style(program, kind);
            style.Clear(); style[field] = field == "columns" ? 3 : 12.5;
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
            Assert.Null(NumericColumnProperty(deleted, kind, field));
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
            Assert.Equal(field == "columns" ? 1 : 0, NumericColumnProperty(restored, kind, field));
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

    private static int? NumericColumnProperty(byte[] bytes, string kind, string field) =>
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
    [InlineData("table", "upright")]
    [InlineData("table", "anchorCenter")]
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

        var otherProgram = Program(kind);
        Style(otherProgram, kind)["forceAntiAlias"] = true;
        var otherSource = PptxCodecTests.RemoveEmbeddedPpj(Compile(otherProgram).File.ToByteArray());
        var otherEdit = Project(otherSource);
        StyleOwner(otherEdit).Remove(kind == "shape" ? "textStyle" : "style");
        Assert.Empty(Compile(otherEdit, otherSource, success: false).File);
    }

    [Theory]
    [InlineData("shape", "upright")]
    [InlineData("shape", "anchorCenter")]
    [InlineData("text", "upright")]
    [InlineData("text", "anchorCenter")]
    [InlineData("master", "upright")]
    [InlineData("master", "anchorCenter")]
    [InlineData("layout", "upright")]
    [InlineData("layout", "anchorCenter")]
    [InlineData("table", "upright")]
    [InlineData("table", "anchorCenter")]
    public void BooleanBodySourceRemovalAndRestorationPreserveOtherState(string kind, string field)
    {
        var program = Program(kind);
        Style(program, kind)[field] = true;
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
        field == "anchorCenter" ? Body(bytes, kind).AnchorCenter?.Value : Body(bytes, kind).UpRight?.Value;

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
        if (field == "autoFit")
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
