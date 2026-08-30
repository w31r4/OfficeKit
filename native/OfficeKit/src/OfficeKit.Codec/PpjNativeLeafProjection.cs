using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjNativeLeafBinding(
    string Id,
    string PageId,
    string ElementId,
    string Kind,
    uint NativeLeafIndex,
    uint TextLeafIndex,
    string ExpectedValue);

/// <summary>
/// Issues the finite scalar leaves that the native token-splice writer can
/// independently prove again. Package paths, relationship IDs and raw XML stay
/// outside PPJ; the binding is internal to the fresh source projection.
/// </summary>
internal static class PpjNativeLeafProjection
{
    private const int MaxIssuedLeavesPerElement = 256;

    private static readonly HashSet<string> StringKinds = new(StringComparer.Ordinal)
    {
        "text", "tableCellText", "nativeText", "paragraphAlignment",
        "paragraphBulletCharacter", "paragraphBulletAutoNumberScheme",
        "paragraphBulletFontFamily", "paragraphBulletColorScheme", "verticalAnchor",
        "textBodyWrap", "textBodyAutoFit", "textBodyVerticalText", "fontFamily",
        "fontFamilyEastAsia", "fontLanguage", "fontUnderline", "fontStrike", "fontColorScheme",
        "fontCaps", "fontHighlightScheme", "fillScheme", "lineScheme", "lineStyle", "lineCap", "lineJoin",
        "lineStartArrow", "lineEndArrow", "imageMaskPreset",
    };

    private static readonly HashSet<string> RgbKinds = new(StringComparer.Ordinal)
    {
        "paragraphBulletColorRgb", "fontColorRgb", "fontHighlightRgb", "fillRgb", "lineRgb",
    };

    private static readonly HashSet<string> BooleanKinds = new(StringComparer.Ordinal)
    {
        "textBodyColumnDirection", "fontBold", "fontItalic", "flipHorizontal", "flipVertical",
    };

    private static readonly HashSet<string> IntegerKinds = new(StringComparer.Ordinal)
    {
        "paragraphMarginLeftEmu", "paragraphIndentEmu", "paragraphBulletAutoNumberStartAt",
        "paragraphLevel", "textBodyInsetLeftEmu", "textBodyInsetTopEmu",
        "textBodyInsetRightEmu", "textBodyInsetBottomEmu", "textBodyColumnCount",
        "fillOpacityThousandthPercent", "imageOpacityThousandthPercent", "lineWidthEmu", "leftEmu", "topEmu",
        "widthEmu", "heightEmu",
    };

    private static readonly IReadOnlyDictionary<string, int> ScaledNumberKinds =
        new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["paragraphLineSpacingPoints"] = 100,
            ["paragraphLineSpacingMultiplier"] = 100_000,
            ["paragraphSpaceBeforePoints"] = 100,
            ["paragraphSpaceBeforeMultiplier"] = 100_000,
            ["paragraphSpaceAfterPoints"] = 100,
            ["paragraphSpaceAfterMultiplier"] = 100_000,
            ["paragraphBulletSizePoints"] = 100,
            ["paragraphBulletSizePercent"] = 1_000,
            ["fontSizePoints"] = 100,
            ["fontKerningPoints"] = 100,
            ["fontBaselinePercent"] = 1_000,
            ["fontSpacingPoints"] = 100,
            ["rotationDegrees"] = 60_000,
        };

    internal static JsonArray Describe(
        string sourceSha256,
        string pageId,
        string elementId,
        PresentationElement element,
        IReadOnlyList<uint> shapeTreePath,
        Action<PpjNativeLeafBinding> record)
    {
        var output = new JsonArray();
        var source = element.Source;
        if (source is null || shapeTreePath.Count == 0) return output;

        void Add(string kind, string expectedValue, JsonNode? value, uint nativeIndex = 0, uint textIndex = 0)
        {
            // nativeRef.leaves is a bounded capability list, not a second copy
            // of the source tree. Keep projection deterministic for unusually
            // large text boxes and tables; the source object remains preserved
            // even when additional scalar edit capabilities are not issued.
            if (output.Count >= MaxIssuedLeavesPerElement) return;
            var expectedHash = Hash(expectedValue);
            var seed = string.Join("\0", sourceSha256, pageId, elementId, string.Join('/', shapeTreePath),
                source.ElementSha256, kind, nativeIndex, textIndex, expectedHash);
            var id = $"nl_{Hash(seed)[..32]}";
            output.Add(new JsonObject
            {
                ["id"] = id,
                ["kind"] = kind,
                ["expectedHash"] = expectedHash,
                ["value"] = value,
            });
            record(new(id, pageId, elementId, kind, nativeIndex, textIndex, expectedValue));
        }

        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                DescribeShape(element.Shape, source, Add);
                break;
            case PresentationElement.ContentOneofCase.Image:
                DescribeImage(element.Image, source, Add);
                break;
            case PresentationElement.ContentOneofCase.Connector:
                DescribeConnector(element.Connector, source, Add);
                break;
            case PresentationElement.ContentOneofCase.Table:
                DescribeTable(element.Table, source, Add);
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (PpjNativeTextProjection.TryRead(element.Opaque.RawXml, out var leaves))
                    for (var index = 0; index < leaves.Count; index++)
                        Add("nativeText", leaves[index], JsonValue.Create(leaves[index]), textIndex: checked((uint)index));
                break;
        }
        return output;
    }

    /// <summary>
    /// Converts the human-facing PPJ scalar into the exact native token that
    /// the source-bound edit-plan codec will independently re-prove. This is a
    /// closed vocabulary: adding a runtime leaf without adding its PPJ value
    /// semantics is a build-time/documentation error, not an implicit string
    /// conversion.
    /// </summary>
    internal static string NormalizeValue(string kind, JsonElement value, string path)
    {
        if (RgbKinds.Contains(kind))
        {
            var token = RequireString(value, kind, path).TrimStart('#');
            if (token.Length != 6 || token.Any(character => !Uri.IsHexDigit(character)))
                throw InvalidValue(kind, path, "an RGB color such as #1A2B3C");
            return token.ToUpperInvariant();
        }
        if (BooleanKinds.Contains(kind))
        {
            if (value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                throw InvalidValue(kind, path, "a boolean");
            return value.GetBoolean() ? "1" : "0";
        }
        if (IntegerKinds.Contains(kind))
        {
            if (!value.TryGetInt64(out var integer)) throw InvalidValue(kind, path, "an integer");
            return integer.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (ScaledNumberKinds.TryGetValue(kind, out var scale))
        {
            if (value.ValueKind != JsonValueKind.Number || !value.TryGetDouble(out var number) || !double.IsFinite(number))
                throw InvalidValue(kind, path, "a finite number");
            var raw = checked((long)Math.Round(number * scale, MidpointRounding.AwayFromZero));
            return raw.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (StringKinds.Contains(kind))
        {
            var token = RequireString(value, kind, path);
            if (kind == "fontUnderline")
                return token switch { "single" => "sng", "double" => "dbl", _ => token };
            if (kind == "fontStrike")
                return token switch { "single" => "sngStrike", "double" => "dblStrike", "none" => "noStrike", _ => token };
            return token;
        }
        throw new CodecException(
            "ppj.nativeRef.leafKind",
            $"Native leaf kind {kind} is not exposed as a directly editable PPJ scalar.",
            path);
    }

    private static string RequireString(JsonElement value, string kind, string path)
    {
        if (value.ValueKind != JsonValueKind.String)
            throw InvalidValue(kind, path, "a string");
        return value.GetString() ?? string.Empty;
    }

    private static CodecException InvalidValue(string kind, string path, string expected) => new(
        "ppj.nativeRef.leafValue",
        $"Native leaf {kind} requires {expected}.",
        path);

    private static void DescribeShape(
        PresentationShape shape,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (source.Editable)
        {
            AddInteger(add, "leftEmu", shape.LeftEmu);
            AddInteger(add, "topEmu", shape.TopEmu);
            AddInteger(add, "widthEmu", shape.WidthEmu);
            AddInteger(add, "heightEmu", shape.HeightEmu);
            if (shape.Transform?.HasRotationAngle60000 == true)
                add("rotationDegrees", shape.Transform.RotationAngle60000.ToStringInvariant(), JsonValue.Create(shape.Transform.RotationAngle60000 / 60_000d), 0, 0);
            if (shape.Transform?.HasFlipHorizontal == true)
                AddBoolean(add, "flipHorizontal", shape.Transform.FlipHorizontal);
            if (shape.Transform?.HasFlipVertical == true)
                AddBoolean(add, "flipVertical", shape.Transform.FlipVertical);
        }

        if (!string.IsNullOrEmpty(shape.FillRgb))
            add("fillRgb", shape.FillRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.FillRgb.ToLowerInvariant()}"), 0, 0);
        if (shape.HasFillOpacityThousandthPercent)
            add("fillOpacityThousandthPercent", shape.FillOpacityThousandthPercent.ToStringInvariant(), JsonValue.Create(shape.FillOpacityThousandthPercent), 0, 0);
        if (!string.IsNullOrEmpty(shape.FillScheme))
            add("fillScheme", shape.FillScheme, JsonValue.Create(shape.FillScheme), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineRgb))
            add("lineRgb", shape.LineRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.LineRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineScheme))
            add("lineScheme", shape.LineScheme, JsonValue.Create(shape.LineScheme), 0, 0);
        if (shape.LineWidthEmu > 0)
            AddInteger(add, "lineWidthEmu", shape.LineWidthEmu);

        // The importer only issues TextEditable after checking the bounded
        // source shape profile. The edit-plan codec repeats that proof against
        // the exact SlidePart during compilation.
        if (!source.TextEditable || shape.TextBody is null) return;
        DescribeBody(shape.TextBody, add);
    }

    private static void DescribeImage(
        PresentationImage image,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (!source.Editable) return;
        AddInteger(add, "leftEmu", image.LeftEmu);
        AddInteger(add, "topEmu", image.TopEmu);
        AddInteger(add, "widthEmu", image.WidthEmu);
        AddInteger(add, "heightEmu", image.HeightEmu);
        if (image.Transform?.HasRotationAngle60000 == true)
            add("rotationDegrees", image.Transform.RotationAngle60000.ToStringInvariant(), JsonValue.Create(image.Transform.RotationAngle60000 / 60_000d), 0, 0);
        if (image.Transform?.HasFlipHorizontal == true)
            AddBoolean(add, "flipHorizontal", image.Transform.FlipHorizontal);
        if (image.Transform?.HasFlipVertical == true)
            AddBoolean(add, "flipVertical", image.Transform.FlipVertical);
        if (image.HasOpacityThousandthPercent)
            AddInteger(add, "imageOpacityThousandthPercent", image.OpacityThousandthPercent);
        if (!string.IsNullOrEmpty(image.MaskPreset) && image.MaskPresetAdjustments.Count == 0)
            add("imageMaskPreset", image.MaskPreset, JsonValue.Create(image.MaskPreset), 0, 0);
    }

    private static void DescribeConnector(
        PresentationConnector connector,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (source is null) return;
        if (!string.IsNullOrEmpty(connector.LineRgb))
            add("lineRgb", connector.LineRgb.ToUpperInvariant(), JsonValue.Create($"#{connector.LineRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(connector.LineScheme))
            add("lineScheme", connector.LineScheme, JsonValue.Create(connector.LineScheme), 0, 0);
        if (connector.LineWidthEmu > 0) AddInteger(add, "lineWidthEmu", connector.LineWidthEmu);
        if (!string.IsNullOrEmpty(connector.LineStyle)) add("lineStyle", connector.LineStyle, JsonValue.Create(connector.LineStyle), 0, 0);
        if (!string.IsNullOrEmpty(connector.LineCap)) add("lineCap", connector.LineCap, JsonValue.Create(connector.LineCap), 0, 0);
        if (!string.IsNullOrEmpty(connector.LineJoin)) add("lineJoin", connector.LineJoin, JsonValue.Create(connector.LineJoin), 0, 0);
        if (!string.IsNullOrEmpty(connector.StartArrow)) add("lineStartArrow", connector.StartArrow, JsonValue.Create(connector.StartArrow), 0, 0);
        if (!string.IsNullOrEmpty(connector.EndArrow)) add("lineEndArrow", connector.EndArrow, JsonValue.Create(connector.EndArrow), 0, 0);
    }

    private static void DescribeTable(
        PresentationTable table,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (!source.Editable) return;
        uint cellIndex = 0;
        foreach (var cell in table.Rows.SelectMany(row => row.Cells))
        {
            add("tableCellText", cell.Text, JsonValue.Create(cell.Text), 0, cellIndex);
            cellIndex++;
        }
    }

    private static void DescribeBody(
        PresentationTextBody body,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        var properties = body.BodyProperties;
        if (properties is not null)
        {
            if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor)
                add("verticalAnchor", properties.VerticalAnchor, JsonValue.Create(properties.VerticalAnchor), 0, 0);
            AddBodyInset(properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, "textBodyInsetLeftEmu", properties.LeftInsetEmu, add);
            AddBodyInset(properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, "textBodyInsetTopEmu", properties.TopInsetEmu, add);
            AddBodyInset(properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, "textBodyInsetRightEmu", properties.RightInsetEmu, add);
            AddBodyInset(properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, "textBodyInsetBottomEmu", properties.BottomInsetEmu, add);
            if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
                add("textBodyWrap", properties.Wrap, JsonValue.Create(properties.Wrap), 0, 0);
            if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
                add("textBodyColumnCount", properties.Columns.ToStringInvariant(), JsonValue.Create(properties.Columns), 0, 0);
            if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode && properties.NormalAutoFit is null)
                add("textBodyAutoFit", properties.AutoFitMode, JsonValue.Create(properties.AutoFitMode), 0, 0);
            if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns)
                add("textBodyColumnDirection", properties.RightToLeftColumns ? "1" : "0", JsonValue.Create(properties.RightToLeftColumns), 0, 0);
            if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode)
                add("textBodyVerticalText", properties.VerticalTextMode, JsonValue.Create(properties.VerticalTextMode), 0, 0);
        }

        uint textLeafIndex = 0;
        uint runStyleIndex = 0;
        for (var paragraphIndex = 0; paragraphIndex < body.Paragraphs.Count; paragraphIndex++)
        {
            var paragraph = body.Paragraphs[paragraphIndex];
            var nativeIndex = checked((uint)paragraphIndex);
            if (paragraph.HasAlignment) add("paragraphAlignment", paragraph.Alignment, JsonValue.Create(paragraph.Alignment), nativeIndex, 0);
            if (paragraph.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                AddScaled(add, "paragraphLineSpacingPoints", paragraph.LineSpacingPoints, 100, nativeIndex);
            if (paragraph.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                AddScaled(add, "paragraphLineSpacingMultiplier", paragraph.LineSpacingMultiplier, 100_000, nativeIndex);
            if (paragraph.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                AddScaled(add, "paragraphSpaceBeforePoints", paragraph.SpaceBeforePoints, 100, nativeIndex);
            if (paragraph.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                AddScaled(add, "paragraphSpaceBeforeMultiplier", paragraph.SpaceBeforeMultiplier, 100_000, nativeIndex);
            if (paragraph.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                AddScaled(add, "paragraphSpaceAfterPoints", paragraph.SpaceAfterPoints, 100, nativeIndex);
            if (paragraph.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                AddScaled(add, "paragraphSpaceAfterMultiplier", paragraph.SpaceAfterMultiplier, 100_000, nativeIndex);
            if (paragraph.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                AddInteger(add, "paragraphMarginLeftEmu", paragraph.MarginLeftEmu, nativeIndex);
            if (paragraph.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                AddInteger(add, "paragraphIndentEmu", paragraph.IndentEmu, nativeIndex);
            if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.BulletCharacter)
                add("paragraphBulletCharacter", paragraph.BulletCharacter, JsonValue.Create(paragraph.BulletCharacter), nativeIndex, 0);
            if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber)
            {
                add("paragraphBulletAutoNumberScheme", paragraph.AutoNumber.Scheme, JsonValue.Create(paragraph.AutoNumber.Scheme), nativeIndex, 0);
                if (paragraph.AutoNumber.HasStartAt)
                    add("paragraphBulletAutoNumberStartAt", paragraph.AutoNumber.StartAt.ToStringInvariant(), JsonValue.Create(paragraph.AutoNumber.StartAt), nativeIndex, 0);
            }
            if (paragraph.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily)
                add("paragraphBulletFontFamily", paragraph.BulletFontFamily, JsonValue.Create(paragraph.BulletFontFamily), nativeIndex, 0);
            if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb)
                add("paragraphBulletColorRgb", paragraph.BulletColorRgb, JsonValue.Create($"#{paragraph.BulletColorRgb.ToLowerInvariant()}"), nativeIndex, 0);
            if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme)
                add("paragraphBulletColorScheme", paragraph.BulletColorScheme, JsonValue.Create(paragraph.BulletColorScheme), nativeIndex, 0);
            if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints)
                AddScaled(add, "paragraphBulletSizePoints", paragraph.BulletSizePoints, 100, nativeIndex);
            if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent)
                AddScaled(add, "paragraphBulletSizePercent", paragraph.BulletSizePercent, 1000, nativeIndex);
            if (paragraph.HasLevel)
                add("paragraphLevel", paragraph.Level.ToStringInvariant(), JsonValue.Create(paragraph.Level), nativeIndex, 0);

            foreach (var run in paragraph.Runs)
            {
                if (run.ContentCase == PresentationTextRun.ContentOneofCase.Text)
                {
                    add("text", run.Text, JsonValue.Create(run.Text), 0, textLeafIndex);
                    textLeafIndex++;
                }
                DescribeRun(run, runStyleIndex, add);
                runStyleIndex++;
            }
        }
    }

    private static void DescribeRun(
        PresentationTextRun run,
        uint index,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (run.HasFontSizePoints) AddScaled(add, "fontSizePoints", run.FontSizePoints, 100, index);
        if (run.HasFontFamily) add("fontFamily", run.FontFamily, JsonValue.Create(run.FontFamily), index, 0);
        if (run.HasFontFamilyEastAsia) add("fontFamilyEastAsia", run.FontFamilyEastAsia, JsonValue.Create(run.FontFamilyEastAsia), index, 0);
        if (run.HasLanguage) add("fontLanguage", run.Language, JsonValue.Create(run.Language), index, 0);
        if (run.HasBold) AddBoolean(add, "fontBold", run.Bold, index);
        if (run.HasItalic) AddBoolean(add, "fontItalic", run.Italic, index);
        if (run.HasUnderline) add("fontUnderline", run.Underline, JsonValue.Create(run.Underline), index, 0);
        if (run.HasStrike) add("fontStrike", run.Strike, JsonValue.Create(run.Strike), index, 0);
        if (run.HasColorRgb) add("fontColorRgb", run.ColorRgb, JsonValue.Create($"#{run.ColorRgb.ToLowerInvariant()}"), index, 0);
        else if (run.HasColorScheme) add("fontColorScheme", run.ColorScheme, JsonValue.Create(run.ColorScheme), index, 0);
        if (run.HasFontKerningPoints) AddScaled(add, "fontKerningPoints", run.FontKerningPoints, 100, index);
        if (run.HasFontBaselinePercent) AddScaled(add, "fontBaselinePercent", run.FontBaselinePercent, 1000, index);
        if (run.HasFontSpacingPoints) AddScaled(add, "fontSpacingPoints", run.FontSpacingPoints, 100, index);
        if (run.HasFontCaps) add("fontCaps", run.FontCaps, JsonValue.Create(run.FontCaps), index, 0);
        if (run.HighlightCase == PresentationTextRun.HighlightOneofCase.HighlightRgb)
            add("fontHighlightRgb", run.HighlightRgb.ToUpperInvariant(), JsonValue.Create($"#{run.HighlightRgb.ToLowerInvariant()}"), index, 0);
        else if (run.HighlightCase == PresentationTextRun.HighlightOneofCase.HighlightScheme)
            add("fontHighlightScheme", run.HighlightScheme, JsonValue.Create(run.HighlightScheme), index, 0);
    }

    private static void AddBodyInset(bool present, string kind, long value, Action<string, string, JsonNode?, uint, uint> add)
    {
        if (present) AddInteger(add, kind, value);
    }

    private static void AddInteger(Action<string, string, JsonNode?, uint, uint> add, string kind, long value, uint nativeIndex = 0) =>
        add(kind, value.ToStringInvariant(), JsonValue.Create(value), nativeIndex, 0);

    private static void AddScaled(Action<string, string, JsonNode?, uint, uint> add, string kind, double value, int scale, uint nativeIndex)
    {
        var raw = checked((long)Math.Round(value * scale, MidpointRounding.AwayFromZero));
        add(kind, raw.ToStringInvariant(), JsonValue.Create(value), nativeIndex, 0);
    }

    private static void AddBoolean(Action<string, string, JsonNode?, uint, uint> add, string kind, bool value, uint nativeIndex = 0) =>
        add(kind, value ? "1" : "0", JsonValue.Create(value), nativeIndex, 0);

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string ToStringInvariant(this long value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this uint value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
