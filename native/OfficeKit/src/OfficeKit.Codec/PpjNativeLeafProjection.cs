using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
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
    string ExpectedValue,
    PpjNativeChartDataBinding? ChartData = null);

internal sealed record PpjNativeChartDataBinding(
    string TargetPartPath,
    string TargetPartSha256,
    string RelationshipId,
    string EmbeddedPackagePartPath,
    string EmbeddedPackageSha256,
    string EmbeddedPackageRelationshipId,
    string EmbeddedWorksheetPartPath,
    string EmbeddedWorksheetSha256,
    string EmbeddedCellReference,
    uint ChartSeriesIndex,
    uint ChartPointIndex,
    string ChartFormula,
    string ChartChannel);

/// <summary>
/// Issues the finite scalar leaves that the native token-splice writer can
/// independently prove again. Package paths, relationship IDs and raw XML stay
/// outside PPJ; the binding is internal to the fresh source projection.
/// </summary>
internal static class PpjNativeLeafProjection
{
    private const int MaxIssuedLeavesPerElement = 256;
    private const int MaxChartCategoryLength = 32_767;

    private static readonly HashSet<string> StringKinds = new(StringComparer.Ordinal)
    {
        "text", "tableCellText", "nativeText", "paragraphAlignment",
        "paragraphBulletCharacter", "paragraphBulletAutoNumberScheme",
        "paragraphBulletFontFamily", "paragraphBulletColorScheme", "verticalAnchor",
        "textBodyWrap", "textBodyAutoFit", "textBodyVerticalText", "fontFamily",
        "fontFamilyEastAsia", "fontLanguage", "fontUnderline", "fontStrike", "fontColorScheme",
        "fontCaps", "fontHighlightScheme", "fillScheme", "shadowAlignment", "shadowColorScheme", "lineScheme", "lineStyle", "lineCap", "lineJoin",
        "lineStartArrow", "lineEndArrow", "imageMaskPreset", "chartDataCategory",
    };

    private static readonly HashSet<string> RgbKinds = new(StringComparer.Ordinal)
    {
        "paragraphBulletColorRgb", "fontColorRgb", "fontHighlightRgb", "fillRgb", "shadowColorRgb", "lineRgb",
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
        "fillOpacityThousandthPercent", "shadowOpacityThousandthPercent", "shadowBlurRadiusEmu", "shadowDistanceEmu",
        "imageOpacityThousandthPercent", "lineWidthEmu", "leftEmu", "topEmu",
        "widthEmu", "heightEmu", "childLeftEmu", "childTopEmu",
        "childWidthEmu", "childHeightEmu",
        "customGeometryAdjustment",
        "presetGeometryAdjustment",
        "imageMaskAdjustment",
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
            ["textBodyNormalAutoFitFontScale"] = 1_000,
            ["textBodyNormalAutoFitLineSpacingReduction"] = 1_000,
            ["rotationDegrees"] = 60_000,
            ["shadowDirectionDegrees"] = 60_000,
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

        void Add(
            string kind,
            string expectedValue,
            JsonNode? value,
            uint nativeIndex = 0,
            uint textIndex = 0,
            PpjNativeChartDataBinding? chartData = null)
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
                ["id"] = StringNode(id),
                ["kind"] = StringNode(kind),
                ["expectedHash"] = StringNode(expectedHash),
                ["value"] = value,
            });
            record(new(id, pageId, elementId, kind, nativeIndex, textIndex, expectedValue, chartData));
        }

        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                DescribeShape(element.Shape, source, (kind, expected, value, nativeIndex, textIndex) =>
                    Add(kind, expected, value, nativeIndex, textIndex));
                break;
            case PresentationElement.ContentOneofCase.Image:
                DescribeImage(element.Image, source, (kind, expected, value, nativeIndex, textIndex) =>
                    Add(kind, expected, value, nativeIndex, textIndex));
                break;
            case PresentationElement.ContentOneofCase.Connector:
                DescribeConnector(element.Connector, source, (kind, expected, value, nativeIndex, textIndex) =>
                    Add(kind, expected, value, nativeIndex, textIndex));
                break;
            case PresentationElement.ContentOneofCase.Table:
                DescribeTable(element.Table, source, (kind, expected, value, nativeIndex, textIndex) =>
                    Add(kind, expected, value, nativeIndex, textIndex));
                break;
            case PresentationElement.ContentOneofCase.Group:
                DescribeGroup(element.Group, source, (kind, expected, value, nativeIndex, textIndex) =>
                    Add(kind, expected, value, nativeIndex, textIndex));
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (PpjNativeTextProjection.TryRead(element.Opaque.RawXml, out var leaves))
                    for (var index = 0; index < leaves.Count; index++)
                        Add("nativeText", leaves[index], JsonValue.Create(leaves[index]), textIndex: checked((uint)index));
                if (string.Equals(element.Opaque.NativeKind, "picture", StringComparison.Ordinal))
                    DescribeOpaquePictureMask(element.Opaque.RawXml, (kind, expected, value, nativeIndex, textIndex) =>
                        Add(kind, expected, value, nativeIndex, textIndex));
                if (element.Opaque.NativeChart is { DataPoints.Count: > 0 } chart)
                {
                    foreach (var point in chart.DataPoints)
                    {
                        if (ChartChannel(point) == "category" ? !ValidTextToken(point.Value) : !ValidNumericToken(point.Value)) continue;
                        Add(
                            ChartDataLeafKind(point),
                            point.Value,
                            JsonValue.Create(point.Value),
                            nativeIndex: point.SeriesIndex,
                            textIndex: point.PointIndex,
                            chartData: new PpjNativeChartDataBinding(
                                chart.PartPath,
                                chart.SourceSha256,
                                chart.RelationshipId,
                                chart.EmbeddedPackagePartPath,
                                chart.EmbeddedPackageSourceSha256,
                                chart.EmbeddedPackageRelationshipId,
                                point.WorksheetPartPath,
                                point.WorksheetSourceSha256,
                                point.CellReference,
                                point.SeriesIndex,
                                point.PointIndex,
                                point.Formula,
                                ChartChannel(point)));
                    }
                }
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
        if (kind == "chartDataCategory")
        {
            var token = RequireString(value, kind, path);
            if (!ValidTextToken(token))
                throw InvalidValue(kind, path, "a bounded chart category label");
            return token;
        }
        if (IsChartDataLeafKind(kind))
        {
            var token = value.ValueKind switch
            {
                JsonValueKind.String => value.GetString() ?? string.Empty,
                JsonValueKind.Number => value.GetRawText(),
                _ => string.Empty,
            };
            if (!ValidNumericToken(token))
                throw InvalidValue(kind, path, "a finite numeric token");
            return token;
        }
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
        if (kind is "customGeometryAdjustment" or "presetGeometryAdjustment" or "imageMaskAdjustment")
        {
            if (!value.TryGetInt64(out var adjustment) || adjustment is < int.MinValue or > int.MaxValue ||
                kind is "presetGeometryAdjustment" or "imageMaskAdjustment" &&
                (adjustment < PptxPresetGeometryAdjustmentCodec.MinimumValue || adjustment > PptxPresetGeometryAdjustmentCodec.MaximumValue))
                throw InvalidValue(kind, path, "a bounded DrawingML signed integer");
            return adjustment.ToString(System.Globalization.CultureInfo.InvariantCulture);
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

    // JsonNode's implicit primitive conversions use reflection metadata that
    // is intentionally unavailable in the NativeAOT PPJ host.  Build string
    // leaves through Utf8JsonWriter so imported source-bound projection stays
    // AOT-safe while preserving the exact JSON primitive value.
    private static JsonNode StringNode(string value)
    {
        var buffer = new System.Buffers.ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer)) writer.WriteStringValue(value);
        return JsonNode.Parse(buffer.WrittenSpan) ?? throw new InvalidOperationException("String JSON primitive could not be created.");
    }

    private static bool ValidNumericToken(string token) =>
        token.Length is > 0 and <= 128 &&
        token == token.Trim() &&
        !token.Any(char.IsControl) &&
        double.TryParse(token, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number) &&
        double.IsFinite(number);

    internal static bool ValidTextToken(string token) =>
        token.Length <= MaxChartCategoryLength &&
        !token.Any(character => character is >= '\u0000' and <= '\u0008' or '\u000B' or '\u000C' or >= '\u000E' and <= '\u001F');

    internal static bool IsChartDataLeafKind(string kind) => kind is
        "chartDataCategory" or "chartDataValue" or "chartDataXValue" or "chartDataYValue" or "chartDataBubbleSize";

    private static string ChartChannel(PresentationNativeChartDataPoint point) =>
        string.IsNullOrEmpty(point.Channel) ? "value" : point.Channel;

    private static string ChartDataLeafKind(PresentationNativeChartDataPoint point) =>
        ChartChannel(point) switch
        {
            "x" => "chartDataXValue",
            "y" => "chartDataYValue",
            "size" => "chartDataBubbleSize",
            "category" => "chartDataCategory",
            _ => "chartDataValue",
        };

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
        if (shape.Shadow is { } shadow)
        {
            if (shadow.HasBlurRadiusEmu)
                AddInteger(add, "shadowBlurRadiusEmu", shadow.BlurRadiusEmu);
            if (shadow.HasDistanceEmu)
                AddInteger(add, "shadowDistanceEmu", shadow.DistanceEmu);
            if (shadow.HasDirectionAngle60000)
                add("shadowDirectionDegrees", shadow.DirectionAngle60000.ToStringInvariant(),
                    JsonValue.Create(shadow.DirectionAngle60000 / 60_000d), 0, 0);
            if (shadow.HasAlignment)
                add("shadowAlignment", shadow.Alignment, JsonValue.Create(shadow.Alignment), 0, 0);
            if (!string.IsNullOrEmpty(shadow.ColorRgb))
                add("shadowColorRgb", shadow.ColorRgb.ToUpperInvariant(),
                    JsonValue.Create($"#{shadow.ColorRgb.ToLowerInvariant()}"), 0, 0);
            else if (shadow.HasColorScheme && !string.IsNullOrEmpty(shadow.ColorScheme))
                add("shadowColorScheme", shadow.ColorScheme, JsonValue.Create(shadow.ColorScheme), 0, 0);
            if (shadow.HasOpacityThousandthPercent)
                AddInteger(add, "shadowOpacityThousandthPercent", shadow.OpacityThousandthPercent);
        }
        if (!string.IsNullOrEmpty(shape.LineRgb))
            add("lineRgb", shape.LineRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.LineRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineScheme))
            add("lineScheme", shape.LineScheme, JsonValue.Create(shape.LineScheme), 0, 0);
        if (shape.LineWidthEmu > 0)
            AddInteger(add, "lineWidthEmu", shape.LineWidthEmu);
        if (shape.HasLineStyleExplicit && shape.LineStyleExplicit && !string.IsNullOrEmpty(shape.LineStyle))
            add("lineStyle", shape.LineStyle, JsonValue.Create(shape.LineStyle), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineCap))
            add("lineCap", shape.LineCap, JsonValue.Create(shape.LineCap), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineJoin))
            add("lineJoin", shape.LineJoin, JsonValue.Create(shape.LineJoin), 0, 0);
        if (!string.IsNullOrEmpty(shape.StartArrow))
            add("lineStartArrow", shape.StartArrow, JsonValue.Create(shape.StartArrow), 0, 0);
        if (!string.IsNullOrEmpty(shape.EndArrow))
            add("lineEndArrow", shape.EndArrow, JsonValue.Create(shape.EndArrow), 0, 0);

        // A custom geometry may be fully recognized even when its path
        // coordinates are driven by DrawingML guide formulas.  Keep the
        // formula graph source-owned, but expose the safest adjustment leaf:
        // an existing ordered `a:avLst/a:gd fmla="val N"` can be changed by
        // token splice without changing guide names, path references, or
        // handle/connection topology.  Formula-valued adjustments and all
        // calculated guides remain undisclosed here and therefore fail closed.
        if (source.Editable && shape.Geometry == "custom")
        {
            for (var index = 0; index < shape.CustomAdjustments.Count; index++)
            {
                var formula = shape.CustomAdjustments[index].Formula;
                if (TryLiteralAdjustment(formula, out var value))
                    add("customGeometryAdjustment", value.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(value), checked((uint)index), 0);
            }
        }
        // Preset geometry adjustments use the fixed guide names and ordered
        // slots from the native preset profile.  A complete profile whose
        // target formula is exactly `val N` is safe to edit by token splice;
        // the preset token, guide order, and all other shape topology remain
        // source-owned.  Partial or calculated guide graphs are not issued.
        if (shape.Geometry != "custom" &&
            shape.PresetAdjustments.Count > 0 &&
            PptxPresetGeometryAdjustmentCodec.TryExpectedCount(shape.Geometry, out var expectedCount) &&
            shape.PresetAdjustments.Count == expectedCount &&
            (source.Editable || shape.PresetAdjustments.Any(PptxPresetGeometryAdjustmentCodec.IsMissingValue)))
        {
            for (var index = 0; index < shape.PresetAdjustments.Count; index++)
            {
                var value = shape.PresetAdjustments[index];
                if (!PptxPresetGeometryAdjustmentCodec.IsMissingValue(value) &&
                    value >= PptxPresetGeometryAdjustmentCodec.MinimumValue && value <= PptxPresetGeometryAdjustmentCodec.MaximumValue)
                    add("presetGeometryAdjustment", value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        JsonValue.Create(value), checked((uint)index), 0);
            }
        }

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
        // A missing PPJ mask is the native rectangle default.  Issue the
        // canonical `rect` leaf as well, so a source-bound continuation can
        // change an otherwise unmasked picture to a supported preset without
        // manufacturing a new relationship or flattening the image.
        if (image.CustomMaskPaths.Count == 0 && image.MaskPresetAdjustments.Count == 0)
        {
            var preset = string.IsNullOrEmpty(image.MaskPreset) ? "rect" : image.MaskPreset;
            add("imageMaskPreset", preset, JsonValue.Create(preset), 0, 0);
        }
        // A complete preset mask already has a typed PPJ representation, but
        // each literal adjustment is also an independently safe source leaf:
        // token-splice the `val N` formula without rebuilding the preset,
        // guide names, or picture relationships.  A partial/formula-valued
        // mask remains opaque as a semantic image; the opaque-picture path
        // below may still expose any independently proven literal sibling.
        if (image.MaskPresetAdjustments.Count > 0 && image.MaskPreset.Length > 0)
        {
            for (var index = 0; index < image.MaskPresetAdjustments.Count; index++)
            {
                var value = image.MaskPresetAdjustments[index];
                add("imageMaskAdjustment", value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    JsonValue.Create(value), checked((uint)index), 0);
            }
        }
    }

    private static void DescribeOpaquePictureMask(
        string rawXml,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        // An unsupported mask makes the picture opaque as a whole, but a
        // direct literal `val N` slot is still independently token-spliceable.
        // Parse only the small owner-local geometry envelope; any namespace,
        // child, or attribute ambiguity suppresses the leaf rather than
        // guessing at a partial mask graph.
        if (rawXml.Length is 0 or > 1_000_000) return;
        try
        {
            const string presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
            const string drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace p = presentationNamespace;
            XNamespace a = drawingNamespace;
            // Open XML's OuterXml commonly declares `p` on an ancestor and
            // only repeats `a` on local children.  Re-home the fragment under
            // a synthetic namespace envelope before parsing so a detached
            // opaque `p:pic` remains inspectable without relying on prefixes.
            var envelope = XElement.Parse(
                $"<root xmlns:p=\"{presentationNamespace}\" xmlns:a=\"{drawingNamespace}\">{rawXml}</root>",
                LoadOptions.PreserveWhitespace);
            var root = envelope.Elements(p + "pic").SingleOrDefault();
            if (root is null) return;
            var properties = root.Elements(p + "spPr").SingleOrDefault();
            var geometry = properties?.Elements(a + "prstGeom").SingleOrDefault();
            var preset = geometry?.Attribute("prst");
            if (properties is null || geometry is null || preset is null ||
                geometry.Elements().Count() != 1 ||
                geometry.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName != "prst")) ||
                !PptxCustomGeometryCodec.TryPresetName(new DocumentFormat.OpenXml.Drawing.ShapeTypeValues(preset.Value), out _))
                return;
            var adjustmentList = geometry.Element(a + "avLst");
            if (adjustmentList is null || adjustmentList.Attributes().Any() ||
                adjustmentList.Elements().Count() != adjustmentList.Elements(a + "gd").Count())
                return;
            var guides = adjustmentList.Elements(a + "gd").ToArray();
            for (var index = 0; index < guides.Length; index++)
            {
                var guide = guides[index];
                var name = guide.Attribute("name");
                var formula = guide.Attribute("fmla")?.Value;
                if (guide.Elements().Any() || name is null || formula is null ||
                    guide.Attributes().Any(attribute => attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("name" or "fmla")) ||
                    name.Value.Length == 0 || !TryLiteralAdjustment(formula, out var value))
                    continue;
                add("imageMaskAdjustment", value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    JsonValue.Create(value), checked((uint)index), 0);
            }
        }
        catch (Exception)
        {
            // Opaque source XML is intentionally best-effort here. The
            // source-bound compiler repeats the structural proof against the
            // Open XML tree before accepting any issued leaf.
        }
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

    private static void DescribeGroup(
        PresentationGroup group,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        // A group has two coordinate spaces (outer off/ext and child
        // chOff/chExt).  Child-space coordinates are exposed separately from
        // the outer frame.  They remain a group-owned transaction surface: a
        // child-space edit is never confused with a child element placement.
        if (!source.Editable) return;
        AddInteger(add, "leftEmu", group.LeftEmu);
        AddInteger(add, "topEmu", group.TopEmu);
        AddInteger(add, "widthEmu", group.WidthEmu);
        AddInteger(add, "heightEmu", group.HeightEmu);
        AddInteger(add, "childLeftEmu", group.ChildLeftEmu);
        AddInteger(add, "childTopEmu", group.ChildTopEmu);
        AddInteger(add, "childWidthEmu", group.ChildWidthEmu);
        AddInteger(add, "childHeightEmu", group.ChildHeightEmu);
        if (group.FrameTransform?.HasRotationAngle60000 == true)
            add("rotationDegrees", group.FrameTransform.RotationAngle60000.ToStringInvariant(),
                JsonValue.Create(group.FrameTransform.RotationAngle60000 / 60_000d), 0, 0);
        if (group.FrameTransform?.HasFlipHorizontal == true)
            AddBoolean(add, "flipHorizontal", group.FrameTransform.FlipHorizontal);
        if (group.FrameTransform?.HasFlipVertical == true)
            AddBoolean(add, "flipVertical", group.FrameTransform.FlipVertical);
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
            if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode &&
                properties.AutoFitMode == "shrinkText" && properties.NormalAutoFit is { } normalAutoFit)
            {
                if (normalAutoFit.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000)
                    AddScaled(add, "textBodyNormalAutoFitFontScale", normalAutoFit.FontScale1000 / 1_000d, 1_000, 0);
                if (normalAutoFit.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000)
                    AddScaled(add, "textBodyNormalAutoFitLineSpacingReduction", normalAutoFit.LineSpacingReduction1000 / 1_000d, 1_000, 0);
            }
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
            if (!paragraph.HasBulletColorOpacityThousandthPercent &&
                paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb)
                add("paragraphBulletColorRgb", paragraph.BulletColorRgb, JsonValue.Create($"#{paragraph.BulletColorRgb.ToLowerInvariant()}"), nativeIndex, 0);
            if (!paragraph.HasBulletColorOpacityThousandthPercent &&
                paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme)
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

    private static bool TryLiteralAdjustment(string? formula, out long value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var tokens = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2 || tokens[0] != "val" ||
            !long.TryParse(tokens[1], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed) ||
            parsed is < int.MinValue or > int.MaxValue)
            return false;
        value = parsed;
        return true;
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string ToStringInvariant(this long value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this uint value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
