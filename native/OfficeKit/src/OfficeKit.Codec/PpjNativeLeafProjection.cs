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
        "textBodyVerticalOverflow", "textBodyHorizontalOverflow",
        "textBodyWarpPreset", "customGeometryGuideFormula", "customGeometryAdjustmentFormula",
        "fontFamilyEastAsia", "fontFamilyComplexScript", "fontLanguage", "fontUnderline", "fontStrike", "fontColorScheme", "textGlowColorScheme", "textDefaultGlowColorScheme", "textInnerShadowColorScheme", "textDefaultInnerShadowColorScheme", "shapeGlowColorScheme", "imageGlowColorScheme", "shapeInnerShadowColorScheme", "imageInnerShadowColorScheme",
        "fontCaps", "fontHighlightScheme", "fillScheme", "shadowAlignment", "imageShadowAlignment", "imageShadowColorScheme", "textDefaultShadowAlignment", "shadowColorScheme", "textDefaultShadowColorScheme", "lineScheme", "lineStyle", "lineCap", "lineJoin",
        "lineStartArrow", "lineEndArrow", "lineStartArrowWidth", "lineStartArrowLength", "lineEndArrowWidth", "lineEndArrowLength", "imageMaskPreset", "shape3dPresetMaterial", "shape3dBevelTopPreset", "shape3dBevelBottomPreset", "shape3dSceneCameraPreset", "shape3dSceneLightRigPreset", "shape3dSceneLightRigDirection", "shape3dContourColorScheme", "shape3dExtrusionColorScheme", "chartDataCategory",
    };

    private static readonly HashSet<string> RgbKinds = new(StringComparer.Ordinal)
    {
        "paragraphBulletColorRgb", "fontColorRgb", "fontHighlightRgb", "fillRgb", "shadowColorRgb", "imageShadowColorRgb", "textDefaultShadowColorRgb", "textGlowColorRgb", "textDefaultGlowColorRgb", "textInnerShadowColorRgb", "textDefaultInnerShadowColorRgb", "shapeGlowColorRgb", "imageGlowColorRgb", "shapeInnerShadowColorRgb", "imageInnerShadowColorRgb", "lineRgb", "shape3dContourRgb", "shape3dExtrusionRgb",
    };

    private static readonly HashSet<string> BooleanKinds = new(StringComparer.Ordinal)
    {
        "textBodyColumnDirection", "textBodyUpright", "textBodyAnchorCenter", "textBodyForceAntiAlias", "textBodySpaceFirstLastParagraph", "textBodyCompatibleLineSpacing", "textBodyFromWordArt", "tableBandedRows", "tableBandedColumns", "tableFirstColumnEmphasis", "tableLastColumnEmphasis", "tableLastRow", "fontBold", "fontItalic", "flipHorizontal", "flipVertical", "customGeometryPathFill", "customGeometryPathStroke", "customGeometryPathExtrusionAllowed",
        "textDefaultShadowRotateWithShape", "shadowRotateWithShape", "imageShadowRotateWithShape",
    };

    private static readonly HashSet<string> IntegerKinds = new(StringComparer.Ordinal)
    {
        "paragraphMarginLeftEmu", "paragraphIndentEmu", "paragraphBulletAutoNumberStartAt",
        "paragraphLevel", "textBodyInsetLeftEmu", "textBodyInsetTopEmu",
        "textBodyInsetRightEmu", "textBodyInsetBottomEmu", "textBodyColumnCount", "textBodyColumnGapEmu", "tableHeaderRows",
        "shape3dSceneCameraZoomThousandthPercent", "shape3dSceneCameraFov60000", "shape3dSceneCameraRotationLatitude60000", "shape3dSceneCameraRotationLongitude60000", "shape3dSceneCameraRotationRevolution60000", "shape3dSceneLightRigRotationLatitude60000", "shape3dSceneLightRigRotationLongitude60000", "shape3dSceneLightRigRotationRevolution60000", "shape3dSceneBackdropAnchorXEmu", "shape3dSceneBackdropAnchorYEmu", "shape3dSceneBackdropAnchorZEmu", "shape3dSceneBackdropNormalDxEmu", "shape3dSceneBackdropNormalDyEmu", "shape3dSceneBackdropNormalDzEmu", "shape3dSceneBackdropUpDxEmu", "shape3dSceneBackdropUpDyEmu", "shape3dSceneBackdropUpDzEmu",
        "fillOpacityThousandthPercent", "shadowOpacityThousandthPercent", "imageShadowBlurRadiusEmu", "imageShadowOpacityThousandthPercent", "imageShadowDistanceEmu", "shadowBlurRadiusEmu", "shadowDistanceEmu", "shapeGlowRadiusEmu", "shapeGlowOpacityThousandthPercent", "imageGlowRadiusEmu", "imageGlowOpacityThousandthPercent", "shapeInnerShadowBlurRadiusEmu", "shapeInnerShadowDistanceEmu", "shapeInnerShadowOpacityThousandthPercent", "imageInnerShadowBlurRadiusEmu", "imageInnerShadowDistanceEmu", "imageInnerShadowOpacityThousandthPercent", "shapeReflectionBlurRadiusEmu", "shapeReflectionStartOpacityThousandthPercent", "shapeReflectionEndOpacityThousandthPercent", "shapeReflectionDistanceEmu", "imageReflectionBlurRadiusEmu", "imageReflectionStartOpacityThousandthPercent", "imageReflectionEndOpacityThousandthPercent", "imageReflectionDistanceEmu", "shapeSoftEdgeRadiusEmu", "imageSoftEdgeRadiusEmu", "textGlowRadiusEmu", "textGlowOpacityThousandthPercent", "textDefaultGlowOpacityThousandthPercent",
        "textInnerShadowBlurRadiusEmu", "textInnerShadowDistanceEmu", "textInnerShadowOpacityThousandthPercent", "textDefaultInnerShadowBlurRadiusEmu", "textDefaultInnerShadowDistanceEmu", "textDefaultInnerShadowDirectionDegrees", "textDefaultInnerShadowOpacityThousandthPercent",
        "textReflectionBlurRadiusEmu", "textDefaultReflectionBlurRadiusEmu", "textReflectionStartOpacityThousandthPercent", "textDefaultReflectionStartOpacityThousandthPercent", "textReflectionEndOpacityThousandthPercent", "textDefaultReflectionEndOpacityThousandthPercent", "textReflectionDistanceEmu", "textDefaultReflectionDistanceEmu", "textReflectionDirectionDegrees", "textDefaultReflectionDirectionDegrees", "textSoftEdgeRadiusEmu", "textDefaultSoftEdgeRadiusEmu", "textDefaultGlowRadiusEmu", "textDefaultShadowBlurRadiusEmu", "textDefaultShadowDistanceEmu", "textDefaultShadowDirectionDegrees", "textDefaultShadowOpacityThousandthPercent",
        "imageOpacityThousandthPercent", "lineOpacityThousandthPercent", "lineWidthEmu", "leftEmu", "topEmu",
        "widthEmu", "heightEmu", "childLeftEmu", "childTopEmu",
        "childWidthEmu", "childHeightEmu",
        "customGeometryAdjustment",
        "customGeometryPathLineToX",
        "customGeometryPathLineToY",
        "customGeometryPathMoveToX",
        "customGeometryPathMoveToY",
        "customGeometryPathArcWidthRadius",
        "customGeometryPathArcHeightRadius",
        "customGeometryPathArcStartAngle60000",
        "customGeometryPathArcSweepAngle60000",
        "customGeometryPathQuadraticEndX",
        "customGeometryPathQuadraticEndY",
        "customGeometryPathQuadraticControlX",
        "customGeometryPathQuadraticControlY",
        "customGeometryPathCubicEndX",
        "customGeometryPathCubicEndY",
        "customGeometryPathCubicControl1X",
        "customGeometryPathCubicControl1Y",
        "customGeometryPathCubicControl2X",
        "customGeometryPathCubicControl2Y",
        "presetGeometryAdjustment",
        "imageMaskAdjustment",
        "textBodyWarpAdjustment",
        "textBodyFlatTextZ",
        "shape3dExtrusionHeightEmu",
        "shape3dDepthEmu",
        "shape3dContourWidthEmu",
        "shape3dBevelTopWidthEmu",
        "shape3dBevelTopHeightEmu",
        "shape3dBevelBottomWidthEmu",
        "shape3dBevelBottomHeightEmu",
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
            ["textBodyRotationDegrees"] = 60_000,
            ["rotationDegrees"] = 60_000,
            ["shadowDirectionDegrees"] = 60_000,
            ["imageShadowDirectionDegrees"] = 60_000,
            ["shapeInnerShadowDirectionDegrees"] = 60_000,
            ["imageInnerShadowDirectionDegrees"] = 60_000,
            ["shapeReflectionDirectionDegrees"] = 60_000,
            ["imageReflectionDirectionDegrees"] = 60_000,
            ["textInnerShadowDirectionDegrees"] = 60_000,
            ["textReflectionDirectionDegrees"] = 60_000,
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
        if (kind == "shape3dExtrusionHeightEmu")
        {
            if (!value.TryGetInt64(out var extrusionHeight) || extrusionHeight is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D extrusion height");
            return extrusionHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dDepthEmu")
        {
            if (!value.TryGetInt64(out var depth) || depth is < int.MinValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded signed 32-bit shape 3-D depth");
            return depth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dContourWidthEmu")
        {
            if (!value.TryGetInt64(out var contourWidth) || contourWidth is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D contour width");
            return contourWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraZoomThousandthPercent")
        {
            if (!value.TryGetInt64(out var cameraZoom) || cameraZoom is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D camera zoom in thousandths of one percent");
            return cameraZoom.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraFov60000")
        {
            if (!value.TryGetInt64(out var cameraFov) || cameraFov is <= 0 or >= 180 * 60_000)
                throw InvalidValue(kind, path, "a bounded positive shape 3-D camera field of view below 180 degrees");
            return cameraFov.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelTopWidthEmu")
        {
            if (!value.TryGetInt64(out var bevelTopWidth) || bevelTopWidth is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D top-bevel width");
            return bevelTopWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelTopHeightEmu")
        {
            if (!value.TryGetInt64(out var bevelTopHeight) || bevelTopHeight is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D top-bevel height");
            return bevelTopHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelBottomWidthEmu")
        {
            if (!value.TryGetInt64(out var bevelBottomWidth) || bevelBottomWidth is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D bottom-bevel width");
            return bevelBottomWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dBevelBottomHeightEmu")
        {
            if (!value.TryGetInt64(out var bevelBottomHeight) || bevelBottomHeight is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative shape 3-D bottom-bevel height");
            return bevelBottomHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigRotationLatitude60000")
        {
            if (!value.TryGetInt64(out var latitude) || latitude is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D light-rig rotation latitude below or equal to 360 degrees");
            return latitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigRotationLongitude60000")
        {
            if (!value.TryGetInt64(out var longitude) || longitude is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D light-rig rotation longitude below or equal to 360 degrees");
            return longitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneLightRigRotationRevolution60000")
        {
            if (!value.TryGetInt64(out var revolution) || revolution is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D light-rig rotation revolution below or equal to 360 degrees");
            return revolution.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationLatitude60000")
        {
            if (!value.TryGetInt64(out var latitude) || latitude is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D camera rotation latitude below or equal to 360 degrees");
            return latitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationLongitude60000")
        {
            if (!value.TryGetInt64(out var longitude) || longitude is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D camera rotation longitude below or equal to 360 degrees");
            return longitude.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind == "shape3dSceneCameraRotationRevolution60000")
        {
            if (!value.TryGetInt64(out var revolution) || revolution is < 0 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded non-negative 3-D camera rotation revolution below or equal to 360 degrees");
            return revolution.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind is "shape3dSceneBackdropAnchorXEmu" or "shape3dSceneBackdropAnchorYEmu" or "shape3dSceneBackdropAnchorZEmu" or "shape3dSceneBackdropNormalDxEmu" or "shape3dSceneBackdropNormalDyEmu" or "shape3dSceneBackdropNormalDzEmu" or "shape3dSceneBackdropUpDxEmu" or "shape3dSceneBackdropUpDyEmu" or "shape3dSceneBackdropUpDzEmu")
        {
            if (!value.TryGetInt64(out var backdropCoordinate) ||
                backdropCoordinate < -27_273_042_329_600 || backdropCoordinate > 27_273_042_316_900)
                throw InvalidValue(kind, path, "a bounded signed 3-D backdrop coordinate in EMUs");
            return backdropCoordinate.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        if (kind is "customGeometryAdjustment" or "customGeometryGuide" or "customGeometryConnectionSiteAngle60000" or "customGeometryConnectionSiteXEmu" or "customGeometryConnectionSiteYEmu" or "customGeometryAdjustmentHandleXEmu" or "customGeometryAdjustmentHandleYEmu" or "customGeometryAdjustmentHandleMinXEmu" or "customGeometryAdjustmentHandleMaxXEmu" or "customGeometryAdjustmentHandleMinYEmu" or "customGeometryAdjustmentHandleMaxYEmu" or "customGeometryAdjustmentHandlePolarMinRadiusEmu" or "customGeometryAdjustmentHandlePolarMaxRadiusEmu" or "customGeometryAdjustmentHandlePolarMinAngle60000" or "customGeometryAdjustmentHandlePolarMaxAngle60000" or "customGeometryAdjustmentHandlePolarXEmu" or "customGeometryAdjustmentHandlePolarYEmu" or "customGeometryPathWidth" or "customGeometryPathHeight" or "customGeometryPathArcWidthRadius" or "customGeometryPathArcHeightRadius" or "customGeometryPathArcStartAngle60000" or "customGeometryPathArcSweepAngle60000" or "customGeometryPathLineToX" or "customGeometryPathLineToY" or "customGeometryPathMoveToX" or "customGeometryPathMoveToY" or "customGeometryPathQuadraticEndX" or "customGeometryTextRectangleLeftEmu" or "customGeometryTextRectangleTopEmu" or "customGeometryTextRectangleRightEmu" or "customGeometryTextRectangleBottomEmu" or "presetGeometryAdjustment" or "imageMaskAdjustment")
        {
            if (!value.TryGetInt64(out var adjustment) || adjustment is < int.MinValue or > int.MaxValue ||
                kind is "presetGeometryAdjustment" or "imageMaskAdjustment" &&
                (adjustment < PptxPresetGeometryAdjustmentCodec.MinimumValue || adjustment > PptxPresetGeometryAdjustmentCodec.MaximumValue))
                throw InvalidValue(kind, path, "a bounded DrawingML signed integer");
            if (kind == "customGeometryConnectionSiteAngle60000" && adjustment is < -21_600_000 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded DrawingML connection-site angle");
            if (kind == "customGeometryConnectionSiteXEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry connection-site x coordinate");
            if (kind == "customGeometryConnectionSiteYEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry connection-site y coordinate");
            if (kind == "customGeometryAdjustmentHandleXEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle x coordinate");
            if (kind == "customGeometryAdjustmentHandleYEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle y coordinate");
            if (kind == "customGeometryAdjustmentHandleMinXEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle minimum-x bound");
            if (kind == "customGeometryAdjustmentHandleMaxXEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle maximum-x bound");
            if (kind == "customGeometryAdjustmentHandleMinYEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle minimum-y bound");
            if (kind == "customGeometryAdjustmentHandleMaxYEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry adjustment-handle maximum-y bound");
            if (kind == "customGeometryAdjustmentHandlePolarMinRadiusEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry polar adjustment-handle minimum radius");
            if (kind == "customGeometryAdjustmentHandlePolarMaxRadiusEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry polar adjustment-handle maximum radius");
            if (kind == "customGeometryAdjustmentHandlePolarMinAngle60000" && adjustment is < -21_600_000 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded custom-geometry polar adjustment-handle minimum angle");
            if (kind == "customGeometryAdjustmentHandlePolarMaxAngle60000" && adjustment is < -21_600_000 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded custom-geometry polar adjustment-handle maximum angle");
            if (kind == "customGeometryAdjustmentHandlePolarXEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry polar adjustment-handle x coordinate");
            if (kind == "customGeometryAdjustmentHandlePolarYEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry polar adjustment-handle y coordinate");
            if (kind == "customGeometryPathWidth" && adjustment is <= 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a positive bounded custom-geometry path coordinate width");
            if (kind == "customGeometryPathHeight" && adjustment is <= 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a positive bounded custom-geometry path coordinate height");
            if (kind == "customGeometryPathArcWidthRadius" && adjustment is <= 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a positive bounded custom-geometry arc width radius");
            if (kind == "customGeometryPathArcHeightRadius" && adjustment is <= 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a positive bounded custom-geometry arc height radius");
            if (kind == "customGeometryPathArcStartAngle60000" && adjustment is < -21_600_000 or > 21_600_000)
                throw InvalidValue(kind, path, "a bounded custom-geometry arc start angle");
            if (kind == "customGeometryPathArcSweepAngle60000" && (adjustment == 0 || adjustment is < -21_600_000 or > 21_600_000))
                throw InvalidValue(kind, path, "a bounded non-zero custom-geometry arc sweep angle");
            if (kind == "customGeometryPathLineToX" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry path-coordinate x value");
            if (kind == "customGeometryPathLineToY" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry path-coordinate y value");
            if (kind == "customGeometryPathMoveToX" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry path-coordinate move-to x value");
            if (kind == "customGeometryPathMoveToY" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry path-coordinate move-to y value");
            if (kind == "customGeometryPathQuadraticEndX" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry quadratic end-point x value");
            if (kind == "customGeometryPathQuadraticEndY" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry quadratic end-point y value");
            if (kind == "customGeometryPathQuadraticControlX" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry quadratic control-point x value");
            if (kind == "customGeometryPathQuadraticControlY" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry quadratic control-point y value");
            if (kind == "customGeometryPathCubicEndX" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic end-point x value");
            if (kind == "customGeometryPathCubicEndY" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic end-point y value");
            if (kind == "customGeometryPathCubicControl1X" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic first-control-point x value");
            if (kind == "customGeometryPathCubicControl1Y" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic first-control-point y value");
            if (kind == "customGeometryPathCubicControl2X" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic second-control-point x value");
            if (kind == "customGeometryPathCubicControl2Y" && adjustment is < -int.MaxValue or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded custom-geometry cubic second-control-point y value");
            if (kind == "customGeometryTextRectangleLeftEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry text-rectangle left coordinate");
            if (kind == "customGeometryTextRectangleTopEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry text-rectangle top coordinate");
            if (kind == "customGeometryTextRectangleRightEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry text-rectangle right coordinate");
            if (kind == "customGeometryTextRectangleBottomEmu" && adjustment is < 0 or > int.MaxValue)
                throw InvalidValue(kind, path, "a bounded non-negative custom-geometry text-rectangle bottom coordinate");
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
            if (kind == "shape3dPresetMaterial" && !PptxShape3DCodec.IsPresetMaterialToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML preset-material token");
            if (kind == "shape3dBevelTopPreset" && !PptxShape3DCodec.IsBevelTopPresetToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML bevel-preset token");
            if (kind == "shape3dBevelBottomPreset" && !PptxShape3DCodec.IsBevelBottomPresetToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML bevel-preset token");
            if (kind == "shape3dSceneCameraPreset" && !PptxShape3DCodec.IsSceneCameraPresetToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML camera-preset token");
            if (kind == "shape3dSceneLightRigPreset" && !PptxShape3DCodec.IsSceneLightRigPresetToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML light-rig preset token");
            if (kind == "shape3dSceneLightRigDirection" && !PptxShape3DCodec.IsSceneLightRigDirectionToken(token))
                throw InvalidValue(kind, path, "a canonical DrawingML light-rig direction token");
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
            if (shadow.HasRotateWithShape)
                AddBoolean(add, "shadowRotateWithShape", shadow.RotateWithShape);
        }
        DescribeGlow(shape.Glow, "shapeGlow", add);
        DescribeInnerShadow(shape.InnerShadow, "shapeInnerShadow", add);
        DescribeReflection(shape.Reflection, "shapeReflection", add);
        if (shape.SoftEdge is { HasRadiusEmu: true } softEdge)
            AddInteger(add, "shapeSoftEdgeRadiusEmu", softEdge.RadiusEmu);
        if (shape.HasShapeExtrusionHeightEmu)
            AddInteger(add, "shape3dExtrusionHeightEmu", shape.ShapeExtrusionHeightEmu);
        if (shape.HasShapeDepthEmu)
            AddInteger(add, "shape3dDepthEmu", shape.ShapeDepthEmu);
        if (shape.HasShapeContourWidthEmu)
            AddInteger(add, "shape3dContourWidthEmu", shape.ShapeContourWidthEmu);
        if (shape.HasShape3DBevelTopWidthEmu)
            AddInteger(add, "shape3dBevelTopWidthEmu", shape.Shape3DBevelTopWidthEmu);
        if (shape.HasShape3DBevelTopHeightEmu)
            AddInteger(add, "shape3dBevelTopHeightEmu", shape.Shape3DBevelTopHeightEmu);
        if (!string.IsNullOrEmpty(shape.Shape3DBevelTopPreset))
            add("shape3dBevelTopPreset", shape.Shape3DBevelTopPreset, JsonValue.Create(shape.Shape3DBevelTopPreset), 0, 0);
        if (shape.HasShape3DBevelBottomWidthEmu)
            AddInteger(add, "shape3dBevelBottomWidthEmu", shape.Shape3DBevelBottomWidthEmu);
        if (shape.HasShape3DBevelBottomHeightEmu)
            AddInteger(add, "shape3dBevelBottomHeightEmu", shape.Shape3DBevelBottomHeightEmu);
        if (!string.IsNullOrEmpty(shape.Shape3DBevelBottomPreset))
            add("shape3dBevelBottomPreset", shape.Shape3DBevelBottomPreset, JsonValue.Create(shape.Shape3DBevelBottomPreset), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DContourRgb))
            add("shape3dContourRgb", shape.Shape3DContourRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.Shape3DContourRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DExtrusionRgb))
            add("shape3dExtrusionRgb", shape.Shape3DExtrusionRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.Shape3DExtrusionRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DContourColorScheme))
            add("shape3dContourColorScheme", shape.Shape3DContourColorScheme, JsonValue.Create(shape.Shape3DContourColorScheme), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DExtrusionColorScheme))
            add("shape3dExtrusionColorScheme", shape.Shape3DExtrusionColorScheme, JsonValue.Create(shape.Shape3DExtrusionColorScheme), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DSceneCameraPreset))
            add("shape3dSceneCameraPreset", shape.Shape3DSceneCameraPreset, JsonValue.Create(shape.Shape3DSceneCameraPreset), 0, 0);
        if (shape.HasShape3DSceneCameraZoomThousandthPercent)
            AddInteger(add, "shape3dSceneCameraZoomThousandthPercent", shape.Shape3DSceneCameraZoomThousandthPercent);
        if (shape.HasShape3DSceneCameraFov60000)
            AddInteger(add, "shape3dSceneCameraFov60000", shape.Shape3DSceneCameraFov60000);
        if (!string.IsNullOrEmpty(shape.Shape3DSceneLightRigPreset))
            add("shape3dSceneLightRigPreset", shape.Shape3DSceneLightRigPreset, JsonValue.Create(shape.Shape3DSceneLightRigPreset), 0, 0);
        if (!string.IsNullOrEmpty(shape.Shape3DSceneLightRigDirection))
            add("shape3dSceneLightRigDirection", shape.Shape3DSceneLightRigDirection, JsonValue.Create(shape.Shape3DSceneLightRigDirection), 0, 0);
        if (shape.HasShape3DSceneLightRigRotationLatitude60000)
            AddInteger(add, "shape3dSceneLightRigRotationLatitude60000", shape.Shape3DSceneLightRigRotationLatitude60000);
        if (shape.HasShape3DSceneLightRigRotationLongitude60000)
            AddInteger(add, "shape3dSceneLightRigRotationLongitude60000", shape.Shape3DSceneLightRigRotationLongitude60000);
        if (shape.HasShape3DSceneLightRigRotationRevolution60000)
            AddInteger(add, "shape3dSceneLightRigRotationRevolution60000", shape.Shape3DSceneLightRigRotationRevolution60000);
        if (shape.HasShape3DSceneCameraRotationLatitude60000)
            AddInteger(add, "shape3dSceneCameraRotationLatitude60000", shape.Shape3DSceneCameraRotationLatitude60000);
        if (shape.HasShape3DSceneCameraRotationLongitude60000)
            AddInteger(add, "shape3dSceneCameraRotationLongitude60000", shape.Shape3DSceneCameraRotationLongitude60000);
        if (shape.HasShape3DSceneCameraRotationRevolution60000)
            AddInteger(add, "shape3dSceneCameraRotationRevolution60000", shape.Shape3DSceneCameraRotationRevolution60000);
        if (shape.HasShape3DSceneBackdropAnchorXEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorXEmu", shape.Shape3DSceneBackdropAnchorXEmu);
        if (shape.HasShape3DSceneBackdropAnchorYEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorYEmu", shape.Shape3DSceneBackdropAnchorYEmu);
        if (shape.HasShape3DSceneBackdropAnchorZEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorZEmu", shape.Shape3DSceneBackdropAnchorZEmu);
        if (shape.HasShape3DSceneBackdropNormalDxEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDxEmu", shape.Shape3DSceneBackdropNormalDxEmu);
        if (shape.HasShape3DSceneBackdropNormalDyEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDyEmu", shape.Shape3DSceneBackdropNormalDyEmu);
        if (shape.HasShape3DSceneBackdropNormalDzEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDzEmu", shape.Shape3DSceneBackdropNormalDzEmu);
        if (shape.HasShape3DSceneBackdropUpDxEmu)
            AddInteger(add, "shape3dSceneBackdropUpDxEmu", shape.Shape3DSceneBackdropUpDxEmu);
        if (shape.HasShape3DSceneBackdropUpDyEmu)
            AddInteger(add, "shape3dSceneBackdropUpDyEmu", shape.Shape3DSceneBackdropUpDyEmu);
        if (shape.HasShape3DSceneBackdropUpDzEmu)
            AddInteger(add, "shape3dSceneBackdropUpDzEmu", shape.Shape3DSceneBackdropUpDzEmu);
        if (!string.IsNullOrEmpty(shape.ShapePresetMaterial))
            add("shape3dPresetMaterial", shape.ShapePresetMaterial, JsonValue.Create(shape.ShapePresetMaterial), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineRgb))
            add("lineRgb", shape.LineRgb.ToUpperInvariant(), JsonValue.Create($"#{shape.LineRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(shape.LineScheme))
            add("lineScheme", shape.LineScheme, JsonValue.Create(shape.LineScheme), 0, 0);
        if (shape.HasLineOpacityThousandthPercent)
            AddInteger(add, "lineOpacityThousandthPercent", shape.LineOpacityThousandthPercent);
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
        if (!string.IsNullOrEmpty(shape.StartArrowWidth))
            add("lineStartArrowWidth", shape.StartArrowWidth, JsonValue.Create(shape.StartArrowWidth), 0, 0);
        if (!string.IsNullOrEmpty(shape.StartArrowLength))
            add("lineStartArrowLength", shape.StartArrowLength, JsonValue.Create(shape.StartArrowLength), 0, 0);
        if (!string.IsNullOrEmpty(shape.EndArrowWidth))
            add("lineEndArrowWidth", shape.EndArrowWidth, JsonValue.Create(shape.EndArrowWidth), 0, 0);
        if (!string.IsNullOrEmpty(shape.EndArrowLength))
            add("lineEndArrowLength", shape.EndArrowLength, JsonValue.Create(shape.EndArrowLength), 0, 0);

        // A custom geometry may be fully recognized even when its path
        // coordinates are driven by DrawingML guide formulas.  Keep the
        // formula graph source-owned, but expose the safest adjustment leaf:
        // an existing ordered `a:avLst/a:gd fmla="val N"` can be changed by
        // token splice without changing guide names, path references, or
        // handle/connection topology.  Formula-valued adjustments and all
        // calculated guides remain undisclosed here and therefore fail closed.
        if (source.Editable && shape.Geometry == "custom")
        {
            if (shape.TextRectangle is { } textRectangle &&
                !textRectangle.HasLeftReference && !textRectangle.HasTopReference &&
                !textRectangle.HasRightReference && !textRectangle.HasBottomReference &&
                TryLiteralCustomGeometryTextRectangle(textRectangle, shape.WidthEmu, shape.HeightEmu,
                    out var textRectangleLeft, out var textRectangleTop, out var textRectangleRight, out var textRectangleBottom))
            {
                add("customGeometryTextRectangleLeftEmu", textRectangleLeft.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(textRectangleLeft), 0, 0);
                add("customGeometryTextRectangleTopEmu", textRectangleTop.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(textRectangleTop), 0, 0);
                add("customGeometryTextRectangleRightEmu", textRectangleRight.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(textRectangleRight), 0, 0);
                add("customGeometryTextRectangleBottomEmu", textRectangleBottom.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(textRectangleBottom), 0, 0);
            }

            for (var index = 0; index < shape.CustomPaths.Count; index++)
            {
                if (TryLiteralCustomGeometryPathWidth(shape.CustomPaths[index].Width, out var pathWidth))
                    add("customGeometryPathWidth", pathWidth.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(pathWidth), checked((uint)index), 0);
                if (TryLiteralCustomGeometryPathHeight(shape.CustomPaths[index].Height, out var pathHeight))
                    add("customGeometryPathHeight", pathHeight.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(pathHeight), checked((uint)index), 0);
                var pathFill = shape.CustomPaths[index].FillMode;
                if (pathFill is PresentationCustomGeometryPath.Types.FillMode.Normal or PresentationCustomGeometryPath.Types.FillMode.None)
                    add("customGeometryPathFill", pathFill == PresentationCustomGeometryPath.Types.FillMode.Normal ? "1" : "0", JsonValue.Create(pathFill == PresentationCustomGeometryPath.Types.FillMode.Normal), checked((uint)index), 0);
                if (shape.CustomPaths[index].HasStroke)
                    add("customGeometryPathStroke", shape.CustomPaths[index].Stroke ? "1" : "0", JsonValue.Create(shape.CustomPaths[index].Stroke), checked((uint)index), 0);
                if (shape.CustomPaths[index].HasExtrusionAllowed)
                    add("customGeometryPathExtrusionAllowed", shape.CustomPaths[index].ExtrusionAllowed ? "1" : "0", JsonValue.Create(shape.CustomPaths[index].ExtrusionAllowed), checked((uint)index), 0);
                for (var commandIndex = 0; commandIndex < shape.CustomPaths[index].Commands.Count; commandIndex++)
                {
                    var command = shape.CustomPaths[index].Commands[commandIndex];
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.LineTo &&
                        TryLiteralCustomGeometryPathLineToX(command.LineTo, out var lineToX))
                    {
                        add("customGeometryPathLineToX", lineToX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(lineToX), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.LineTo &&
                        TryLiteralCustomGeometryPathLineToY(command.LineTo, out var lineToY))
                    {
                        add("customGeometryPathLineToY", lineToY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(lineToY), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.MoveTo &&
                        TryLiteralCustomGeometryPathMoveToX(command.MoveTo, out var moveToX))
                    {
                        add("customGeometryPathMoveToX", moveToX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(moveToX), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.MoveTo &&
                        TryLiteralCustomGeometryPathMoveToY(command.MoveTo, out var moveToY))
                    {
                        add("customGeometryPathMoveToY", moveToY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(moveToY), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.ArcTo &&
                        TryLiteralCustomGeometryPathArcWidthRadius(command.ArcTo, out var arcWidthRadius))
                    {
                        add("customGeometryPathArcWidthRadius", arcWidthRadius.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(arcWidthRadius), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.ArcTo &&
                        TryLiteralCustomGeometryPathArcHeightRadius(command.ArcTo, out var arcHeightRadius))
                    {
                        add("customGeometryPathArcHeightRadius", arcHeightRadius.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(arcHeightRadius), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.ArcTo &&
                        TryLiteralCustomGeometryPathArcStartAngle(command.ArcTo, out var arcStartAngle))
                    {
                        add("customGeometryPathArcStartAngle60000", arcStartAngle.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(arcStartAngle), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.ArcTo &&
                        TryLiteralCustomGeometryPathArcSweepAngle(command.ArcTo, out var arcSweepAngle))
                    {
                        add("customGeometryPathArcSweepAngle60000", arcSweepAngle.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(arcSweepAngle), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo &&
                        TryLiteralCustomGeometryPathQuadraticEndX(command.QuadraticBezierTo, out var quadraticEndX))
                    {
                        add("customGeometryPathQuadraticEndX", quadraticEndX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(quadraticEndX), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo &&
                        TryLiteralCustomGeometryPathQuadraticEndY(command.QuadraticBezierTo, out var quadraticEndY))
                    {
                        add("customGeometryPathQuadraticEndY", quadraticEndY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(quadraticEndY), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo &&
                        TryLiteralCustomGeometryPathQuadraticControlX(command.QuadraticBezierTo, out var quadraticControlX))
                    {
                        add("customGeometryPathQuadraticControlX", quadraticControlX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(quadraticControlX), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo &&
                        TryLiteralCustomGeometryPathQuadraticControlY(command.QuadraticBezierTo, out var quadraticControlY))
                    {
                        add("customGeometryPathQuadraticControlY", quadraticControlY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(quadraticControlY), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicEndX(command.CubicBezierTo, out var cubicEndX))
                    {
                        add("customGeometryPathCubicEndX", cubicEndX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicEndX), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicEndY(command.CubicBezierTo, out var cubicEndY))
                    {
                        add("customGeometryPathCubicEndY", cubicEndY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicEndY), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicControl1X(command.CubicBezierTo, out var cubicControl1X))
                    {
                        add("customGeometryPathCubicControl1X", cubicControl1X.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicControl1X), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicControl1Y(command.CubicBezierTo, out var cubicControl1Y))
                    {
                        add("customGeometryPathCubicControl1Y", cubicControl1Y.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicControl1Y), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicControl2X(command.CubicBezierTo, out var cubicControl2X))
                    {
                        add("customGeometryPathCubicControl2X", cubicControl2X.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicControl2X), checked((uint)index), checked((uint)commandIndex));
                    }
                    if (command.CommandCase == PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo &&
                        TryLiteralCustomGeometryPathCubicControl2Y(command.CubicBezierTo, out var cubicControl2Y))
                    {
                        add("customGeometryPathCubicControl2Y", cubicControl2Y.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(cubicControl2Y), checked((uint)index), checked((uint)commandIndex));
                    }
                }
            }

            for (var index = 0; index < shape.CustomAdjustments.Count; index++)
            {
                var formula = shape.CustomAdjustments[index].Formula;
                if (TryLiteralAdjustment(formula, out var value))
                    add("customGeometryAdjustment", value.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(value), checked((uint)index), 0);
                else if (shape.CustomPaths.Count > 0 && formula.Length > 0)
                    add("customGeometryAdjustmentFormula", formula, JsonValue.Create(formula), checked((uint)index), 0);
            }

            // A direct custom-geometry guide has a stable native location in
            // the ordered gdLst. Expose only uniquely named guides, so
            // formula references keep one unambiguous source identity. A
            // literal guide gets the numeric leaf; a calculated guide gets a
            // formula-string leaf. The formula graph, handles, connection
            // sites, text rectangle, paths, and all other guides remain
            // source-owned.
            var guideNames = new HashSet<string>(StringComparer.Ordinal);
            var uniqueGuides = true;
            for (var index = 0; index < shape.CustomGuides.Count; index++)
            {
                if (shape.CustomGuides[index].Name.Length == 0 || !guideNames.Add(shape.CustomGuides[index].Name))
                {
                    uniqueGuides = false;
                    break;
                }
            }
            if (uniqueGuides)
            {
                for (var index = 0; index < shape.CustomGuides.Count; index++)
                {
                    var formula = shape.CustomGuides[index].Formula;
                    if (TryLiteralAdjustment(formula, out var value))
                        add("customGeometryGuide", value.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(value), checked((uint)index), 0);
                    else if (formula.Length > 0)
                        add("customGeometryGuideFormula", formula, JsonValue.Create(formula), checked((uint)index), 0);
                }
            }

            for (var index = 0; index < shape.CustomConnectionSites.Count; index++)
            {
                var site = shape.CustomConnectionSites[index];
                if (!site.HasAngleReference && TryLiteralConnectionSiteAngle(site.Angle60000, out var angle))
                    add("customGeometryConnectionSiteAngle60000", angle.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(angle), checked((uint)index), 0);
                if (!site.HasXReference && TryLiteralConnectionSiteX(site.XEmu, shape.WidthEmu, out var x))
                    add("customGeometryConnectionSiteXEmu", x.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(x), checked((uint)index), 0);
                if (!site.HasYReference && TryLiteralConnectionSiteY(site.YEmu, shape.HeightEmu, out var y))
                    add("customGeometryConnectionSiteYEmu", y.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(y), checked((uint)index), 0);
            }

            for (var index = 0; index < shape.CustomAdjustmentHandles.Count; index++)
            {
                var handle = shape.CustomAdjustmentHandles[index];
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.Position is { } position && !position.HasXReference &&
                    TryLiteralCustomGeometryHandleX(position.X, shape.WidthEmu, out var x))
                    add("customGeometryAdjustmentHandleXEmu", x.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(x), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.Position is { } positionY && !positionY.HasYReference &&
                    TryLiteralCustomGeometryHandleY(positionY.Y, shape.HeightEmu, out var y))
                    add("customGeometryAdjustmentHandleYEmu", y.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(y), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.HasMinX && TryLiteralCustomGeometryHandleMinX(handle.Xy.MinX, shape.WidthEmu, out var minX))
                    add("customGeometryAdjustmentHandleMinXEmu", minX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(minX), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.HasMaxX && TryLiteralCustomGeometryHandleMaxX(handle.Xy.MaxX, shape.WidthEmu, out var maxX))
                    add("customGeometryAdjustmentHandleMaxXEmu", maxX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(maxX), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.HasMinY && TryLiteralCustomGeometryHandleMinY(handle.Xy.MinY, shape.HeightEmu, out var minY))
                    add("customGeometryAdjustmentHandleMinYEmu", minY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(minY), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy &&
                    handle.Xy.HasMaxY && TryLiteralCustomGeometryHandleMaxY(handle.Xy.MaxY, shape.HeightEmu, out var maxY))
                    add("customGeometryAdjustmentHandleMaxYEmu", maxY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(maxY), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.HasMinRadius && TryLiteralCustomGeometryHandlePolarMinRadius(handle.Polar.MinRadius, out var minRadius))
                    add("customGeometryAdjustmentHandlePolarMinRadiusEmu", minRadius.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(minRadius), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.HasMaxRadius && TryLiteralCustomGeometryHandlePolarMaxRadius(handle.Polar.MaxRadius, out var maxRadius))
                    add("customGeometryAdjustmentHandlePolarMaxRadiusEmu", maxRadius.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(maxRadius), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.HasMinAngle60000 && TryLiteralCustomGeometryHandlePolarMinAngle(handle.Polar.MinAngle60000, out var minAngle))
                    add("customGeometryAdjustmentHandlePolarMinAngle60000", minAngle.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(minAngle), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.HasMaxAngle60000 && TryLiteralCustomGeometryHandlePolarMaxAngle(handle.Polar.MaxAngle60000, out var maxAngle))
                    add("customGeometryAdjustmentHandlePolarMaxAngle60000", maxAngle.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(maxAngle), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.Position is { } polarPositionX && !polarPositionX.HasXReference &&
                    TryLiteralCustomGeometryHandleX(polarPositionX.X, shape.WidthEmu, out var polarX))
                    add("customGeometryAdjustmentHandlePolarXEmu", polarX.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(polarX), checked((uint)index), 0);
                if (handle.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar &&
                    handle.Polar.Position is { } polarPositionY && !polarPositionY.HasYReference &&
                    TryLiteralCustomGeometryHandleY(polarPositionY.Y, shape.HeightEmu, out var polarY))
                    add("customGeometryAdjustmentHandlePolarYEmu", polarY.ToString(System.Globalization.CultureInfo.InvariantCulture), JsonValue.Create(polarY), checked((uint)index), 0);
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
        if (image.HasShape3DDepthEmu)
            AddInteger(add, "shape3dDepthEmu", image.Shape3DDepthEmu);
        if (image.HasShape3DExtrusionHeightEmu)
            AddInteger(add, "shape3dExtrusionHeightEmu", image.Shape3DExtrusionHeightEmu);
        if (image.HasShape3DContourWidthEmu)
            AddInteger(add, "shape3dContourWidthEmu", image.Shape3DContourWidthEmu);
        if (!string.IsNullOrEmpty(image.Shape3DPresetMaterial))
            add("shape3dPresetMaterial", image.Shape3DPresetMaterial, JsonValue.Create(image.Shape3DPresetMaterial), 0, 0);
        if (image.HasShape3DBevelTopWidthEmu)
            AddInteger(add, "shape3dBevelTopWidthEmu", image.Shape3DBevelTopWidthEmu);
        if (image.HasShape3DBevelTopHeightEmu)
            AddInteger(add, "shape3dBevelTopHeightEmu", image.Shape3DBevelTopHeightEmu);
        if (!string.IsNullOrEmpty(image.Shape3DBevelTopPreset))
            add("shape3dBevelTopPreset", image.Shape3DBevelTopPreset, JsonValue.Create(image.Shape3DBevelTopPreset), 0, 0);
        if (image.HasShape3DBevelBottomWidthEmu)
            AddInteger(add, "shape3dBevelBottomWidthEmu", image.Shape3DBevelBottomWidthEmu);
        if (image.HasShape3DBevelBottomHeightEmu)
            AddInteger(add, "shape3dBevelBottomHeightEmu", image.Shape3DBevelBottomHeightEmu);
        if (!string.IsNullOrEmpty(image.Shape3DBevelBottomPreset))
            add("shape3dBevelBottomPreset", image.Shape3DBevelBottomPreset, JsonValue.Create(image.Shape3DBevelBottomPreset), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DContourRgb))
            add("shape3dContourRgb", image.Shape3DContourRgb.ToUpperInvariant(), JsonValue.Create($"#{image.Shape3DContourRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionRgb))
            add("shape3dExtrusionRgb", image.Shape3DExtrusionRgb.ToUpperInvariant(), JsonValue.Create($"#{image.Shape3DExtrusionRgb.ToLowerInvariant()}"), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DContourColorScheme))
            add("shape3dContourColorScheme", image.Shape3DContourColorScheme, JsonValue.Create(image.Shape3DContourColorScheme), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionColorScheme))
            add("shape3dExtrusionColorScheme", image.Shape3DExtrusionColorScheme, JsonValue.Create(image.Shape3DExtrusionColorScheme), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DSceneCameraPreset))
            add("shape3dSceneCameraPreset", image.Shape3DSceneCameraPreset, JsonValue.Create(image.Shape3DSceneCameraPreset), 0, 0);
        if (image.HasShape3DSceneCameraZoomThousandthPercent)
            AddInteger(add, "shape3dSceneCameraZoomThousandthPercent", image.Shape3DSceneCameraZoomThousandthPercent);
        if (image.HasShape3DSceneCameraFov60000)
            AddInteger(add, "shape3dSceneCameraFov60000", image.Shape3DSceneCameraFov60000);
        if (image.HasShape3DSceneCameraRotationLatitude60000)
            AddInteger(add, "shape3dSceneCameraRotationLatitude60000", image.Shape3DSceneCameraRotationLatitude60000);
        if (image.HasShape3DSceneCameraRotationLongitude60000)
            AddInteger(add, "shape3dSceneCameraRotationLongitude60000", image.Shape3DSceneCameraRotationLongitude60000);
        if (image.HasShape3DSceneCameraRotationRevolution60000)
            AddInteger(add, "shape3dSceneCameraRotationRevolution60000", image.Shape3DSceneCameraRotationRevolution60000);
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigPreset))
            add("shape3dSceneLightRigPreset", image.Shape3DSceneLightRigPreset, JsonValue.Create(image.Shape3DSceneLightRigPreset), 0, 0);
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigDirection))
            add("shape3dSceneLightRigDirection", image.Shape3DSceneLightRigDirection, JsonValue.Create(image.Shape3DSceneLightRigDirection), 0, 0);
        if (image.HasShape3DSceneLightRigRotationLatitude60000)
            AddInteger(add, "shape3dSceneLightRigRotationLatitude60000", image.Shape3DSceneLightRigRotationLatitude60000);
        if (image.HasShape3DSceneLightRigRotationLongitude60000)
            AddInteger(add, "shape3dSceneLightRigRotationLongitude60000", image.Shape3DSceneLightRigRotationLongitude60000);
        if (image.HasShape3DSceneLightRigRotationRevolution60000)
            AddInteger(add, "shape3dSceneLightRigRotationRevolution60000", image.Shape3DSceneLightRigRotationRevolution60000);
        if (image.HasShape3DSceneBackdropAnchorXEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorXEmu", image.Shape3DSceneBackdropAnchorXEmu);
        if (image.HasShape3DSceneBackdropAnchorYEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorYEmu", image.Shape3DSceneBackdropAnchorYEmu);
        if (image.HasShape3DSceneBackdropAnchorZEmu)
            AddInteger(add, "shape3dSceneBackdropAnchorZEmu", image.Shape3DSceneBackdropAnchorZEmu);
        if (image.HasShape3DSceneBackdropNormalDxEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDxEmu", image.Shape3DSceneBackdropNormalDxEmu);
        if (image.HasShape3DSceneBackdropNormalDyEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDyEmu", image.Shape3DSceneBackdropNormalDyEmu);
        if (image.HasShape3DSceneBackdropNormalDzEmu)
            AddInteger(add, "shape3dSceneBackdropNormalDzEmu", image.Shape3DSceneBackdropNormalDzEmu);
        if (image.HasShape3DSceneBackdropUpDxEmu)
            AddInteger(add, "shape3dSceneBackdropUpDxEmu", image.Shape3DSceneBackdropUpDxEmu);
        if (image.HasShape3DSceneBackdropUpDyEmu)
            AddInteger(add, "shape3dSceneBackdropUpDyEmu", image.Shape3DSceneBackdropUpDyEmu);
        if (image.HasShape3DSceneBackdropUpDzEmu)
            AddInteger(add, "shape3dSceneBackdropUpDzEmu", image.Shape3DSceneBackdropUpDzEmu);
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
        if (image.Shadow is { HasRotateWithShape: true } shadow)
            AddBoolean(add, "imageShadowRotateWithShape", shadow.RotateWithShape);
        if (image.Shadow is { HasBlurRadiusEmu: true } shadowBlur)
            AddInteger(add, "imageShadowBlurRadiusEmu", shadowBlur.BlurRadiusEmu);
        if (image.Shadow is { HasDistanceEmu: true } shadowDistance)
            AddInteger(add, "imageShadowDistanceEmu", shadowDistance.DistanceEmu);
        if (image.Shadow is { HasDirectionAngle60000: true } shadowDirection)
            add("imageShadowDirectionDegrees", shadowDirection.DirectionAngle60000.ToStringInvariant(),
                JsonValue.Create(shadowDirection.DirectionAngle60000 / 60_000d), 0, 0);
        if (image.Shadow is { HasAlignment: true } shadowAlignment)
            add("imageShadowAlignment", shadowAlignment.Alignment,
                JsonValue.Create(shadowAlignment.Alignment), 0, 0);
        if (image.Shadow is { ColorRgb.Length: > 0 } shadowColor)
            add("imageShadowColorRgb", shadowColor.ColorRgb.ToUpperInvariant(),
                JsonValue.Create($"#{shadowColor.ColorRgb.ToLowerInvariant()}"), 0, 0);
        else if (image.Shadow is { HasColorScheme: true } shadowScheme &&
                 !string.IsNullOrEmpty(shadowScheme.ColorScheme))
            add("imageShadowColorScheme", shadowScheme.ColorScheme,
                JsonValue.Create(shadowScheme.ColorScheme), 0, 0);
        if (image.Shadow is { HasOpacityThousandthPercent: true } shadowOpacity)
            AddInteger(add, "imageShadowOpacityThousandthPercent", shadowOpacity.OpacityThousandthPercent);
        DescribeGlow(image.Glow, "imageGlow", add);
        DescribeInnerShadow(image.InnerShadow, "imageInnerShadow", add);
        DescribeReflection(image.Reflection, "imageReflection", add);
        if (image.SoftEdge is { HasRadiusEmu: true } softEdge)
            AddInteger(add, "imageSoftEdgeRadiusEmu", softEdge.RadiusEmu);
    }

    private static void DescribeGlow(
        PresentationGlow? glow,
        string prefix,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (glow is null) return;
        if (glow.HasRadiusEmu)
            AddInteger(add, prefix + "RadiusEmu", glow.RadiusEmu);
        if (!string.IsNullOrEmpty(glow.ColorRgb))
            add(prefix + "ColorRgb", glow.ColorRgb.ToUpperInvariant(),
                JsonValue.Create($"#{glow.ColorRgb.ToLowerInvariant()}"), 0, 0);
        else if (glow.HasColorScheme && !string.IsNullOrEmpty(glow.ColorScheme))
            add(prefix + "ColorScheme", glow.ColorScheme, JsonValue.Create(glow.ColorScheme), 0, 0);
        if (glow.HasOpacityThousandthPercent)
            AddInteger(add, prefix + "OpacityThousandthPercent", glow.OpacityThousandthPercent);
    }

    private static void DescribeInnerShadow(
        PresentationInnerShadow? shadow,
        string prefix,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (shadow is null) return;
        if (shadow.HasBlurRadiusEmu)
            AddInteger(add, prefix + "BlurRadiusEmu", shadow.BlurRadiusEmu);
        if (shadow.HasDistanceEmu)
            AddInteger(add, prefix + "DistanceEmu", shadow.DistanceEmu);
        if (shadow.HasDirectionAngle60000)
            add(prefix + "DirectionDegrees", shadow.DirectionAngle60000.ToStringInvariant(),
                JsonValue.Create(shadow.DirectionAngle60000 / 60_000d), 0, 0);
        if (!string.IsNullOrEmpty(shadow.ColorRgb))
            add(prefix + "ColorRgb", shadow.ColorRgb.ToUpperInvariant(),
                JsonValue.Create($"#{shadow.ColorRgb.ToLowerInvariant()}"), 0, 0);
        else if (shadow.HasColorScheme && !string.IsNullOrEmpty(shadow.ColorScheme))
            add(prefix + "ColorScheme", shadow.ColorScheme, JsonValue.Create(shadow.ColorScheme), 0, 0);
        if (shadow.HasOpacityThousandthPercent)
            AddInteger(add, prefix + "OpacityThousandthPercent", shadow.OpacityThousandthPercent);
    }

    private static void DescribeReflection(
        PresentationReflection? reflection,
        string prefix,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (reflection is null) return;
        if (reflection.HasBlurRadiusEmu)
            AddInteger(add, prefix + "BlurRadiusEmu", reflection.BlurRadiusEmu);
        if (reflection.HasStartOpacityThousandthPercent)
            AddInteger(add, prefix + "StartOpacityThousandthPercent", reflection.StartOpacityThousandthPercent);
        if (reflection.HasEndOpacityThousandthPercent)
            AddInteger(add, prefix + "EndOpacityThousandthPercent", reflection.EndOpacityThousandthPercent);
        if (reflection.HasDistanceEmu)
            AddInteger(add, prefix + "DistanceEmu", reflection.DistanceEmu);
        if (reflection.HasDirectionAngle60000)
            add(prefix + "DirectionDegrees", reflection.DirectionAngle60000.ToStringInvariant(),
                JsonValue.Create(reflection.DirectionAngle60000 / 60_000d), 0, 0);
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
        if (connector.HasLineOpacityThousandthPercent)
            AddInteger(add, "lineOpacityThousandthPercent", connector.LineOpacityThousandthPercent);
        if (connector.LineWidthEmu > 0) AddInteger(add, "lineWidthEmu", connector.LineWidthEmu);
        if (!string.IsNullOrEmpty(connector.LineStyle)) add("lineStyle", connector.LineStyle, JsonValue.Create(connector.LineStyle), 0, 0);
        if (!string.IsNullOrEmpty(connector.LineCap)) add("lineCap", connector.LineCap, JsonValue.Create(connector.LineCap), 0, 0);
        if (!string.IsNullOrEmpty(connector.LineJoin)) add("lineJoin", connector.LineJoin, JsonValue.Create(connector.LineJoin), 0, 0);
        if (!string.IsNullOrEmpty(connector.StartArrow)) add("lineStartArrow", connector.StartArrow, JsonValue.Create(connector.StartArrow), 0, 0);
        if (!string.IsNullOrEmpty(connector.EndArrow)) add("lineEndArrow", connector.EndArrow, JsonValue.Create(connector.EndArrow), 0, 0);
        if (!string.IsNullOrEmpty(connector.StartArrowWidth)) add("lineStartArrowWidth", connector.StartArrowWidth, JsonValue.Create(connector.StartArrowWidth), 0, 0);
        if (!string.IsNullOrEmpty(connector.StartArrowLength)) add("lineStartArrowLength", connector.StartArrowLength, JsonValue.Create(connector.StartArrowLength), 0, 0);
        if (!string.IsNullOrEmpty(connector.EndArrowWidth)) add("lineEndArrowWidth", connector.EndArrowWidth, JsonValue.Create(connector.EndArrowWidth), 0, 0);
        if (!string.IsNullOrEmpty(connector.EndArrowLength)) add("lineEndArrowLength", connector.EndArrowLength, JsonValue.Create(connector.EndArrowLength), 0, 0);
    }

    private static void DescribeTable(
        PresentationTable table,
        PresentationElementSourceBinding source,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (!source.Editable) return;
        if (table.HasFirstRow)
            AddInteger(add, "tableHeaderRows", table.FirstRow ? 1 : 0);
        if (table.HasBandedRows)
            AddBoolean(add, "tableBandedRows", table.BandedRows);
        if (table.HasBandedColumns)
            AddBoolean(add, "tableBandedColumns", table.BandedColumns);
        if (table.HasFirstColumn)
            AddBoolean(add, "tableFirstColumnEmphasis", table.FirstColumn);
        if (table.HasLastColumn)
            AddBoolean(add, "tableLastColumnEmphasis", table.LastColumn);
        if (table.HasLastRow)
            AddBoolean(add, "tableLastRow", table.LastRow);
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
            if (properties.HasAnchorCenter)
                AddBoolean(add, "textBodyAnchorCenter", properties.AnchorCenter);
            if (properties.HasForceAntiAlias)
                AddBoolean(add, "textBodyForceAntiAlias", properties.ForceAntiAlias);
            if (properties.HasSpaceFirstLastParagraph)
                AddBoolean(add, "textBodySpaceFirstLastParagraph", properties.SpaceFirstLastParagraph);
            if (properties.HasCompatibleLineSpacing)
                AddBoolean(add, "textBodyCompatibleLineSpacing", properties.CompatibleLineSpacing);
            if (properties.HasFromWordArt)
                AddBoolean(add, "textBodyFromWordArt", properties.FromWordArt);
            if (properties.HasTextWarpPreset)
                add("textBodyWarpPreset", properties.TextWarpPreset, JsonValue.Create(properties.TextWarpPreset), 0, 0);
            for (var index = 0; index < properties.TextWarpAdjustments.Count; index++)
                AddInteger(add, "textBodyWarpAdjustment", properties.TextWarpAdjustments[index].Value, checked((uint)index));
            if (properties.HasFlatTextZ)
                AddInteger(add, "textBodyFlatTextZ", properties.FlatTextZ);
            AddBodyInset(properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, "textBodyInsetLeftEmu", properties.LeftInsetEmu, add);
            AddBodyInset(properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, "textBodyInsetTopEmu", properties.TopInsetEmu, add);
            AddBodyInset(properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, "textBodyInsetRightEmu", properties.RightInsetEmu, add);
            AddBodyInset(properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, "textBodyInsetBottomEmu", properties.BottomInsetEmu, add);
            if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
                add("textBodyWrap", properties.Wrap, JsonValue.Create(properties.Wrap), 0, 0);
            if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
                add("textBodyColumnCount", properties.Columns.ToStringInvariant(), JsonValue.Create(properties.Columns), 0, 0);
            if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu)
                AddInteger(add, "textBodyColumnGapEmu", properties.ColumnSpacingEmu);
            if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000)
                add("textBodyRotationDegrees", properties.RotationAngle60000.ToStringInvariant(), JsonValue.Create(properties.RotationAngle60000 / 60_000d), 0, 0);
            if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode)
                add("textBodyVerticalOverflow", properties.VerticalOverflowMode, JsonValue.Create(properties.VerticalOverflowMode), 0, 0);
            if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode)
                add("textBodyHorizontalOverflow", properties.HorizontalOverflowMode, JsonValue.Create(properties.HorizontalOverflowMode), 0, 0);
            if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright)
                AddBoolean(add, "textBodyUpright", properties.Upright);
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
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.SoftEdge is { HasRadiusEmu: true } softEdge)
                AddInteger(add, "textDefaultSoftEdgeRadiusEmu", softEdge.RadiusEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasBlurRadiusEmu: true } shadow)
                AddInteger(add, "textDefaultShadowBlurRadiusEmu", shadow.BlurRadiusEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasDistanceEmu: true } shadowDistance)
                AddInteger(add, "textDefaultShadowDistanceEmu", shadowDistance.DistanceEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasDirectionAngle60000: true } shadowDirection)
                AddInteger(add, "textDefaultShadowDirectionDegrees", shadowDirection.DirectionAngle60000, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasAlignment: true } shadowAlignment)
                add("textDefaultShadowAlignment", shadowAlignment.Alignment,
                    JsonValue.Create(shadowAlignment.Alignment), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { ColorRgb.Length: > 0 } shadowColor)
                add("textDefaultShadowColorRgb", shadowColor.ColorRgb.ToUpperInvariant(),
                    JsonValue.Create($"#{shadowColor.ColorRgb.ToLowerInvariant()}"), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasColorScheme: true } shadowScheme &&
                !string.IsNullOrEmpty(shadowScheme.ColorScheme))
                add("textDefaultShadowColorScheme", shadowScheme.ColorScheme,
                    JsonValue.Create(shadowScheme.ColorScheme), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasOpacityThousandthPercent: true } shadowOpacity)
                AddInteger(add, "textDefaultShadowOpacityThousandthPercent", shadowOpacity.OpacityThousandthPercent, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Shadow is { HasRotateWithShape: true } shadowRotation)
                AddBoolean(add, "textDefaultShadowRotateWithShape", shadowRotation.RotateWithShape, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Glow is { HasRadiusEmu: true } glow)
                AddInteger(add, "textDefaultGlowRadiusEmu", glow.RadiusEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Glow is { HasOpacityThousandthPercent: true } glowOpacity)
                AddInteger(add, "textDefaultGlowOpacityThousandthPercent", glowOpacity.OpacityThousandthPercent, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Glow is { ColorRgb.Length: > 0 } glowColor)
                add("textDefaultGlowColorRgb", glowColor.ColorRgb.ToUpperInvariant(),
                    JsonValue.Create($"#{glowColor.ColorRgb.ToLowerInvariant()}"), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Glow is { HasColorScheme: true } glowScheme &&
                !string.IsNullOrEmpty(glowScheme.ColorScheme))
                add("textDefaultGlowColorScheme", glowScheme.ColorScheme,
                    JsonValue.Create(glowScheme.ColorScheme), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { HasBlurRadiusEmu: true } innerShadow)
                AddInteger(add, "textDefaultInnerShadowBlurRadiusEmu", innerShadow.BlurRadiusEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { HasDistanceEmu: true } innerShadowDistance)
                AddInteger(add, "textDefaultInnerShadowDistanceEmu", innerShadowDistance.DistanceEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { HasDirectionAngle60000: true } innerShadowDirection)
                AddInteger(add, "textDefaultInnerShadowDirectionDegrees", innerShadowDirection.DirectionAngle60000, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { ColorRgb.Length: > 0 } innerShadowColor)
                add("textDefaultInnerShadowColorRgb", innerShadowColor.ColorRgb.ToUpperInvariant(),
                    JsonValue.Create($"#{innerShadowColor.ColorRgb.ToLowerInvariant()}"), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { HasColorScheme: true } innerShadowScheme &&
                !string.IsNullOrEmpty(innerShadowScheme.ColorScheme))
                add("textDefaultInnerShadowColorScheme", innerShadowScheme.ColorScheme,
                    JsonValue.Create(innerShadowScheme.ColorScheme), nativeIndex, 0);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.InnerShadow is { HasOpacityThousandthPercent: true } innerShadowOpacity)
                AddInteger(add, "textDefaultInnerShadowOpacityThousandthPercent", innerShadowOpacity.OpacityThousandthPercent, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Reflection is { HasBlurRadiusEmu: true } reflection)
                AddInteger(add, "textDefaultReflectionBlurRadiusEmu", reflection.BlurRadiusEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Reflection is { HasDistanceEmu: true } reflectionDistance)
                AddInteger(add, "textDefaultReflectionDistanceEmu", reflectionDistance.DistanceEmu, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Reflection is { HasStartOpacityThousandthPercent: true } reflectionStartOpacity)
                AddInteger(add, "textDefaultReflectionStartOpacityThousandthPercent", reflectionStartOpacity.StartOpacityThousandthPercent, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Reflection is { HasEndOpacityThousandthPercent: true } reflectionEndOpacity)
                AddInteger(add, "textDefaultReflectionEndOpacityThousandthPercent", reflectionEndOpacity.EndOpacityThousandthPercent, nativeIndex);
            if (paragraph.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                paragraph.DefaultRunProperties.Reflection is { HasDirectionAngle60000: true } reflectionDirection)
                AddInteger(add, "textDefaultReflectionDirectionDegrees", reflectionDirection.DirectionAngle60000, nativeIndex);

            foreach (var run in paragraph.Runs)
            {
                if (run.ContentCase == PresentationTextRun.ContentOneofCase.Text)
                {
                    add("text", run.Text, JsonValue.Create(run.Text), 0, textLeafIndex);
                    DescribeGlow(run, runStyleIndex, textLeafIndex, add);
                    DescribeInnerShadow(run, runStyleIndex, textLeafIndex, add);
                    DescribeReflection(run, runStyleIndex, textLeafIndex, add);
                    DescribeSoftEdge(run, runStyleIndex, textLeafIndex, add);
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
        if (run.HasFontFamilyComplexScript) add("fontFamilyComplexScript", run.FontFamilyComplexScript, JsonValue.Create(run.FontFamilyComplexScript), index, 0);
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

    private static void DescribeGlow(
        PresentationTextRun run,
        uint runIndex,
        uint textIndex,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (run.Glow is not { } glow) return;
        if (glow.HasRadiusEmu)
            AddInteger(add, "textGlowRadiusEmu", glow.RadiusEmu, runIndex, textIndex);
        if (!string.IsNullOrEmpty(glow.ColorRgb))
            add("textGlowColorRgb", glow.ColorRgb.ToUpperInvariant(),
                JsonValue.Create($"#{glow.ColorRgb.ToLowerInvariant()}"), runIndex, textIndex);
        else if (glow.HasColorScheme && !string.IsNullOrEmpty(glow.ColorScheme))
            add("textGlowColorScheme", glow.ColorScheme,
                JsonValue.Create(glow.ColorScheme), runIndex, textIndex);
            if (glow.HasOpacityThousandthPercent)
            AddInteger(add, "textGlowOpacityThousandthPercent",
                glow.OpacityThousandthPercent, runIndex, textIndex);
    }

    private static void DescribeInnerShadow(
        PresentationTextRun run,
        uint runIndex,
        uint textIndex,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (run.InnerShadow is not { } shadow) return;
        if (shadow.HasBlurRadiusEmu)
            AddInteger(add, "textInnerShadowBlurRadiusEmu", shadow.BlurRadiusEmu, runIndex, textIndex);
        if (shadow.HasDistanceEmu)
            AddInteger(add, "textInnerShadowDistanceEmu", shadow.DistanceEmu, runIndex, textIndex);
        if (shadow.HasDirectionAngle60000)
            add("textInnerShadowDirectionDegrees", shadow.DirectionAngle60000.ToStringInvariant(),
                JsonValue.Create(shadow.DirectionAngle60000 / 60_000d), runIndex, textIndex);
        if (!string.IsNullOrEmpty(shadow.ColorRgb))
            add("textInnerShadowColorRgb", shadow.ColorRgb.ToUpperInvariant(),
                JsonValue.Create($"#{shadow.ColorRgb.ToLowerInvariant()}"), runIndex, textIndex);
        else if (shadow.HasColorScheme && !string.IsNullOrEmpty(shadow.ColorScheme))
            add("textInnerShadowColorScheme", shadow.ColorScheme,
                JsonValue.Create(shadow.ColorScheme), runIndex, textIndex);
        if (shadow.HasOpacityThousandthPercent)
            AddInteger(add, "textInnerShadowOpacityThousandthPercent",
                shadow.OpacityThousandthPercent, runIndex, textIndex);
    }

    private static void DescribeReflection(
        PresentationTextRun run,
        uint runIndex,
        uint textIndex,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (run.Reflection is not { } reflection) return;
        if (reflection.HasBlurRadiusEmu)
            AddInteger(add, "textReflectionBlurRadiusEmu", reflection.BlurRadiusEmu, runIndex, textIndex);
        if (reflection.HasStartOpacityThousandthPercent)
            AddInteger(add, "textReflectionStartOpacityThousandthPercent",
                reflection.StartOpacityThousandthPercent, runIndex, textIndex);
        if (reflection.HasEndOpacityThousandthPercent)
            AddInteger(add, "textReflectionEndOpacityThousandthPercent",
                reflection.EndOpacityThousandthPercent, runIndex, textIndex);
        if (reflection.HasDistanceEmu)
            AddInteger(add, "textReflectionDistanceEmu", reflection.DistanceEmu, runIndex, textIndex);
        if (reflection.HasDirectionAngle60000)
            add("textReflectionDirectionDegrees", reflection.DirectionAngle60000.ToStringInvariant(),
                JsonValue.Create(reflection.DirectionAngle60000 / 60_000d), runIndex, textIndex);
    }

    private static void DescribeSoftEdge(
        PresentationTextRun run,
        uint runIndex,
        uint textIndex,
        Action<string, string, JsonNode?, uint, uint> add)
    {
        if (run.SoftEdge is { HasRadiusEmu: true } softEdge)
            AddInteger(add, "textSoftEdgeRadiusEmu", softEdge.RadiusEmu, runIndex, textIndex);
    }

    private static void AddBodyInset(bool present, string kind, long value, Action<string, string, JsonNode?, uint, uint> add)
    {
        if (present) AddInteger(add, kind, value);
    }

    private static void AddInteger(
        Action<string, string, JsonNode?, uint, uint> add,
        string kind,
        long value,
        uint nativeIndex = 0,
        uint textIndex = 0) =>
        add(kind, value.ToStringInvariant(), JsonValue.Create(value), nativeIndex, textIndex);

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

    private static bool TryLiteralConnectionSiteAngle(int angle, out int value)
    {
        value = angle;
        return angle is >= -21_600_000 and <= 21_600_000;
    }

    private static bool TryLiteralConnectionSiteX(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralConnectionSiteY(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleX(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleY(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleMinX(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleMaxX(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleMinY(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandleMaxY(long coordinate, long extent, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue && coordinate <= extent;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMinRadius(long coordinate, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMaxRadius(long coordinate, out long value)
    {
        value = coordinate;
        return coordinate is >= 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMinAngle(long angle, out long value)
    {
        value = angle;
        return angle is >= -21_600_000 and <= 21_600_000;
    }

    private static bool TryLiteralCustomGeometryHandlePolarMaxAngle(long angle, out long value)
    {
        value = angle;
        return angle is >= -21_600_000 and <= 21_600_000;
    }

    private static bool TryLiteralCustomGeometryTextRectangle(
        PresentationCustomGeometryTextRectangle rectangle,
        long width,
        long height,
        out long left,
        out long top,
        out long right,
        out long bottom)
    {
        left = rectangle.LeftEmu;
        top = rectangle.TopEmu;
        right = rectangle.RightEmu;
        bottom = rectangle.BottomEmu;
        return rectangle.LeftEmu is >= 0 and <= int.MaxValue && rectangle.LeftEmu <= width &&
            rectangle.TopEmu is >= 0 and <= int.MaxValue && rectangle.TopEmu <= height &&
            rectangle.RightEmu is >= 0 and <= int.MaxValue && rectangle.RightEmu <= width &&
            rectangle.BottomEmu is >= 0 and <= int.MaxValue && rectangle.BottomEmu <= height &&
            rectangle.LeftEmu < rectangle.RightEmu && rectangle.TopEmu < rectangle.BottomEmu;
    }

    private static bool TryLiteralCustomGeometryPathWidth(long width, out long value)
    {
        value = width;
        return width is > 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathHeight(long height, out long value)
    {
        value = height;
        return height is > 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathLineToX(
        PresentationCustomGeometryPoint point,
        out long value)
    {
        value = point.X;
        return !point.HasXReference && point.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathLineToY(
        PresentationCustomGeometryPoint point,
        out long value)
    {
        value = point.Y;
        return !point.HasYReference && point.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathMoveToX(
        PresentationCustomGeometryPoint point,
        out long value)
    {
        value = point.X;
        return !point.HasXReference && point.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathMoveToY(
        PresentationCustomGeometryPoint point,
        out long value)
    {
        value = point.Y;
        return !point.HasYReference && point.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathArcWidthRadius(
        PresentationCustomGeometryArc arc,
        out long value)
    {
        value = arc.WidthRadius;
        return !arc.HasWidthRadiusReference && arc.WidthRadius is > 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathArcHeightRadius(
        PresentationCustomGeometryArc arc,
        out long value)
    {
        value = arc.HeightRadius;
        return !arc.HasHeightRadiusReference && arc.HeightRadius is > 0 and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathArcStartAngle(
        PresentationCustomGeometryArc arc,
        out long value)
    {
        value = arc.StartAngle;
        return !arc.HasStartAngleReference && arc.StartAngle is >= -21_600_000 and <= 21_600_000;
    }

    private static bool TryLiteralCustomGeometryPathArcSweepAngle(
        PresentationCustomGeometryArc arc,
        out long value)
    {
        value = arc.SweepAngle;
        return !arc.HasSweepAngleReference && arc.SweepAngle != 0 && arc.SweepAngle is >= -21_600_000 and <= 21_600_000;
    }

    private static bool TryLiteralCustomGeometryPathQuadraticEndX(
        PresentationCustomGeometryQuadraticBezier quadratic,
        out long value)
    {
        value = quadratic.End.X;
        return !quadratic.End.HasXReference && quadratic.End.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathQuadraticEndY(
        PresentationCustomGeometryQuadraticBezier quadratic,
        out long value)
    {
        value = quadratic.End.Y;
        return !quadratic.End.HasYReference && quadratic.End.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathQuadraticControlX(
        PresentationCustomGeometryQuadraticBezier quadratic,
        out long value)
    {
        value = quadratic.Control.X;
        return !quadratic.Control.HasXReference && quadratic.Control.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathQuadraticControlY(
        PresentationCustomGeometryQuadraticBezier quadratic,
        out long value)
    {
        value = quadratic.Control.Y;
        return !quadratic.Control.HasYReference && quadratic.Control.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicEndX(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.End.X;
        return !cubic.End.HasXReference && cubic.End.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicEndY(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.End.Y;
        return !cubic.End.HasYReference && cubic.End.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicControl1X(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.Control1.X;
        return !cubic.Control1.HasXReference && cubic.Control1.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicControl1Y(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.Control1.Y;
        return !cubic.Control1.HasYReference && cubic.Control1.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicControl2X(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.Control2.X;
        return !cubic.Control2.HasXReference && cubic.Control2.X is >= -int.MaxValue and <= int.MaxValue;
    }

    private static bool TryLiteralCustomGeometryPathCubicControl2Y(
        PresentationCustomGeometryCubicBezier cubic,
        out long value)
    {
        value = cubic.Control2.Y;
        return !cubic.Control2.HasYReference && cubic.Control2.Y is >= -int.MaxValue and <= int.MaxValue;
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string ToStringInvariant(this long value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this uint value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private static string ToStringInvariant(this int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
