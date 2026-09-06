using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the deliberately bounded, source-preserving p:pic projection. The
// semantic image owns a compact canonical picture profile: asset, frame, crop,
// accessibility, opacity, preset or bounded custom mask, border, glow, inner shadow, reflection, soft edge, and one outer shadow. SVG
// fallback and bounded presentation extensions remain source-preserved.
internal static class PptxPictureCodec
{
    private const int MaxTextLength = 1_024;
    // A source-bound image may legitimately bleed past the slide edge. Keep
    // that valid DrawingML geometry bounded so malformed imports cannot turn
    // into unbounded JS coordinates or allocations.
    private const long MaxFrameCoordinateEmu = 100_000_000L;
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Office2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string SvgBlipNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private const string UseLocalDpiUri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private const string HiddenFillUri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}";

    internal static bool TryRead(P.Picture source, PptxPartContext context, out PresentationImage image)
    {
        image = new PresentationImage();
        var sourceProperties = source.ShapeProperties;
        var shadowSupported = PptxShadowCodec.TryRead(sourceProperties, out var shadow);
        var glowSupported = PptxGlowCodec.TryRead(sourceProperties, out var glow);
        var innerShadowSupported = PptxInnerShadowCodec.TryRead(sourceProperties, out var innerShadow);
        var reflectionSupported = PptxReflectionCodec.TryRead(sourceProperties, out var reflection);
        var softEdgeSupported = PptxSoftEdgeCodec.TryRead(sourceProperties, out var softEdge);
        if (!TryParts(source, out var nonVisual, out var blip, out var properties, out var transform, out var geometry, out var crop, out var tiled) ||
            !TryReadBorder(properties.GetFirstChild<A.Outline>(), out var border) ||
            !shadowSupported && !glowSupported && !innerShadowSupported && !reflectionSupported && !softEdgeSupported) return false;
        // Accessibility is a leaf capability, not ownership of the whole
        // picture. An ambiguous known extension hides the modeled metadata and
        // disables only that setter; unrelated picture edits remain residual-
        // preserving and source-bound.
        _ = PptxNonVisualAccessibilityCodec.TryReadResidual(nonVisual, out var accessibility);
        var relationshipId = blip.Embed?.Value ?? string.Empty;
        if (relationshipId.Length == 0) return false;
        try
        {
            var asset = context.ReadEmbeddedPicture(relationshipId);
            var svgRelationshipId = SvgFallbackRelationshipId(blip);
            var svgAsset = svgRelationshipId is null ? null : context.ReadEmbeddedPicture(svgRelationshipId);
            if (svgAsset is not null && !svgAsset.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                return false;
            var offset = transform.Offset!;
            var extents = transform.Extents!;
            image = new PresentationImage
            {
                AssetId = asset.Id,
                SvgAssetId = svgAsset?.Id ?? string.Empty,
                AltText = accessibility?.HasDescription == true ? accessibility.Description : string.Empty,
                AccessibilityTitle = accessibility?.HasTitle == true ? accessibility.Title : string.Empty,
                LeftEmu = offset.X?.Value ?? 0,
                TopEmu = offset.Y?.Value ?? 0,
                WidthEmu = extents.Cx?.Value ?? 0,
                HeightEmu = extents.Cy?.Value ?? 0,
            };
            if (accessibility?.HasDecorative == true) image.AccessibilityDecorative = accessibility.Decorative;
            if (crop is not null) image.Crop = ReadCrop(crop);
            image.Tiled = tiled;
            if (properties.GetFirstChild<A.Shape3DType>() is { } shape3d &&
                PptxShape3DCodec.TryReadDepth(shape3d, out var shape3dDepth))
                image.Shape3DDepthEmu = shape3dDepth;
            if (properties.GetFirstChild<A.Shape3DType>() is { } extrusionShape3d &&
                PptxShape3DCodec.TryReadExtrusionHeight(extrusionShape3d, out var shape3dExtrusionHeight))
                image.Shape3DExtrusionHeightEmu = shape3dExtrusionHeight;
            if (properties.GetFirstChild<A.Shape3DType>() is { } contourShape3d &&
                PptxShape3DCodec.TryReadContourWidth(contourShape3d, out var shape3dContourWidth))
                image.Shape3DContourWidthEmu = shape3dContourWidth;
            if (properties.GetFirstChild<A.Shape3DType>() is { } materialShape3d &&
                PptxShape3DCodec.TryReadPresetMaterial(materialShape3d, out var shape3dPresetMaterial))
                image.Shape3DPresetMaterial = shape3dPresetMaterial;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelTopShape3d &&
                PptxShape3DCodec.TryReadBevelTopWidth(bevelTopShape3d, out var shape3dBevelTopWidth))
                image.Shape3DBevelTopWidthEmu = shape3dBevelTopWidth;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelTopHeightShape3d &&
                PptxShape3DCodec.TryReadBevelTopHeight(bevelTopHeightShape3d, out var shape3dBevelTopHeight))
                image.Shape3DBevelTopHeightEmu = shape3dBevelTopHeight;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelTopPresetShape3d &&
                PptxShape3DCodec.TryReadBevelTopPreset(bevelTopPresetShape3d, out var shape3dBevelTopPreset))
                image.Shape3DBevelTopPreset = shape3dBevelTopPreset;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelBottomShape3d &&
                PptxShape3DCodec.TryReadBevelBottomWidth(bevelBottomShape3d, out var shape3dBevelBottomWidth))
                image.Shape3DBevelBottomWidthEmu = shape3dBevelBottomWidth;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelBottomHeightShape3d &&
                PptxShape3DCodec.TryReadBevelBottomHeight(bevelBottomHeightShape3d, out var shape3dBevelBottomHeight))
                image.Shape3DBevelBottomHeightEmu = shape3dBevelBottomHeight;
            if (properties.GetFirstChild<A.Shape3DType>() is { } bevelBottomPresetShape3d &&
                PptxShape3DCodec.TryReadBevelBottomPreset(bevelBottomPresetShape3d, out var shape3dBevelBottomPreset))
                image.Shape3DBevelBottomPreset = shape3dBevelBottomPreset;
            if (properties.GetFirstChild<A.Shape3DType>() is { } contourRgbShape3d &&
                PptxShape3DCodec.TryReadContourRgb(contourRgbShape3d, out var shape3dContourRgb))
                image.Shape3DContourRgb = shape3dContourRgb;
            if (properties.GetFirstChild<A.Shape3DType>() is { } extrusionRgbShape3d &&
                PptxShape3DCodec.TryReadExtrusionRgb(extrusionRgbShape3d, out var shape3dExtrusionRgb))
                image.Shape3DExtrusionRgb = shape3dExtrusionRgb;
            if (properties.GetFirstChild<A.Shape3DType>() is { } contourColorSchemeShape3d &&
                PptxShape3DCodec.TryReadContourColorScheme(contourColorSchemeShape3d, out var shape3dContourColorScheme))
                image.Shape3DContourColorScheme = shape3dContourColorScheme;
            if (properties.GetFirstChild<A.Shape3DType>() is { } extrusionColorSchemeShape3d &&
                PptxShape3DCodec.TryReadExtrusionColorScheme(extrusionColorSchemeShape3d, out var shape3dExtrusionColorScheme))
                image.Shape3DExtrusionColorScheme = shape3dExtrusionColorScheme;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } scene3d &&
                PptxShape3DCodec.TryReadSceneCameraPreset(scene3d, out var sceneCameraPreset))
                image.Shape3DSceneCameraPreset = sceneCameraPreset;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } cameraZoomScene3d &&
                PptxShape3DCodec.TryReadSceneCameraZoom(cameraZoomScene3d, out var sceneCameraZoom))
                image.Shape3DSceneCameraZoomThousandthPercent = sceneCameraZoom;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } cameraFovScene3d &&
                PptxShape3DCodec.TryReadSceneCameraFov(cameraFovScene3d, out var sceneCameraFov))
                image.Shape3DSceneCameraFov60000 = sceneCameraFov;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationLatitudeScene3d &&
                PptxShape3DCodec.TryReadSceneCameraRotationLatitude(cameraRotationLatitudeScene3d, out var sceneCameraRotationLatitude))
                image.Shape3DSceneCameraRotationLatitude60000 = sceneCameraRotationLatitude;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationLongitudeScene3d &&
                PptxShape3DCodec.TryReadSceneCameraRotationLongitude(cameraRotationLongitudeScene3d, out var sceneCameraRotationLongitude))
                image.Shape3DSceneCameraRotationLongitude60000 = sceneCameraRotationLongitude;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationRevolutionScene3d &&
                PptxShape3DCodec.TryReadSceneCameraRotationRevolution(cameraRotationRevolutionScene3d, out var sceneCameraRotationRevolution))
                image.Shape3DSceneCameraRotationRevolution60000 = sceneCameraRotationRevolution;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } lightRigPresetScene3d &&
                PptxShape3DCodec.TryReadSceneLightRigPreset(lightRigPresetScene3d, out var sceneLightRigPreset))
                image.Shape3DSceneLightRigPreset = sceneLightRigPreset;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } lightRigDirectionScene3d &&
                PptxShape3DCodec.TryReadSceneLightRigDirection(lightRigDirectionScene3d, out var sceneLightRigDirection))
                image.Shape3DSceneLightRigDirection = sceneLightRigDirection;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationLatitudeScene3d &&
                PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(lightRigRotationLatitudeScene3d, out var sceneLightRigRotationLatitude))
                image.Shape3DSceneLightRigRotationLatitude60000 = sceneLightRigRotationLatitude;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationLongitudeScene3d &&
                PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(lightRigRotationLongitudeScene3d, out var sceneLightRigRotationLongitude))
                image.Shape3DSceneLightRigRotationLongitude60000 = sceneLightRigRotationLongitude;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationRevolutionScene3d &&
                PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(lightRigRotationRevolutionScene3d, out var sceneLightRigRotationRevolution))
                image.Shape3DSceneLightRigRotationRevolution60000 = sceneLightRigRotationRevolution;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorXScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropAnchorX(backdropAnchorXScene3d, out var sceneBackdropAnchorX))
                image.Shape3DSceneBackdropAnchorXEmu = sceneBackdropAnchorX;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorYScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropAnchorY(backdropAnchorYScene3d, out var sceneBackdropAnchorY))
                image.Shape3DSceneBackdropAnchorYEmu = sceneBackdropAnchorY;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorZScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropAnchorZ(backdropAnchorZScene3d, out var sceneBackdropAnchorZ))
                image.Shape3DSceneBackdropAnchorZEmu = sceneBackdropAnchorZ;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDxScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropNormalDx(backdropNormalDxScene3d, out var sceneBackdropNormalDx))
                image.Shape3DSceneBackdropNormalDxEmu = sceneBackdropNormalDx;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDyScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropNormalDy(backdropNormalDyScene3d, out var sceneBackdropNormalDy))
                image.Shape3DSceneBackdropNormalDyEmu = sceneBackdropNormalDy;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDzScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropNormalDz(backdropNormalDzScene3d, out var sceneBackdropNormalDz))
                image.Shape3DSceneBackdropNormalDzEmu = sceneBackdropNormalDz;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDxScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropUpDx(backdropUpDxScene3d, out var sceneBackdropUpDx))
                image.Shape3DSceneBackdropUpDxEmu = sceneBackdropUpDx;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDyScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropUpDy(backdropUpDyScene3d, out var sceneBackdropUpDy))
                image.Shape3DSceneBackdropUpDyEmu = sceneBackdropUpDy;
            if (properties.Elements<A.Scene3DType>().Count() == 1 &&
                properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDzScene3d &&
                PptxShape3DCodec.TryReadSceneBackdropUpDz(backdropUpDzScene3d, out var sceneBackdropUpDz))
                image.Shape3DSceneBackdropUpDzEmu = sceneBackdropUpDz;
            // An empty alphaModFix is a valid no-op effect, not an explicit
            // opacity token. Preserve it as opaque source structure rather
            // than manufacturing an editable 100% leaf that the writer
            // cannot later prove. Only an authored `amt` is a controlled
            // source-bound opacity value.
            if (blip.GetFirstChild<A.AlphaModulationFixed>() is { Amount.Value: { } amount })
                image.OpacityThousandthPercent = checked((uint)amount);
            if (!TryReadMask(geometry, image.WidthEmu, image.HeightEmu, image)) return false;
            image.Border = border;
            image.Shadow = shadow;
            image.Glow = glow;
            image.InnerShadow = innerShadow;
            image.Reflection = reflection;
            image.SoftEdge = softEdge;
            var visual = ReadTransform(transform);
            if (visual is not null) image.Transform = visual;
            PptxNonVisualAccessibilityCodec.Validate(Accessibility(image), "source", "image");
            return FrameCoordinateSupported(image.LeftEmu, allowNegative: true) &&
                   FrameCoordinateSupported(image.TopEmu, allowNegative: true) &&
                   FrameExtentSupported(image.WidthEmu) && FrameExtentSupported(image.HeightEmu) &&
                   (nonVisual.Name?.Value?.Length ?? 0) <= MaxTextLength;
        }
        catch (CodecException)
        {
            image = new PresentationImage();
            return false;
        }
    }

    internal static void Validate(PresentationImage? image, string elementId, PptxAssetCatalog assets, bool sourceBound = false)
    {
        if (image is null)
            throw Invalid(elementId, "payload is missing");
        if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512)
            throw Invalid(elementId, "asset ID must contain 1 through 512 characters");
        _ = assets.Get(image.AssetId);
        if (image.SvgAssetId.Length > 512)
            throw Invalid(elementId, "SVG fallback asset ID must contain at most 512 characters");
        if (image.SvgAssetId.Length > 0 && !assets.Get(image.SvgAssetId).ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
            throw Invalid(elementId, "SVG fallback asset must use content type image/svg+xml");
        PptxNonVisualAccessibilityCodec.Validate(Accessibility(image), elementId, "image");
        if (!FrameCoordinateSupported(image.LeftEmu, sourceBound) || !FrameCoordinateSupported(image.TopEmu, sourceBound) ||
            !FrameExtentSupported(image.WidthEmu) || !FrameExtentSupported(image.HeightEmu))
            throw Invalid(elementId, sourceBound
                ? "source-bound frame coordinates must stay within the bounded DrawingML range and extents must be positive"
                : "frame must use non-negative coordinates and positive extents");
        if (image.Crop is not null && !CropValuesValid(image.Crop))
            throw Invalid(elementId, "crop edges must be between -100% and 100% and opposing sums must remain below 100%");
        if (image.HasOpacityThousandthPercent && image.OpacityThousandthPercent > 100_000)
            throw Invalid(elementId, "opacity must be between 0% and 100%");
        if (image.HasShape3DDepthEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D depth is source-bound only");
            if (image.Shape3DDepthEmu < int.MinValue || image.Shape3DDepthEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D depth must fit the signed 32-bit DrawingML range");
        }
        if (image.HasShape3DExtrusionHeightEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D extrusion height is source-bound only");
            if (image.Shape3DExtrusionHeightEmu < 0 || image.Shape3DExtrusionHeightEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D extrusion height must fit the non-negative signed 32-bit DrawingML range");
        }
        if (image.HasShape3DContourWidthEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D contour width is source-bound only");
            if (image.Shape3DContourWidthEmu < 0 || image.Shape3DContourWidthEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D contour width must fit the non-negative signed 32-bit DrawingML range");
        }
        if (!string.IsNullOrEmpty(image.Shape3DPresetMaterial))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D preset material is source-bound only");
            if (!PptxShape3DCodec.IsPresetMaterialToken(image.Shape3DPresetMaterial))
                throw Invalid(elementId, "picture 3-D preset material must use a bounded DrawingML token");
        }
        if (image.HasShape3DBevelTopWidthEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D top-bevel width is source-bound only");
            if (image.Shape3DBevelTopWidthEmu < 0 || image.Shape3DBevelTopWidthEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D top-bevel width must fit the non-negative signed 32-bit DrawingML range");
        }
        if (image.HasShape3DBevelTopHeightEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D top-bevel height is source-bound only");
            if (image.Shape3DBevelTopHeightEmu < 0 || image.Shape3DBevelTopHeightEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D top-bevel height must fit the non-negative signed 32-bit DrawingML range");
        }
        if (!string.IsNullOrEmpty(image.Shape3DBevelTopPreset))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D top-bevel preset is source-bound only");
            if (!PptxShape3DCodec.IsBevelTopPresetToken(image.Shape3DBevelTopPreset))
                throw Invalid(elementId, "picture 3-D top-bevel preset must use a bounded DrawingML token");
        }
        if (image.HasShape3DBevelBottomWidthEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D bottom-bevel width is source-bound only");
            if (image.Shape3DBevelBottomWidthEmu < 0 || image.Shape3DBevelBottomWidthEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D bottom-bevel width must fit the non-negative signed 32-bit DrawingML range");
        }
        if (image.HasShape3DBevelBottomHeightEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D bottom-bevel height is source-bound only");
            if (image.Shape3DBevelBottomHeightEmu < 0 || image.Shape3DBevelBottomHeightEmu > int.MaxValue)
                throw Invalid(elementId, "picture 3-D bottom-bevel height must fit the non-negative signed 32-bit DrawingML range");
        }
        if (!string.IsNullOrEmpty(image.Shape3DBevelBottomPreset))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D bottom-bevel preset is source-bound only");
            if (!PptxShape3DCodec.IsBevelBottomPresetToken(image.Shape3DBevelBottomPreset))
                throw Invalid(elementId, "picture 3-D bottom-bevel preset must use a bounded DrawingML token");
        }
        if (!string.IsNullOrEmpty(image.Shape3DContourRgb))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D contour color is source-bound only");
            _ = PptxColor.Normalize(image.Shape3DContourRgb);
        }
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionRgb))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D extrusion color is source-bound only");
            _ = PptxColor.Normalize(image.Shape3DExtrusionRgb);
        }
        if (!string.IsNullOrEmpty(image.Shape3DContourColorScheme))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D contour theme color is source-bound only");
            _ = PptxColor.NormalizeScheme(image.Shape3DContourColorScheme);
        }
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionColorScheme))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D extrusion theme color is source-bound only");
            _ = PptxColor.NormalizeScheme(image.Shape3DExtrusionColorScheme);
        }
        if (!string.IsNullOrEmpty(image.Shape3DSceneCameraPreset))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera preset is source-bound only");
            if (!PptxShape3DCodec.IsSceneCameraPresetToken(image.Shape3DSceneCameraPreset))
                throw Invalid(elementId, "picture 3-D scene camera preset must use a bounded DrawingML token");
        }
        if (image.HasShape3DSceneCameraZoomThousandthPercent)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera zoom is source-bound only");
            if (image.Shape3DSceneCameraZoomThousandthPercent < 0 ||
                image.Shape3DSceneCameraZoomThousandthPercent > int.MaxValue)
                throw Invalid(elementId, "picture 3-D scene camera zoom must fit the non-negative signed 32-bit DrawingML range");
        }
        if (image.HasShape3DSceneCameraFov60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera FOV is source-bound only");
            if (image.Shape3DSceneCameraFov60000 <= 0 || image.Shape3DSceneCameraFov60000 >= 180L * 60_000)
                throw Invalid(elementId, "picture 3-D scene camera FOV must be positive and below 180 degrees");
        }
        if (image.HasShape3DSceneCameraRotationLatitude60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera rotation latitude is source-bound only");
            if (image.Shape3DSceneCameraRotationLatitude60000 < 0 ||
                image.Shape3DSceneCameraRotationLatitude60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene camera rotation latitude must be non-negative and at most 360 degrees");
        }
        if (image.HasShape3DSceneCameraRotationLongitude60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera rotation longitude is source-bound only");
            if (image.Shape3DSceneCameraRotationLongitude60000 < 0 ||
                image.Shape3DSceneCameraRotationLongitude60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene camera rotation longitude must be non-negative and at most 360 degrees");
        }
        if (image.HasShape3DSceneCameraRotationRevolution60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene camera rotation revolution is source-bound only");
            if (image.Shape3DSceneCameraRotationRevolution60000 < 0 ||
                image.Shape3DSceneCameraRotationRevolution60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene camera rotation revolution must be non-negative and at most 360 degrees");
        }
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigPreset))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene light-rig preset is source-bound only");
            if (!PptxShape3DCodec.IsSceneLightRigPresetToken(image.Shape3DSceneLightRigPreset))
                throw Invalid(elementId, "picture 3-D scene light-rig preset must use a bounded DrawingML token");
        }
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigDirection))
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene light-rig direction is source-bound only");
            if (!PptxShape3DCodec.IsSceneLightRigDirectionToken(image.Shape3DSceneLightRigDirection))
                throw Invalid(elementId, "picture 3-D scene light-rig direction must use a bounded DrawingML token");
        }
        if (image.HasShape3DSceneLightRigRotationLatitude60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation latitude is source-bound only");
            if (image.Shape3DSceneLightRigRotationLatitude60000 < 0 ||
                image.Shape3DSceneLightRigRotationLatitude60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation latitude must be non-negative and at most 360 degrees");
        }
        if (image.HasShape3DSceneLightRigRotationLongitude60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation longitude is source-bound only");
            if (image.Shape3DSceneLightRigRotationLongitude60000 < 0 ||
                image.Shape3DSceneLightRigRotationLongitude60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation longitude must be non-negative and at most 360 degrees");
        }
        if (image.HasShape3DSceneLightRigRotationRevolution60000)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation revolution is source-bound only");
            if (image.Shape3DSceneLightRigRotationRevolution60000 < 0 ||
                image.Shape3DSceneLightRigRotationRevolution60000 > 360L * 60_000)
                throw Invalid(elementId, "picture 3-D scene light-rig rotation revolution must be non-negative and at most 360 degrees");
        }
        if (image.HasShape3DSceneBackdropAnchorXEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop anchor X is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropAnchorXEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop anchor X must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropAnchorYEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop anchor Y is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropAnchorYEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop anchor Y must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropAnchorZEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop anchor Z is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropAnchorZEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop anchor Z must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropNormalDxEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop normal X is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropNormalDxEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop normal X must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropNormalDyEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop normal Y is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropNormalDyEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop normal Y must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropNormalDzEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop normal Z is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropNormalDzEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop normal Z must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropUpDxEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector X is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropUpDxEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector X must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropUpDyEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector Y is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropUpDyEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector Y must be a bounded signed coordinate");
        }
        if (image.HasShape3DSceneBackdropUpDzEmu)
        {
            if (!sourceBound)
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector Z is source-bound only");
            if (!PptxShape3DCodec.TryLiteralBackdropCoordinate(
                    image.Shape3DSceneBackdropUpDzEmu.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    out _))
                throw Invalid(elementId, "picture 3-D scene backdrop up-vector Z must be a bounded signed coordinate");
        }
        if (image.CustomMaskPaths.Count > 0)
        {
            if (image.MaskPreset.Length > 0 || image.MaskPresetAdjustments.Count > 0)
                throw Invalid(elementId, "custom image mask paths cannot be combined with preset mask state");
            PptxCustomGeometryCodec.Validate(CustomMaskShape(image), elementId + " image mask");
        }
        else
        {
            var maskPreset = image.MaskPreset.Length == 0 ? "rect" : image.MaskPreset;
            PptxPresetGeometryAdjustmentCodec.Validate(maskPreset, image.MaskPresetAdjustments, elementId + " image mask");
        }
        ValidateBorder(image.Border, elementId);
        PptxShadowCodec.Validate(image.Shadow, elementId, "image");
        PptxGlowCodec.Validate(image.Glow, elementId, "image");
        PptxInnerShadowCodec.Validate(image.InnerShadow, elementId, "image");
        PptxReflectionCodec.Validate(image.Reflection, elementId, "image");
        PptxSoftEdgeCodec.Validate(image.SoftEdge, elementId, "image");
        ValidateTransform(image.Transform, elementId);
    }

    internal static P.Picture Build(PresentationElement source, uint nativeId, PptxPartContext context)
    {
        var image = source.Image;
        // SVG fallbacks are currently source-bound only. Refuse an authored
        // image carrying one instead of silently dropping the fallback while
        // constructing a new picture.
        if (image.SvgAssetId.Length > 0)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author an SVG fallback outside the bounded source-preserving path.");
        if (image.HasShape3DDepthEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D depth in a new picture.");
        if (image.HasShape3DExtrusionHeightEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D extrusion height in a new picture.");
        if (image.HasShape3DContourWidthEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D contour width in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DPresetMaterial))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D preset material in a new picture.");
        if (image.HasShape3DBevelTopWidthEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D top-bevel width in a new picture.");
        if (image.HasShape3DBevelTopHeightEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D top-bevel height in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DBevelTopPreset))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D top-bevel preset in a new picture.");
        if (image.HasShape3DBevelBottomWidthEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D bottom-bevel width in a new picture.");
        if (image.HasShape3DBevelBottomHeightEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D bottom-bevel height in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DBevelBottomPreset))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D bottom-bevel preset in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DContourRgb))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D contour color in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionRgb))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D extrusion color in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DContourColorScheme))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D contour theme color in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DExtrusionColorScheme))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D extrusion theme color in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DSceneCameraPreset))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera preset in a new picture.");
        if (image.HasShape3DSceneCameraZoomThousandthPercent)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera zoom in a new picture.");
        if (image.HasShape3DSceneCameraFov60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera FOV in a new picture.");
        if (image.HasShape3DSceneCameraRotationLatitude60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera rotation latitude in a new picture.");
        if (image.HasShape3DSceneCameraRotationLongitude60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera rotation longitude in a new picture.");
        if (image.HasShape3DSceneCameraRotationRevolution60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene camera rotation revolution in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigPreset))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene light-rig preset in a new picture.");
        if (!string.IsNullOrEmpty(image.Shape3DSceneLightRigDirection))
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene light-rig direction in a new picture.");
        if (image.HasShape3DSceneLightRigRotationLatitude60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene light-rig rotation latitude in a new picture.");
        if (image.HasShape3DSceneLightRigRotationLongitude60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene light-rig rotation longitude in a new picture.");
        if (image.HasShape3DSceneLightRigRotationRevolution60000)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene light-rig rotation revolution in a new picture.");
        if (image.HasShape3DSceneBackdropAnchorXEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop anchor X in a new picture.");
        if (image.HasShape3DSceneBackdropAnchorYEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop anchor Y in a new picture.");
        if (image.HasShape3DSceneBackdropAnchorZEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop anchor Z in a new picture.");
        if (image.HasShape3DSceneBackdropNormalDxEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop normal X in a new picture.");
        if (image.HasShape3DSceneBackdropNormalDyEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop normal Y in a new picture.");
        if (image.HasShape3DSceneBackdropNormalDzEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop normal Z in a new picture.");
        if (image.HasShape3DSceneBackdropUpDxEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop up-vector X in a new picture.");
        if (image.HasShape3DSceneBackdropUpDyEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop up-vector Y in a new picture.");
        if (image.HasShape3DSceneBackdropUpDzEmu)
            throw new CodecException(
                "unsupported_presentation_image",
                $"Presentation image {source.Id} cannot author source-bound 3-D scene backdrop up-vector Z in a new picture.");
        var transform = new A.Transform2D(
            new A.Offset { X = image.LeftEmu, Y = image.TopEmu },
            new A.Extents { Cx = image.WidthEmu, Cy = image.HeightEmu });
        ApplyTransform(transform, image.Transform);
        var blip = new A.Blip { Embed = context.AddEmbeddedPicture(image.AssetId) };
        ApplyOpacity(blip, image.HasOpacityThousandthPercent ? image.OpacityThousandthPercent : null);
        var fill = new P.BlipFill(blip);
        if (image.Crop is not null) fill.Append(BuildCrop(image.Crop));
        fill.Append(image.Tiled ? new A.Tile() : new A.Stretch(new A.FillRectangle()));
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, Accessibility(image));
        var properties = new P.ShapeProperties(transform);
        if (image.CustomMaskPaths.Count > 0)
            PptxCustomGeometryCodec.Apply(properties, CustomMaskShape(image), source.Id + " image mask");
        else
            properties.Append(BuildMask(image.MaskPreset, image.MaskPresetAdjustments, source.Id));
        if (image.Border is not null) properties.Append(BuildBorder(image.Border));
        PptxShadowCodec.Apply(properties, image.Shadow);
        PptxGlowCodec.Apply(properties, image.Glow);
        PptxInnerShadowCodec.Apply(properties, image.InnerShadow);
        PptxReflectionCodec.Apply(properties, image.Reflection);
        PptxSoftEdgeCodec.Apply(properties, image.SoftEdge);
        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisual,
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            fill,
            properties);
    }

    internal static void Apply(P.Picture source, PresentationElement requested, PptxPartContext context)
    {
        if (!TryRead(source, context, out var currentImage) ||
            !TryParts(source, out var nonVisual, out var blip, out var properties, out var transform, out var geometry, out _, out _))
            throw new CodecException("unsupported_presentation_edit", $"Presentation image {requested.Id} no longer matches the editable picture profile.");
        var current = context.ReadEmbeddedPicture(blip.Embed?.Value ?? string.Empty);
        var replacement = context.Assets?.Get(requested.Image.AssetId) ??
            throw new CodecException("invalid_presentation_asset", $"Presentation image {requested.Id} requires an asset catalog.");
        if (!current.Id.Equals(replacement.Id, StringComparison.Ordinal))
        {
            if (!current.ContentType.Equals(replacement.ContentType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} replacement must retain content type {current.ContentType}.");
            if (HasSvgFallback(blip) && requested.Image.SvgAssetId.Length == 0)
                throw new CodecException(
                    "unsupported_presentation_image",
                    $"Presentation image {requested.Id} has a paired SVG fallback; replacing only the primary asset would leave the fallback stale.");
            blip.Embed = context.AddEmbeddedPicture(replacement.Id);
        }
        var currentSvgRelationshipId = SvgFallbackRelationshipId(blip);
        if (currentSvgRelationshipId is null)
        {
            if (requested.Image.SvgAssetId.Length > 0)
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} has no source SVG fallback relationship to replace.");
        }
        else
        {
            var currentSvg = context.ReadEmbeddedPicture(currentSvgRelationshipId);
            var requestedSvg = requested.Image.SvgAssetId.Length == 0
                ? throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} cannot remove a paired SVG fallback in the bounded source-preserving path.")
                : context.Assets.Get(requested.Image.SvgAssetId);
            if (!currentSvg.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) ||
                !requestedSvg.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} SVG fallback must retain content type image/svg+xml.");
            if (!currentSvg.Id.Equals(requestedSvg.Id, StringComparison.Ordinal))
            {
                var fallback = SvgFallbackElement(blip) ??
                    throw new CodecException("unsupported_presentation_image", $"Presentation image {requested.Id} SVG fallback relationship is malformed.");
                fallback.SetAttribute(new OpenXmlAttribute("r", "embed", OfficeRelationshipsNamespace, context.AddEmbeddedPicture(requestedSvg.Id)));
            }
        }
        nonVisual.Name = requested.Name;
        if (!AccessibilityEqual(currentImage, requested.Image))
            PptxNonVisualAccessibilityCodec.ApplyResidualBound(nonVisual, Accessibility(requested.Image), "image");
        transform.Offset!.X = requested.Image.LeftEmu;
        transform.Offset.Y = requested.Image.TopEmu;
        transform.Extents!.Cx = requested.Image.WidthEmu;
        transform.Extents.Cy = requested.Image.HeightEmu;
        ApplyCrop(source.BlipFill!, requested.Image.Crop);
        if (currentImage.Tiled != requested.Image.Tiled)
        {
            source.BlipFill!.ChildElements.Last().Remove();
            source.BlipFill.Append(requested.Image.Tiled ? new A.Tile() : new A.Stretch(new A.FillRectangle()));
        }
        ApplyTransform(transform, requested.Image.Transform);
        if (currentImage.HasShape3DDepthEmu != requested.Image.HasShape3DDepthEmu ||
            currentImage.HasShape3DDepthEmu && currentImage.Shape3DDepthEmu != requested.Image.Shape3DDepthEmu)
        {
            if (!currentImage.HasShape3DDepthEmu || !requested.Image.HasShape3DDepthEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } shape3d ||
                !PptxShape3DCodec.TryReadDepth(shape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D depth is outside the bounded source-preserving profile.");
            shape3d.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "z",
                string.Empty,
                requested.Image.Shape3DDepthEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DExtrusionHeightEmu != requested.Image.HasShape3DExtrusionHeightEmu ||
            currentImage.HasShape3DExtrusionHeightEmu && currentImage.Shape3DExtrusionHeightEmu != requested.Image.Shape3DExtrusionHeightEmu)
        {
            if (!currentImage.HasShape3DExtrusionHeightEmu || !requested.Image.HasShape3DExtrusionHeightEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } extrusionShape3d ||
                !PptxShape3DCodec.TryReadExtrusionHeight(extrusionShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D extrusion height is outside the bounded source-preserving profile.");
            extrusionShape3d.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "extrusionH",
                string.Empty,
                requested.Image.Shape3DExtrusionHeightEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DContourWidthEmu != requested.Image.HasShape3DContourWidthEmu ||
            currentImage.HasShape3DContourWidthEmu && currentImage.Shape3DContourWidthEmu != requested.Image.Shape3DContourWidthEmu)
        {
            if (!currentImage.HasShape3DContourWidthEmu || !requested.Image.HasShape3DContourWidthEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } contourShape3d ||
                !PptxShape3DCodec.TryReadContourWidth(contourShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D contour width is outside the bounded source-preserving profile.");
            contourShape3d.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "contourW",
                string.Empty,
                requested.Image.Shape3DContourWidthEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (!string.Equals(currentImage.Shape3DPresetMaterial, requested.Image.Shape3DPresetMaterial, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DPresetMaterial) ||
                string.IsNullOrEmpty(requested.Image.Shape3DPresetMaterial) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } materialShape3d ||
                !PptxShape3DCodec.TryReadPresetMaterial(materialShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D preset material is outside the bounded source-preserving profile.");
            materialShape3d.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "prstMaterial",
                string.Empty,
                requested.Image.Shape3DPresetMaterial));
        }
        if (currentImage.HasShape3DBevelTopWidthEmu != requested.Image.HasShape3DBevelTopWidthEmu ||
            currentImage.HasShape3DBevelTopWidthEmu && currentImage.Shape3DBevelTopWidthEmu != requested.Image.Shape3DBevelTopWidthEmu)
        {
            if (!currentImage.HasShape3DBevelTopWidthEmu || !requested.Image.HasShape3DBevelTopWidthEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelTopShape3d ||
                !PptxShape3DCodec.TryReadBevelTopWidth(bevelTopShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel width is outside the bounded source-preserving profile.");
            var bevelTop = bevelTopShape3d.GetFirstChild<A.BevelTop>();
            if (bevelTop is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel owner is outside the bounded source-preserving profile.");
            bevelTop.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "w",
                string.Empty,
                requested.Image.Shape3DBevelTopWidthEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DBevelTopHeightEmu != requested.Image.HasShape3DBevelTopHeightEmu ||
            currentImage.HasShape3DBevelTopHeightEmu && currentImage.Shape3DBevelTopHeightEmu != requested.Image.Shape3DBevelTopHeightEmu)
        {
            if (!currentImage.HasShape3DBevelTopHeightEmu || !requested.Image.HasShape3DBevelTopHeightEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelTopHeightShape3d ||
                !PptxShape3DCodec.TryReadBevelTopHeight(bevelTopHeightShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel height is outside the bounded source-preserving profile.");
            var bevelTop = bevelTopHeightShape3d.GetFirstChild<A.BevelTop>();
            if (bevelTop is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel owner is outside the bounded source-preserving profile.");
            bevelTop.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "h",
                string.Empty,
                requested.Image.Shape3DBevelTopHeightEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (!string.Equals(currentImage.Shape3DBevelTopPreset, requested.Image.Shape3DBevelTopPreset, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DBevelTopPreset) ||
                string.IsNullOrEmpty(requested.Image.Shape3DBevelTopPreset) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelTopPresetShape3d ||
                !PptxShape3DCodec.TryReadBevelTopPreset(bevelTopPresetShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel preset is outside the bounded source-preserving profile.");
            var bevelTop = bevelTopPresetShape3d.GetFirstChild<A.BevelTop>();
            if (bevelTop is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D top-bevel owner is outside the bounded source-preserving profile.");
            bevelTop.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "prst",
                string.Empty,
                requested.Image.Shape3DBevelTopPreset));
        }
        if (currentImage.HasShape3DBevelBottomWidthEmu != requested.Image.HasShape3DBevelBottomWidthEmu ||
            currentImage.HasShape3DBevelBottomWidthEmu && currentImage.Shape3DBevelBottomWidthEmu != requested.Image.Shape3DBevelBottomWidthEmu)
        {
            if (!currentImage.HasShape3DBevelBottomWidthEmu || !requested.Image.HasShape3DBevelBottomWidthEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelBottomShape3d ||
                !PptxShape3DCodec.TryReadBevelBottomWidth(bevelBottomShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel width is outside the bounded source-preserving profile.");
            var bevelBottom = bevelBottomShape3d.GetFirstChild<A.BevelBottom>();
            if (bevelBottom is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel owner is outside the bounded source-preserving profile.");
            bevelBottom.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "w",
                string.Empty,
                requested.Image.Shape3DBevelBottomWidthEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DBevelBottomHeightEmu != requested.Image.HasShape3DBevelBottomHeightEmu ||
            currentImage.HasShape3DBevelBottomHeightEmu && currentImage.Shape3DBevelBottomHeightEmu != requested.Image.Shape3DBevelBottomHeightEmu)
        {
            if (!currentImage.HasShape3DBevelBottomHeightEmu || !requested.Image.HasShape3DBevelBottomHeightEmu ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelBottomHeightShape3d ||
                !PptxShape3DCodec.TryReadBevelBottomHeight(bevelBottomHeightShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel height is outside the bounded source-preserving profile.");
            var bevelBottom = bevelBottomHeightShape3d.GetFirstChild<A.BevelBottom>();
            if (bevelBottom is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel owner is outside the bounded source-preserving profile.");
            bevelBottom.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "h",
                string.Empty,
                requested.Image.Shape3DBevelBottomHeightEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (!string.Equals(currentImage.Shape3DBevelBottomPreset, requested.Image.Shape3DBevelBottomPreset, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DBevelBottomPreset) ||
                string.IsNullOrEmpty(requested.Image.Shape3DBevelBottomPreset) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } bevelBottomPresetShape3d ||
                !PptxShape3DCodec.TryReadBevelBottomPreset(bevelBottomPresetShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel preset is outside the bounded source-preserving profile.");
            var bevelBottom = bevelBottomPresetShape3d.GetFirstChild<A.BevelBottom>();
            if (bevelBottom is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D bottom-bevel owner is outside the bounded source-preserving profile.");
            bevelBottom.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "prst",
                string.Empty,
                requested.Image.Shape3DBevelBottomPreset));
        }
        if (!string.Equals(currentImage.Shape3DContourRgb, requested.Image.Shape3DContourRgb, StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DContourRgb) ||
                string.IsNullOrEmpty(requested.Image.Shape3DContourRgb) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } contourRgbShape3d ||
                !PptxShape3DCodec.TryReadContourRgb(contourRgbShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D contour color is outside the bounded source-preserving profile.");
            var contourColor = contourRgbShape3d.GetFirstChild<A.ContourColor>();
            var contourRgb = contourColor?.GetFirstChild<A.RgbColorModelHex>();
            if (contourRgb is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D contour color owner is outside the bounded source-preserving profile.");
            contourRgb.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "val",
                string.Empty,
                PptxColor.Normalize(requested.Image.Shape3DContourRgb)));
        }
        if (!string.Equals(currentImage.Shape3DContourColorScheme, requested.Image.Shape3DContourColorScheme, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DContourColorScheme) ||
                string.IsNullOrEmpty(requested.Image.Shape3DContourColorScheme) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } contourColorSchemeShape3d ||
                !PptxShape3DCodec.TryReadContourColorScheme(contourColorSchemeShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D contour theme color is outside the bounded source-preserving profile.");
            var contourColor = contourColorSchemeShape3d.GetFirstChild<A.ContourColor>();
            var contourScheme = contourColor?.GetFirstChild<A.SchemeColor>();
            if (contourScheme is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D contour theme color owner is outside the bounded source-preserving profile.");
            contourScheme.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "val",
                string.Empty,
                PptxColor.NormalizeScheme(requested.Image.Shape3DContourColorScheme)));
        }
        if (!string.Equals(currentImage.Shape3DExtrusionColorScheme, requested.Image.Shape3DExtrusionColorScheme, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DExtrusionColorScheme) ||
                string.IsNullOrEmpty(requested.Image.Shape3DExtrusionColorScheme) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } extrusionColorSchemeShape3d ||
                !PptxShape3DCodec.TryReadExtrusionColorScheme(extrusionColorSchemeShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D extrusion theme color is outside the bounded source-preserving profile.");
            var extrusionColor = extrusionColorSchemeShape3d.GetFirstChild<A.ExtrusionColor>();
            var extrusionScheme = extrusionColor?.GetFirstChild<A.SchemeColor>();
            if (extrusionScheme is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D extrusion theme color owner is outside the bounded source-preserving profile.");
            extrusionScheme.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "val",
                string.Empty,
                PptxColor.NormalizeScheme(requested.Image.Shape3DExtrusionColorScheme)));
        }
        if (!string.Equals(currentImage.Shape3DExtrusionRgb, requested.Image.Shape3DExtrusionRgb, StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DExtrusionRgb) ||
                string.IsNullOrEmpty(requested.Image.Shape3DExtrusionRgb) ||
                properties.GetFirstChild<A.Shape3DType>() is not { } extrusionRgbShape3d ||
                !PptxShape3DCodec.TryReadExtrusionRgb(extrusionRgbShape3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D extrusion color is outside the bounded source-preserving profile.");
            var extrusionColor = extrusionRgbShape3d.GetFirstChild<A.ExtrusionColor>();
            var extrusionRgb = extrusionColor?.GetFirstChild<A.RgbColorModelHex>();
            if (extrusionRgb is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D extrusion color owner is outside the bounded source-preserving profile.");
            extrusionRgb.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "val",
                string.Empty,
                PptxColor.Normalize(requested.Image.Shape3DExtrusionRgb)));
        }
        if (!string.Equals(currentImage.Shape3DSceneCameraPreset, requested.Image.Shape3DSceneCameraPreset, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DSceneCameraPreset) ||
                string.IsNullOrEmpty(requested.Image.Shape3DSceneCameraPreset) ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraPreset(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera preset is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            if (camera is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera owner is outside the bounded source-preserving profile.");
            camera.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "prst",
                string.Empty,
                requested.Image.Shape3DSceneCameraPreset));
        }
        if (currentImage.HasShape3DSceneCameraZoomThousandthPercent != requested.Image.HasShape3DSceneCameraZoomThousandthPercent ||
            currentImage.Shape3DSceneCameraZoomThousandthPercent != requested.Image.Shape3DSceneCameraZoomThousandthPercent)
        {
            if (!currentImage.HasShape3DSceneCameraZoomThousandthPercent ||
                !requested.Image.HasShape3DSceneCameraZoomThousandthPercent ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraZoom(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera zoom is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            if (camera is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera owner is outside the bounded source-preserving profile.");
            camera.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "zoom",
                string.Empty,
                requested.Image.Shape3DSceneCameraZoomThousandthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneCameraFov60000 != requested.Image.HasShape3DSceneCameraFov60000 ||
            currentImage.Shape3DSceneCameraFov60000 != requested.Image.Shape3DSceneCameraFov60000)
        {
            if (!currentImage.HasShape3DSceneCameraFov60000 ||
                !requested.Image.HasShape3DSceneCameraFov60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraFov(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera FOV is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            if (camera is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera owner is outside the bounded source-preserving profile.");
            camera.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "fov",
                string.Empty,
                requested.Image.Shape3DSceneCameraFov60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneCameraRotationLatitude60000 != requested.Image.HasShape3DSceneCameraRotationLatitude60000 ||
            currentImage.Shape3DSceneCameraRotationLatitude60000 != requested.Image.Shape3DSceneCameraRotationLatitude60000)
        {
            if (!currentImage.HasShape3DSceneCameraRotationLatitude60000 ||
                !requested.Image.HasShape3DSceneCameraRotationLatitude60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraRotationLatitude(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation latitude is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            var rotation = camera?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "lat",
                string.Empty,
                requested.Image.Shape3DSceneCameraRotationLatitude60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneCameraRotationLongitude60000 != requested.Image.HasShape3DSceneCameraRotationLongitude60000 ||
            currentImage.Shape3DSceneCameraRotationLongitude60000 != requested.Image.Shape3DSceneCameraRotationLongitude60000)
        {
            if (!currentImage.HasShape3DSceneCameraRotationLongitude60000 ||
                !requested.Image.HasShape3DSceneCameraRotationLongitude60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraRotationLongitude(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation longitude is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            var rotation = camera?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "lon",
                string.Empty,
                requested.Image.Shape3DSceneCameraRotationLongitude60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneCameraRotationRevolution60000 != requested.Image.HasShape3DSceneCameraRotationRevolution60000 ||
            currentImage.Shape3DSceneCameraRotationRevolution60000 != requested.Image.Shape3DSceneCameraRotationRevolution60000)
        {
            if (!currentImage.HasShape3DSceneCameraRotationRevolution60000 ||
                !requested.Image.HasShape3DSceneCameraRotationRevolution60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneCameraRotationRevolution(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation revolution is outside the bounded source-preserving profile.");
            var camera = scene3d.GetFirstChild<A.Camera>();
            var rotation = camera?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene camera rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "rev",
                string.Empty,
                requested.Image.Shape3DSceneCameraRotationRevolution60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (!string.Equals(currentImage.Shape3DSceneLightRigPreset, requested.Image.Shape3DSceneLightRigPreset, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DSceneLightRigPreset) ||
                string.IsNullOrEmpty(requested.Image.Shape3DSceneLightRigPreset) ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneLightRigPreset(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig preset is outside the bounded source-preserving profile.");
            var lightRig = scene3d.GetFirstChild<A.LightRig>();
            if (lightRig is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig owner is outside the bounded source-preserving profile.");
            lightRig.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "rig",
                string.Empty,
                requested.Image.Shape3DSceneLightRigPreset));
        }
        if (!string.Equals(currentImage.Shape3DSceneLightRigDirection, requested.Image.Shape3DSceneLightRigDirection, StringComparison.Ordinal))
        {
            if (string.IsNullOrEmpty(currentImage.Shape3DSceneLightRigDirection) ||
                string.IsNullOrEmpty(requested.Image.Shape3DSceneLightRigDirection) ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneLightRigDirection(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig direction is outside the bounded source-preserving profile.");
            var lightRig = scene3d.GetFirstChild<A.LightRig>();
            if (lightRig is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig owner is outside the bounded source-preserving profile.");
            lightRig.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dir",
                string.Empty,
                requested.Image.Shape3DSceneLightRigDirection));
        }
        if (currentImage.HasShape3DSceneLightRigRotationLatitude60000 != requested.Image.HasShape3DSceneLightRigRotationLatitude60000 ||
            currentImage.Shape3DSceneLightRigRotationLatitude60000 != requested.Image.Shape3DSceneLightRigRotationLatitude60000)
        {
            if (!currentImage.HasShape3DSceneLightRigRotationLatitude60000 ||
                !requested.Image.HasShape3DSceneLightRigRotationLatitude60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation latitude is outside the bounded source-preserving profile.");
            var lightRig = scene3d.GetFirstChild<A.LightRig>();
            var rotation = lightRig?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "lat",
                string.Empty,
                requested.Image.Shape3DSceneLightRigRotationLatitude60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneLightRigRotationLongitude60000 != requested.Image.HasShape3DSceneLightRigRotationLongitude60000 ||
            currentImage.Shape3DSceneLightRigRotationLongitude60000 != requested.Image.Shape3DSceneLightRigRotationLongitude60000)
        {
            if (!currentImage.HasShape3DSceneLightRigRotationLongitude60000 ||
                !requested.Image.HasShape3DSceneLightRigRotationLongitude60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation longitude is outside the bounded source-preserving profile.");
            var lightRig = scene3d.GetFirstChild<A.LightRig>();
            var rotation = lightRig?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "lon",
                string.Empty,
                requested.Image.Shape3DSceneLightRigRotationLongitude60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneLightRigRotationRevolution60000 != requested.Image.HasShape3DSceneLightRigRotationRevolution60000 ||
            currentImage.Shape3DSceneLightRigRotationRevolution60000 != requested.Image.Shape3DSceneLightRigRotationRevolution60000)
        {
            if (!currentImage.HasShape3DSceneLightRigRotationRevolution60000 ||
                !requested.Image.HasShape3DSceneLightRigRotationRevolution60000 ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation revolution is outside the bounded source-preserving profile.");
            var lightRig = scene3d.GetFirstChild<A.LightRig>();
            var rotation = lightRig?.GetFirstChild<A.Rotation>();
            if (rotation is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene light-rig rotation owner is outside the bounded source-preserving profile.");
            rotation.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "rev",
                string.Empty,
                requested.Image.Shape3DSceneLightRigRotationRevolution60000.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropAnchorXEmu != requested.Image.HasShape3DSceneBackdropAnchorXEmu ||
            currentImage.Shape3DSceneBackdropAnchorXEmu != requested.Image.Shape3DSceneBackdropAnchorXEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropAnchorXEmu ||
                !requested.Image.HasShape3DSceneBackdropAnchorXEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropAnchorX(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor X is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var anchor = backdrop?.GetFirstChild<A.Anchor>();
            if (anchor is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor owner is outside the bounded source-preserving profile.");
            anchor.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "x",
                string.Empty,
                requested.Image.Shape3DSceneBackdropAnchorXEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropAnchorYEmu != requested.Image.HasShape3DSceneBackdropAnchorYEmu ||
            currentImage.Shape3DSceneBackdropAnchorYEmu != requested.Image.Shape3DSceneBackdropAnchorYEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropAnchorYEmu ||
                !requested.Image.HasShape3DSceneBackdropAnchorYEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropAnchorY(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor Y is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var anchor = backdrop?.GetFirstChild<A.Anchor>();
            if (anchor is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor owner is outside the bounded source-preserving profile.");
            anchor.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "y",
                string.Empty,
                requested.Image.Shape3DSceneBackdropAnchorYEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropAnchorZEmu != requested.Image.HasShape3DSceneBackdropAnchorZEmu ||
            currentImage.Shape3DSceneBackdropAnchorZEmu != requested.Image.Shape3DSceneBackdropAnchorZEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropAnchorZEmu ||
                !requested.Image.HasShape3DSceneBackdropAnchorZEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropAnchorZ(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor Z is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var anchor = backdrop?.GetFirstChild<A.Anchor>();
            if (anchor is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop anchor owner is outside the bounded source-preserving profile.");
            anchor.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "z",
                string.Empty,
                requested.Image.Shape3DSceneBackdropAnchorZEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropNormalDxEmu != requested.Image.HasShape3DSceneBackdropNormalDxEmu ||
            currentImage.Shape3DSceneBackdropNormalDxEmu != requested.Image.Shape3DSceneBackdropNormalDxEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropNormalDxEmu ||
                !requested.Image.HasShape3DSceneBackdropNormalDxEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropNormalDx(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal X is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var normal = backdrop?.GetFirstChild<A.Normal>();
            if (normal is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal owner is outside the bounded source-preserving profile.");
            normal.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dx",
                string.Empty,
                requested.Image.Shape3DSceneBackdropNormalDxEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropNormalDyEmu != requested.Image.HasShape3DSceneBackdropNormalDyEmu ||
            currentImage.Shape3DSceneBackdropNormalDyEmu != requested.Image.Shape3DSceneBackdropNormalDyEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropNormalDyEmu ||
                !requested.Image.HasShape3DSceneBackdropNormalDyEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropNormalDy(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal Y is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var normal = backdrop?.GetFirstChild<A.Normal>();
            if (normal is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal owner is outside the bounded source-preserving profile.");
            normal.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dy",
                string.Empty,
                requested.Image.Shape3DSceneBackdropNormalDyEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropNormalDzEmu != requested.Image.HasShape3DSceneBackdropNormalDzEmu ||
            currentImage.Shape3DSceneBackdropNormalDzEmu != requested.Image.Shape3DSceneBackdropNormalDzEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropNormalDzEmu ||
                !requested.Image.HasShape3DSceneBackdropNormalDzEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropNormalDz(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal Z is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var normal = backdrop?.GetFirstChild<A.Normal>();
            if (normal is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop normal owner is outside the bounded source-preserving profile.");
            normal.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dz",
                string.Empty,
                requested.Image.Shape3DSceneBackdropNormalDzEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropUpDxEmu != requested.Image.HasShape3DSceneBackdropUpDxEmu ||
            currentImage.Shape3DSceneBackdropUpDxEmu != requested.Image.Shape3DSceneBackdropUpDxEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropUpDxEmu ||
                !requested.Image.HasShape3DSceneBackdropUpDxEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropUpDx(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector X is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var up = backdrop?.GetFirstChild<A.UpVector>();
            if (up is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector owner is outside the bounded source-preserving profile.");
            up.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dx",
                string.Empty,
                requested.Image.Shape3DSceneBackdropUpDxEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropUpDyEmu != requested.Image.HasShape3DSceneBackdropUpDyEmu ||
            currentImage.Shape3DSceneBackdropUpDyEmu != requested.Image.Shape3DSceneBackdropUpDyEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropUpDyEmu ||
                !requested.Image.HasShape3DSceneBackdropUpDyEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropUpDy(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector Y is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var up = backdrop?.GetFirstChild<A.UpVector>();
            if (up is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector owner is outside the bounded source-preserving profile.");
            up.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dy",
                string.Empty,
                requested.Image.Shape3DSceneBackdropUpDyEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasShape3DSceneBackdropUpDzEmu != requested.Image.HasShape3DSceneBackdropUpDzEmu ||
            currentImage.Shape3DSceneBackdropUpDzEmu != requested.Image.Shape3DSceneBackdropUpDzEmu)
        {
            if (!currentImage.HasShape3DSceneBackdropUpDzEmu ||
                !requested.Image.HasShape3DSceneBackdropUpDzEmu ||
                properties.Elements<A.Scene3DType>().Count() != 1 ||
                properties.GetFirstChild<A.Scene3DType>() is not { } scene3d ||
                !PptxShape3DCodec.TryReadSceneBackdropUpDz(scene3d, out _))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector Z is outside the bounded source-preserving profile.");
            var backdrop = scene3d.GetFirstChild<A.Backdrop>();
            var up = backdrop?.GetFirstChild<A.UpVector>();
            if (up is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation image {requested.Id} 3-D scene backdrop up-vector owner is outside the bounded source-preserving profile.");
            up.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "dz",
                string.Empty,
                requested.Image.Shape3DSceneBackdropUpDzEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        if (currentImage.HasOpacityThousandthPercent != requested.Image.HasOpacityThousandthPercent ||
            currentImage.OpacityThousandthPercent != requested.Image.OpacityThousandthPercent)
            ApplyOpacity(blip, requested.Image.HasOpacityThousandthPercent ? requested.Image.OpacityThousandthPercent : null);
        var customMaskChanged = !currentImage.CustomMaskPaths.SequenceEqual(requested.Image.CustomMaskPaths);
        var presetMaskChanged = !currentImage.MaskPreset.Equals(requested.Image.MaskPreset, StringComparison.Ordinal) ||
            !currentImage.MaskPresetAdjustments.SequenceEqual(requested.Image.MaskPresetAdjustments);
        if (customMaskChanged && requested.Image.CustomMaskPaths.Count > 0)
        {
            // A literal custom mask is an owner-local picture geometry. Keep
            // the native picture relationship and replace the existing
            // preset/custom geometry only when the requested profile remains
            // fully bounded.  PptxCustomGeometryCodec owns the geometry
            // replacement, so a preset <-> custom transition does not require
            // a new picture or any relationship rewrite.
            PptxCustomGeometryCodec.Validate(CustomMaskShape(requested.Image), requested.Id + " image mask");
            PptxCustomGeometryCodec.Apply(properties, CustomMaskShape(requested.Image), requested.Id + " image mask");
        }
        else if (requested.Image.CustomMaskPaths.Count == 0 && (presetMaskChanged || customMaskChanged))
        {
            var maskPreset = requested.Image.MaskPreset.Length == 0 ? "rect" : requested.Image.MaskPreset;
            PptxPresetGeometryAdjustmentCodec.Validate(maskPreset, requested.Image.MaskPresetAdjustments, requested.Id + " image mask");
            if (geometry is A.PresetGeometry presetGeometry)
            {
                presetGeometry.Preset = MaskPreset(requested.Image.MaskPreset);
                PptxPresetGeometryAdjustmentCodec.Apply(presetGeometry, maskPreset, requested.Image.MaskPresetAdjustments, requested.Id + " image mask");
            }
            else
            {
                // A custom -> preset transition is still local to the
                // picture's shape properties.  Build a bounded preset shape
                // and let the shared geometry codec replace the old node.
                var presetShape = new PresentationShape { Geometry = maskPreset };
                presetShape.PresetAdjustments.Add(requested.Image.MaskPresetAdjustments);
                PptxCustomGeometryCodec.Apply(properties, presetShape, requested.Id + " image mask");
            }
        }
        else if (customMaskChanged || presetMaskChanged)
        {
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation image {requested.Id} mask topology is outside the bounded preset/custom profile.");
        }
        if (!Equals(currentImage.Border, requested.Image.Border)) ApplyBorder(properties, requested.Image.Border);
        if (!Equals(currentImage.Shadow, requested.Image.Shadow)) PptxShadowCodec.Apply(properties, requested.Image.Shadow);
        if (!Equals(currentImage.Glow, requested.Image.Glow)) PptxGlowCodec.Apply(properties, requested.Image.Glow);
        if (!Equals(currentImage.InnerShadow, requested.Image.InnerShadow)) PptxInnerShadowCodec.Apply(properties, requested.Image.InnerShadow);
        if (!Equals(currentImage.Reflection, requested.Image.Reflection)) PptxReflectionCodec.Apply(properties, requested.Image.Reflection);
        if (!Equals(currentImage.SoftEdge, requested.Image.SoftEdge)) PptxSoftEdgeCodec.Apply(properties, requested.Image.SoftEdge);
    }

    internal static void ScrubModeledContent(P.Picture source)
    {
        if (source.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            nonVisual.Name = string.Empty;
            PptxNonVisualAccessibilityCodec.ScrubResidualModeledContent(nonVisual);
            nonVisual.Description = string.Empty;
        }
        if (source.BlipFill?.GetFirstChild<A.Blip>() is { } blip)
        {
            blip.Embed = string.Empty;
            SvgFallbackElement(blip)?.SetAttribute(new OpenXmlAttribute(
                "r",
                "embed",
                OfficeRelationshipsNamespace,
                string.Empty));
            blip.GetFirstChild<A.AlphaModulationFixed>()?.Remove();
        }
        source.BlipFill?.GetFirstChild<A.SourceRectangle>()?.Remove();
        if (source.BlipFill?.ChildElements.LastOrDefault() is A.Tile or A.Stretch)
        {
            source.BlipFill.ChildElements.Last().Remove();
            source.BlipFill.Append(new A.Stretch(new A.FillRectangle()));
        }
        if (source.ShapeProperties?.Transform2D is { } transform)
        {
            if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
            if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
            transform.Rotation = null;
            transform.HorizontalFlip = null;
            transform.VerticalFlip = null;
        }
        if (source.ShapeProperties is { } properties)
        {
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry)
            {
                // Recreate the sentinel instead of mutating the existing
                // node. Open XML SDK can serialize namespace declarations in
                // a different attribute order after a geometry type change;
                // using one fresh sentinel for every preset/custom source
                // makes the residual hash insensitive to that modeled slot
                // representation detail.
                var preset = new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle,
                };
                geometry.InsertAfterSelf(preset);
                geometry.Remove();
            }
            else
            {
                // Custom geometry is a modeled mask profile.  Normalize it to
                // the same empty rectangle sentinel used by preset masks so
                // a bounded preset <-> custom transition does not look like
                // an unrelated native child insertion/deletion in the
                // residual hash.
                if (properties.GetFirstChild<A.CustomGeometry>() is { } custom)
                {
                    var preset = new A.PresetGeometry(new A.AdjustValueList())
                    {
                        Preset = A.ShapeTypeValues.Rectangle,
                    };
                    custom.InsertAfterSelf(preset);
                    custom.Remove();
                }
            }
            properties.GetFirstChild<A.Outline>()?.Remove();
            properties.GetFirstChild<A.EffectList>()?.Remove();
        }
    }

    private static bool TryParts(
        P.Picture source,
        out P.NonVisualDrawingProperties nonVisual,
        out A.Blip blip,
        out P.ShapeProperties properties,
        out A.Transform2D transform,
        out OpenXmlElement geometry,
        out A.SourceRectangle? crop,
        out bool tiled)
    {
        nonVisual = null!;
        blip = null!;
        properties = null!;
        transform = null!;
        geometry = null!;
        crop = null;
        tiled = false;
        var nonVisualContainer = source.NonVisualPictureProperties;
        var fill = source.BlipFill;
        var pictureProperties = source.ShapeProperties;
        if (nonVisualContainer?.NonVisualDrawingProperties is not { } nv ||
            nonVisualContainer.NonVisualPictureDrawingProperties is null ||
            nonVisualContainer.ApplicationNonVisualDrawingProperties is null ||
            fill is null || pictureProperties is null) return false;
        var fillChildren = fill.ChildElements.ToArray();
        if (source.ChildElements.Count != 3 ||
            nonVisualContainer.ChildElements.Count != 3 ||
            fillChildren.Length is < 2 or > 3 || fillChildren[0] is not A.Blip embedded ||
            fillChildren[^1] is not (A.Stretch or A.Tile) ||
            fillChildren.Length == 3 && fillChildren[1] is not A.SourceRectangle ||
            fillChildren[^1] is A.Stretch stretch && !StretchSupported(stretch) ||
            fillChildren[^1] is A.Tile tile && !TileSupported(tile) ||
            !BlipSupported(embedded) || !BlipFillSupported(fill) ||
            !ShapePropertiesSupported(pictureProperties) ||
            pictureProperties.Elements<A.Transform2D>().SingleOrDefault() is not { } xfrm ||
            !TransformSupported(xfrm)) return false;
        var geometries = pictureProperties.ChildElements
            .Where(child => child is A.PresetGeometry or A.CustomGeometry)
            .ToArray();
        if (geometries.Length != 1) return false;
        if (geometries[0] is A.PresetGeometry presetGeometry)
        {
            if (presetGeometry.Preset is null || presetGeometry.GetAttributes().Count != 1 ||
                presetGeometry.GetAttributes()[0].LocalName != "prst" || presetGeometry.GetAttributes()[0].NamespaceUri.Length != 0 ||
                presetGeometry.ChildElements.Count != 1 || presetGeometry.GetFirstChild<A.AdjustValueList>() is not { } adjustments ||
                adjustments.HasAttributes) return false;
        }
        else if (geometries[0] is not A.CustomGeometry customGeometry ||
                 !PptxCustomGeometryCodec.Supports(
                     customGeometry,
                     xfrm.Extents?.Cx?.Value ?? 0,
                     xfrm.Extents?.Cy?.Value ?? 0)) return false;
        crop = fillChildren.Length == 3 ? (A.SourceRectangle)fillChildren[1] : null;
        tiled = fillChildren[^1] is A.Tile;
        if (crop is not null && !CropSupported(crop)) return false;
        nonVisual = nv;
        blip = embedded;
        properties = pictureProperties;
        transform = xfrm;
        geometry = geometries[0];
        return true;
    }

    private static bool BlipSupported(A.Blip blip)
    {
        var attributes = blip.GetAttributes().ToArray();
        var embed = attributes.Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        var link = attributes.Where(attribute => attribute.LocalName == "link" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        if (embed.Length != 1 || link.Length != 0 || string.IsNullOrWhiteSpace(embed[0].Value)) return false;
        if (attributes.Any(attribute =>
                !((attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)) ||
                  (attribute.LocalName == "cstate" && attribute.NamespaceUri.Length == 0)))) return false;
        if (attributes.Any(attribute => attribute.LocalName == "cstate" && attribute.Value is not ("email" or "screen" or "print" or "hqprint"))) return false;
        return BlipChildrenSupported(blip);
    }

    private static bool BlipChildrenSupported(A.Blip blip)
    {
        var extensionLists = blip.ChildElements.Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace).ToArray();
        if (blip.ChildElements.Any(child => child.LocalName is not ("alphaModFix" or "clrChange" or "extLst") || child.NamespaceUri != DrawingNamespace) ||
            extensionLists.Length > 1) return false;
        var alpha = blip.ChildElements.Where(child => child.LocalName == "alphaModFix").ToArray();
        var colors = blip.ChildElements.Where(child => child.LocalName == "clrChange").ToArray();
        return alpha.Length <= 1 && colors.Length <= 1 &&
               alpha.All(AlphaModulationSupported) && colors.All(ColorChangeSupported) &&
               extensionLists.All(BlipExtensionListSupported);
    }

    private static bool AlphaModulationSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        return element.NamespaceUri == DrawingNamespace && !element.HasChildren &&
               attributes.All(attribute => attribute.LocalName == "amt" && attribute.NamespaceUri.Length == 0) &&
               attributes.Length <= 1 && attributes.All(attribute => UIntInRange(attribute.Value ?? string.Empty, 0, 100_000));
    }

    private static bool ColorChangeSupported(OpenXmlElement element)
    {
        if (element.NamespaceUri != DrawingNamespace || element.HasAttributes || element.ChildElements.Count != 2) return false;
        var from = element.ChildElements.FirstOrDefault(child => child.LocalName == "clrFrom" && child.NamespaceUri == DrawingNamespace);
        var to = element.ChildElements.FirstOrDefault(child => child.LocalName == "clrTo" && child.NamespaceUri == DrawingNamespace);
        return from is not null && to is not null && element.ChildElements.All(child => ReferenceEquals(child, from) || ReferenceEquals(child, to)) &&
               ColorChoiceSupported(from, allowAlpha: false) && ColorChoiceSupported(to, allowAlpha: true);
    }

    private static bool ColorChoiceSupported(OpenXmlElement element, bool allowAlpha)
    {
        if (element.HasAttributes || element.ChildElements.Count != 1 || element.ChildElements[0].LocalName != "srgbClr" ||
            element.ChildElements[0].NamespaceUri != DrawingNamespace) return false;
        var color = element.ChildElements[0];
        var attributes = color.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.LocalName != "val" || attribute.NamespaceUri.Length != 0) ||
            attributes.Length != 1 || !IsHexColor(attributes[0].Value ?? string.Empty)) return false;
        var alpha = color.ChildElements.ToArray();
        return alpha.Length == 0 || allowAlpha && alpha.Length == 1 && alpha[0].LocalName == "alpha" &&
               alpha[0].NamespaceUri == DrawingNamespace && alpha[0].ChildElements.Count == 0 &&
               alpha[0].GetAttributes().Count == 1 && alpha[0].GetAttributes()[0].LocalName == "val" &&
               alpha[0].GetAttributes()[0].NamespaceUri.Length == 0 && UIntInRange(alpha[0].GetAttributes()[0].Value ?? string.Empty, 0, 100_000);
    }

    private static bool BlipExtensionListSupported(OpenXmlElement element)
    {
        if (element is not A.BlipExtensionList || element.NamespaceUri != DrawingNamespace) return false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var extension in element.ChildElements)
        {
            if (extension.LocalName != "ext" || extension.NamespaceUri != DrawingNamespace || extension.GetAttributes().Count != 1 ||
                extension.GetAttributes()[0].LocalName != "uri" || extension.GetAttributes()[0].NamespaceUri.Length != 0 ||
                !seen.Add(extension.GetAttributes()[0].Value ?? string.Empty) || extension.ChildElements.Count != 1) return false;
            var uri = extension.GetAttributes()[0].Value ?? string.Empty;
            if (uri == UseLocalDpiUri && !UseLocalDpiSupported(extension.ChildElements[0])) return false;
            if (uri == SvgExtensionUri && !SvgFallbackSupported(extension.ChildElements[0])) return false;
            if (uri is not (UseLocalDpiUri or SvgExtensionUri)) return false;
        }
        return true;
    }

    private static bool UseLocalDpiSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        return element.LocalName == "useLocalDpi" && element.NamespaceUri == Office2010Namespace && !element.HasChildren &&
               attributes.All(attribute => attribute.LocalName == "val" && attribute.NamespaceUri.Length == 0 && attribute.Value is "0" or "1") &&
               attributes.Length <= 1;
    }

    private static bool SvgFallbackSupported(OpenXmlElement element)
    {
        var attributes = element.GetAttributes().ToArray();
        var embeds = attributes.Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri)).ToArray();
        return element.LocalName == "svgBlip" && element.NamespaceUri == SvgBlipNamespace && !element.HasChildren && embeds.Length == 1 &&
               attributes.Length == 1 && !string.IsNullOrWhiteSpace(embeds[0].Value);
    }

    private static OpenXmlElement? SvgFallbackElement(A.Blip blip) => blip.ChildElements
        .Where(child => child.LocalName == "extLst" && child.NamespaceUri == DrawingNamespace)
        .SelectMany(child => child.ChildElements)
        .Where(child => child.LocalName == "ext" && child.NamespaceUri == DrawingNamespace &&
            child.GetAttributes().Any(attribute => attribute.LocalName == "uri" && attribute.NamespaceUri.Length == 0 && attribute.Value == SvgExtensionUri))
        .SelectMany(child => child.ChildElements)
        .SingleOrDefault(child => child.LocalName == "svgBlip" && child.NamespaceUri == SvgBlipNamespace);

    private static string? SvgFallbackRelationshipId(A.Blip blip)
    {
        var fallback = SvgFallbackElement(blip);
        if (fallback is null) return null;
        var embeds = fallback.GetAttributes()
            .Where(attribute => attribute.LocalName == "embed" && RelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        return embeds.Length == 1 ? embeds[0].Value : string.Empty;
    }

    private static bool HasSvgFallback(A.Blip blip) => SvgFallbackElement(blip) is not null;

    private static bool BlipFillSupported(P.BlipFill fill)
    {
        var attributes = fill.GetAttributes().ToArray();
        return attributes.All(attribute => attribute.LocalName == "rotWithShape" && attribute.NamespaceUri.Length == 0 && attribute.Value is "0" or "1");
    }

    private static bool StretchSupported(A.Stretch stretch)
    {
        if (stretch.HasAttributes || stretch.ChildElements.Count > 1) return false;
        return stretch.ChildElements.Count == 0 || stretch.GetFirstChild<A.FillRectangle>() is { } fillRect && !fillRect.HasAttributes && !fillRect.HasChildren;
    }

    private static bool TileSupported(A.Tile tile) => !tile.HasAttributes && !tile.HasChildren;

    private static bool ShapePropertiesSupported(P.ShapeProperties properties)
    {
        var attributes = properties.GetAttributes().ToArray();
        if (attributes.Any(attribute => attribute.LocalName != "bwMode" || attribute.NamespaceUri.Length != 0 || attribute.Value is not ("auto" or "gray" or "ltGray" or "invGray" or "grayWhite" or "blackGray" or "blackWhite" or "clr"))) return false;
        if (properties.ChildElements.Count < 2 || properties.ChildElements.Count > 7) return false;
        if (properties.ChildElements.Count(child => child is A.Transform2D) != 1 ||
            properties.ChildElements.Count(child => child is A.PresetGeometry or A.CustomGeometry) != 1) return false;
        return properties.ChildElements.All(child => child switch
        {
            A.Transform2D or A.PresetGeometry or A.CustomGeometry => true,
            A.Shape3DType shape3d => PptxShape3DCodec.TryReadDepth(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadExtrusionHeight(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadContourWidth(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadPresetMaterial(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelTopWidth(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelTopHeight(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelTopPreset(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelBottomWidth(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelBottomHeight(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadBevelBottomPreset(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadContourRgb(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadExtrusionRgb(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadContourColorScheme(shape3d, out _) ||
                                     PptxShape3DCodec.TryReadExtrusionColorScheme(shape3d, out _),
            A.Scene3DType scene3d => PptxShape3DCodec.TryReadSceneCameraPreset(scene3d, out _) ||
                                     PptxShape3DCodec.TryReadSceneCameraRotationLatitude(scene3d, out _) ||
                                     PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(scene3d, out _) ||
                                     PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(scene3d, out _) ||
                                     PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(scene3d, out _) ||
                                     PptxShape3DCodec.TryReadSceneBackdropAnchorX(scene3d, out _),
            A.NoFill noFill => !noFill.HasAttributes && !noFill.HasChildren,
            A.Outline outline => TryReadBorder(outline, out _),
            A.EffectList => PptxShadowCodec.TryRead(properties, out _) ||
                            PptxGlowCodec.TryRead(properties, out _) ||
                            PptxInnerShadowCodec.TryRead(properties, out _) ||
                            PptxReflectionCodec.TryRead(properties, out _) ||
                            PptxSoftEdgeCodec.TryRead(properties, out _),
            A.ShapePropertiesExtensionList extensions => ShapeExtensionListSupported(extensions),
            _ => false,
        });
    }

    private static bool ShapeExtensionListSupported(A.ShapePropertiesExtensionList extensions)
    {
        if (extensions.HasAttributes) return false;
        foreach (var extension in extensions.ChildElements)
        {
            var attributes = extension.GetAttributes().ToArray();
            if (extension.LocalName != "ext" || extension.NamespaceUri != DrawingNamespace || attributes.Length != 1 ||
                attributes[0].LocalName != "uri" || attributes[0].NamespaceUri.Length != 0 || attributes[0].Value != HiddenFillUri ||
                extension.ChildElements.Count != 1) return false;
            var hidden = extension.ChildElements[0];
            if (hidden.LocalName != "hiddenFill" || hidden.NamespaceUri != Office2010Namespace || hidden.ChildElements.Count != 1 || hidden.GetAttributes().Count != 0)
                return false;
            var solid = hidden.ChildElements[0];
            if (solid.LocalName != "solidFill" || solid.NamespaceUri != DrawingNamespace || solid.ChildElements.Count != 1 || solid.GetAttributes().Count != 0)
                return false;
            var color = solid.ChildElements[0];
            if (color.LocalName != "srgbClr" || color.NamespaceUri != DrawingNamespace || color.ChildElements.Count != 0 || color.GetAttributes().Count != 1 ||
                color.GetAttributes()[0].LocalName != "val" || color.GetAttributes()[0].NamespaceUri.Length != 0 || !IsHexColor(color.GetAttributes()[0].Value ?? string.Empty)) return false;
        }
        return true;
    }

    private static bool IsHexColor(string value) => value.Length == 6 && value.All(Uri.IsHexDigit);

    private static bool TryReadBorder(A.Outline? outline, out PresentationImageBorder? border)
    {
        border = null;
        if (outline is null) return true;
        // PowerPoint and Google Slides commonly serialize a picture frame's
        // default single/centered line attributes even when the frame is
        // explicitly invisible.  These attributes do not change the rendered
        // picture, so accept only the canonical defaults while retaining the
        // existing source XML until a caller explicitly edits the border.
        if (outline.CompoundLineType?.Value is { } compound &&
            !compound.Equals(A.CompoundLineValues.Single) ||
            outline.Alignment?.Value is { } alignment &&
            !alignment.Equals(A.PenAlignmentValues.Center) ||
            !HasOnlyAttributes(outline, "w", "cap", "cmpd", "algn") ||
            outline.Width?.Value is < 0 or > int.MaxValue)
            return false;
        var noFill = outline.Elements<A.NoFill>().ToArray();
        var solidFill = outline.Elements<A.SolidFill>().ToArray();
        if (noFill.Length + solidFill.Length != 1) return false;
        var noLine = noFill.Length == 1;
        if (noLine && (noFill[0].HasAttributes || noFill[0].HasChildren)) return false;

        string rgb = string.Empty;
        string scheme = string.Empty;
        uint? opacity = null;
        if (!noLine)
        {
            var solid = solidFill[0];
            if (solid.ChildElements.Count != 1 || !HasOnlyAttributes(solid)) return false;
            if (solid.FirstChild is A.RgbColorModelHex color)
            {
                if (!HasOnlyAttributes(color, "val") ||
                    color.Val?.Value is not { Length: 6 } directRgb || !directRgb.All(Uri.IsHexDigit)) return false;
                rgb = directRgb;
                var alphas = color.Elements<A.Alpha>().ToArray();
                if (color.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
                    alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
                    return false;
                if (alphas.SingleOrDefault()?.Val?.Value is { } alphaValue)
                    opacity = checked((uint)alphaValue);
            }
            else if (solid.FirstChild is A.SchemeColor schemeColor)
            {
                if (!HasOnlyAttributes(schemeColor, "val") || schemeColor.Val?.Value is not { } rawScheme ||
                    !PptxColor.TrySchemeToken(rawScheme, out scheme)) return false;
                var alphas = schemeColor.Elements<A.Alpha>().ToArray();
                if (schemeColor.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
                    alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
                    return false;
                if (alphas.SingleOrDefault()?.Val?.Value is { } alphaValue)
                    opacity = checked((uint)alphaValue);
            }
            else return false;
        }
        var dashes = outline.Elements<A.PresetDash>().ToArray();
        if (dashes.Length > 1 || dashes.SingleOrDefault() is { } dash &&
            (dash.ChildElements.Any() || !HasOnlyAttributes(dash, "val"))) return false;
        if (noLine && dashes.Length > 0) return false;
        var style = dashes.SingleOrDefault()?.Val?.Value switch
        {
            null => "solid",
            var value when value.Equals(A.PresetLineDashValues.Solid) => "solid",
            var value when value.Equals(A.PresetLineDashValues.Dash) => "dashed",
            var value when value.Equals(A.PresetLineDashValues.Dot) => "dotted",
            var value when value.Equals(A.PresetLineDashValues.DashDot) => "dash-dot",
            var value when value.Equals(A.PresetLineDashValues.LargeDashDotDot) => "dash-dot-dot",
            _ => string.Empty,
        };
        if (noLine) style = "none";
        if (style.Length == 0 || !TryCap(outline.CapType?.Value, out var cap) || !TryJoin(outline, out var join) ||
            !TryReadInertLineEnds(outline) ||
            outline.ChildElements.Any(child => child is not A.NoFill and not A.SolidFill and not A.PresetDash and not A.Round and not A.LineJoinBevel and not A.Miter and not A.HeadEnd and not A.TailEnd))
            return false;
        if (noLine) return true;
        border = new PresentationImageBorder
        {
            ColorRgb = string.IsNullOrEmpty(scheme) ? PptxColor.Normalize(rgb) : string.Empty,
            ColorScheme = scheme,
            WidthEmu = outline.Width?.Value ?? 0L,
            Style = style,
            Cap = cap,
            Join = join,
        };
        if (opacity is { } alpha)
            border.OpacityThousandthPercent = alpha;
        return true;
    }

    private static void ValidateBorder(PresentationImageBorder? border, string elementId)
    {
        if (border is null) return;
        if (string.IsNullOrEmpty(border.ColorRgb) == string.IsNullOrEmpty(border.ColorScheme))
            throw Invalid(elementId, "border must use exactly one RGB or theme color");
        if (!string.IsNullOrEmpty(border.ColorRgb)) PptxColor.Normalize(border.ColorRgb);
        else PptxColor.NormalizeScheme(border.ColorScheme);
        if (border.WidthEmu is < 0 or > int.MaxValue ||
            border.Style is not ("solid" or "dashed" or "dotted" or "dash-dot" or "dash-dot-dot") ||
            border.Cap is not ("" or "flat" or "round" or "square") ||
            border.Join is not ("" or "miter" or "round" or "bevel") ||
            border.HasOpacityThousandthPercent && border.OpacityThousandthPercent > 100_000)
            throw Invalid(elementId, "border uses unsupported width, dash, cap, join, or opacity");
    }

    private static A.Outline BuildBorder(PresentationImageBorder border)
    {
        var outline = new A.Outline { Width = checked((int)border.WidthEmu) };
        outline.CapType = border.Cap switch
        {
            "round" => A.LineCapValues.Round,
            "square" => A.LineCapValues.Square,
            "flat" => A.LineCapValues.Flat,
            _ => null,
        };
        OpenXmlElement color = !string.IsNullOrEmpty(border.ColorScheme)
            ? new A.SchemeColor { Val = PptxColor.SchemeValue(border.ColorScheme) }
            : new A.RgbColorModelHex { Val = PptxColor.Normalize(border.ColorRgb) };
        if (border.HasOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)border.OpacityThousandthPercent) });
        outline.Append(new A.SolidFill(color));
        if (border.Style != "solid") outline.Append(new A.PresetDash { Val = border.Style switch
        {
            "dotted" => A.PresetLineDashValues.Dot,
            "dash-dot" => A.PresetLineDashValues.DashDot,
            "dash-dot-dot" => A.PresetLineDashValues.LargeDashDotDot,
            _ => A.PresetLineDashValues.Dash,
        }});
        if (border.Join.Length > 0) outline.Append(border.Join switch
        {
            "round" => new A.Round(),
            "bevel" => new A.LineJoinBevel(),
            _ => new A.Miter(),
        });
        return outline;
    }

    private static void ApplyBorder(P.ShapeProperties properties, PresentationImageBorder? border)
    {
        properties.GetFirstChild<A.Outline>()?.Remove();
        if (border is null) return;
        var outline = BuildBorder(border);
        OpenXmlElement? anchor = properties.ChildElements.LastOrDefault(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill);
        anchor ??= properties.GetFirstChild<A.PresetGeometry>();
        anchor ??= properties.GetFirstChild<A.Transform2D>();
        if (anchor is null) properties.PrependChild(outline);
        else properties.InsertAfter(outline, anchor);
    }

    private static A.PresetGeometry BuildMask(string value, IEnumerable<int> adjustments, string elementId)
    {
        var name = value.Length == 0 ? "rect" : value;
        var output = new A.PresetGeometry { Preset = MaskPreset(value) };
        PptxPresetGeometryAdjustmentCodec.Apply(output, name, adjustments, elementId + " image mask");
        return output;
    }

    private static bool TryReadMask(
        OpenXmlElement geometry,
        long widthEmu,
        long heightEmu,
        PresentationImage image)
    {
        if (geometry is A.PresetGeometry presetGeometry)
        {
            if (presetGeometry.Preset?.Value is not { } preset ||
                !PptxCustomGeometryCodec.TryPresetName(preset, out var maskPreset) ||
                !PptxPresetGeometryAdjustmentCodec.TryRead(presetGeometry, maskPreset, out var maskAdjustments))
                return false;
            if (!maskPreset.Equals("rect", StringComparison.Ordinal)) image.MaskPreset = maskPreset;
            image.MaskPresetAdjustments.Add(maskAdjustments);
            return true;
        }
        if (geometry is not A.CustomGeometry customGeometry) return false;
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = widthEmu,
            HeightEmu = heightEmu,
        };
        PptxCustomGeometryCodec.Read(customGeometry, widthEmu, heightEmu, shape);
        if (shape.CustomPaths.Count == 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
            shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 || shape.TextRectangle is not null)
            return false;
        image.CustomMaskPaths.Add(shape.CustomPaths);
        return true;
    }

    private static PresentationShape CustomMaskShape(PresentationImage image)
    {
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = image.WidthEmu,
            HeightEmu = image.HeightEmu,
        };
        shape.CustomPaths.Add(image.CustomMaskPaths);
        return shape;
    }

    private static A.ShapeTypeValues MaskPreset(string value)
    {
        var name = value.Length == 0 ? "rect" : value;
        if (!PptxCustomGeometryCodec.TryPreset(name, out var preset))
            throw new CodecException("unsupported_presentation_image", $"Presentation image mask preset {name} is unsupported.");
        return preset;
    }

    private static void ApplyOpacity(A.Blip blip, uint? opacity)
    {
        blip.GetFirstChild<A.AlphaModulationFixed>()?.Remove();
        if (opacity is null) return;
        var alpha = new A.AlphaModulationFixed { Amount = checked((int)opacity.Value) };
        var before = blip.ChildElements.FirstOrDefault(child => child is A.ColorChange or A.BlipExtensionList);
        if (before is null) blip.Append(alpha);
        else blip.InsertBefore(alpha, before);
    }

    private static bool TryCap(A.LineCapValues? value, out string cap)
    {
        cap = string.Empty;
        if (value is null) return true;
        if (value.Value.Equals(A.LineCapValues.Flat)) cap = "flat";
        else if (value.Value.Equals(A.LineCapValues.Round)) cap = "round";
        else if (value.Value.Equals(A.LineCapValues.Square)) cap = "square";
        else return false;
        return true;
    }

    private static bool TryJoin(A.Outline outline, out string join)
    {
        join = string.Empty;
        var joins = outline.ChildElements.Where(child => child is A.Round or A.LineJoinBevel or A.Miter).ToArray();
        if (joins.Length > 1) return false;
        if (joins.SingleOrDefault() is A.Round round)
        {
            if (round.HasAttributes || round.HasChildren) return false;
            join = "round";
        }
        else if (joins.SingleOrDefault() is A.LineJoinBevel bevel)
        {
            if (bevel.HasAttributes || bevel.HasChildren) return false;
            join = "bevel";
        }
        else if (joins.SingleOrDefault() is A.Miter miter)
        {
            // 800000 is the DrawingML default miter limit.  It is inert for
            // the picture frame profile, but is emitted explicitly by several
            // native exporters alongside noFill.
            if (miter.HasChildren || miter.GetAttributes().Any(attribute =>
                    attribute.NamespaceUri.Length != 0 || attribute.LocalName != "lim" || attribute.Value != "800000")) return false;
            join = "miter";
        }
        return true;
    }

    private static bool TryReadInertLineEnds(A.Outline outline)
    {
        var heads = outline.Elements<A.HeadEnd>().ToArray();
        var tails = outline.Elements<A.TailEnd>().ToArray();
        if (heads.Length > 1 || tails.Length > 1) return false;
        return InertLineEnd(heads.SingleOrDefault()) && InertLineEnd(tails.SingleOrDefault());
    }

    private static bool InertLineEnd(OpenXmlElement? source)
    {
        if (source is null) return true;
        if (source.ChildElements.Any() || !HasOnlyAttributes(source, "type", "w", "len")) return false;
        var type = source switch
        {
            A.HeadEnd head => head.Type?.Value,
            A.TailEnd tail => tail.Type?.Value,
            _ => null,
        };
        return type is null || type.Value.Equals(A.LineEndValues.None);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }

    private static bool UIntInRange(string value, uint min, uint max) =>
        uint.TryParse(value, out var parsed) && parsed >= min && parsed <= max;

    private static bool RelationshipNamespace(string value) => PptxNativeObjectCatalog.IsRelationshipNamespace(value);

    private static PresentationNonVisualAccessibility? Accessibility(PresentationImage image)
    {
        if (image.AccessibilityTitle.Length == 0 && image.AltText.Length == 0 && !image.HasAccessibilityDecorative) return null;
        var value = new PresentationNonVisualAccessibility();
        if (image.AccessibilityTitle.Length > 0) value.Title = image.AccessibilityTitle;
        if (image.AltText.Length > 0) value.Description = image.AltText;
        if (image.HasAccessibilityDecorative) value.Decorative = image.AccessibilityDecorative;
        return value;
    }

    private static bool AccessibilityEqual(PresentationImage left, PresentationImage right) =>
        left.AccessibilityTitle.Equals(right.AccessibilityTitle, StringComparison.Ordinal) &&
        left.AltText.Equals(right.AltText, StringComparison.Ordinal) &&
        left.HasAccessibilityDecorative == right.HasAccessibilityDecorative &&
        left.AccessibilityDecorative == right.AccessibilityDecorative;

    private static bool CropSupported(A.SourceRectangle source)
    {
        var known = new HashSet<string>(StringComparer.Ordinal) { "l", "t", "r", "b" };
        return !source.HasChildren && source.GetAttributes().All(attribute => known.Contains(attribute.LocalName)) &&
               CropValuesValid(ReadCrop(source));
    }

    private static PresentationImageCrop ReadCrop(A.SourceRectangle source) => new()
    {
        LeftThousandthPercent = source.Left?.Value ?? 0,
        TopThousandthPercent = source.Top?.Value ?? 0,
        RightThousandthPercent = source.Right?.Value ?? 0,
        BottomThousandthPercent = source.Bottom?.Value ?? 0,
    };

    private static bool CropValuesValid(PresentationImageCrop crop) =>
        crop.LeftThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.TopThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.RightThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.BottomThousandthPercent is >= -100_000 and <= 100_000 &&
        crop.LeftThousandthPercent + crop.RightThousandthPercent < 100_000 &&
        crop.TopThousandthPercent + crop.BottomThousandthPercent < 100_000;

    private static A.SourceRectangle BuildCrop(PresentationImageCrop crop) => new()
    {
        Left = crop.LeftThousandthPercent,
        Top = crop.TopThousandthPercent,
        Right = crop.RightThousandthPercent,
        Bottom = crop.BottomThousandthPercent,
    };

    private static void ApplyCrop(P.BlipFill target, PresentationImageCrop? crop)
    {
        var current = target.GetFirstChild<A.SourceRectangle>();
        if (crop is null)
        {
            current?.Remove();
            return;
        }
        var replacement = BuildCrop(crop);
        if (current is not null)
        {
            current.InsertAfterSelf(replacement);
            current.Remove();
            return;
        }
        var blip = target.GetFirstChild<A.Blip>() ?? throw new CodecException("unsupported_presentation_edit", "Presentation picture lost its embedded blip.");
        blip.InsertAfterSelf(replacement);
    }

    private static bool TransformSupported(A.Transform2D? transform)
    {
        if (transform is null || transform.ChildElements.Count != 2 ||
            transform.Elements<A.Offset>().SingleOrDefault() is not { } offset ||
            transform.Elements<A.Extents>().SingleOrDefault() is not { } extents ||
            offset.X is null || offset.Y is null || extents.Cx is null || extents.Cy is null ||
            offset.HasChildren || extents.HasChildren || offset.GetAttributes().Count != 2 || extents.GetAttributes().Count != 2)
            return false;
        var known = new HashSet<string>(StringComparer.Ordinal) { "rot", "flipH", "flipV" };
        if (transform.GetAttributes().Any(attribute => !known.Contains(attribute.LocalName))) return false;
        var rotation = transform.Rotation?.Value;
        return rotation is null || Math.Abs((long)rotation.Value) <= 21_600_000L;
    }

    private static bool FrameCoordinateSupported(long value, bool allowNegative) =>
        value <= MaxFrameCoordinateEmu && value >= (allowNegative ? -MaxFrameCoordinateEmu : 0L);

    private static bool FrameExtentSupported(long value) => value > 0 && value <= MaxFrameCoordinateEmu;

    private static PresentationImageTransform? ReadTransform(A.Transform2D source)
    {
        var result = new PresentationImageTransform();
        if (source.Rotation is not null) result.RotationAngle60000 = source.Rotation.Value;
        if (source.HorizontalFlip is not null) result.FlipHorizontal = source.HorizontalFlip.Value;
        if (source.VerticalFlip is not null) result.FlipVertical = source.VerticalFlip.Value;
        return result.CalculateSize() == 0 ? null : result;
    }

    private static void ValidateTransform(PresentationImageTransform? transform, string elementId)
    {
        if (transform is null) return;
        if (!transform.HasRotationAngle60000 && !transform.HasFlipHorizontal && !transform.HasFlipVertical)
            throw Invalid(elementId, "transform must define rotation or a flip");
        if (transform.HasRotationAngle60000 && Math.Abs((long)transform.RotationAngle60000) > 21_600_000L)
            throw Invalid(elementId, "rotation must be between -360 and 360 degrees");
    }

    private static void ApplyTransform(A.Transform2D target, PresentationImageTransform? source)
    {
        target.Rotation = source?.HasRotationAngle60000 == true ? source.RotationAngle60000 : null;
        target.HorizontalFlip = source?.HasFlipHorizontal == true ? source.FlipHorizontal : null;
        target.VerticalFlip = source?.HasFlipVertical == true ? source.FlipVertical : null;
    }

    private static CodecException Invalid(string elementId, string message) =>
        new("invalid_presentation_image", $"Presentation image {elementId} {message}.");
}
