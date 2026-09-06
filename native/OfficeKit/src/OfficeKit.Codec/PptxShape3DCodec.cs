using System.Globalization;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Narrow source-bound owners for direct DrawingML 3-D scalars. The wider
// scene, transformed colors, and bevel graph remain source-owned; no 3-D
// scene is rebuilt.
internal static class PptxShape3DCodec
{
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const long MaxCoordinate = int.MaxValue;
    private const long MinBackdropCoordinate = -27_273_042_329_600;
    private const long MaxBackdropCoordinate = 27_273_042_316_900;
    private static readonly HashSet<string> PresetMaterials = new(StringComparer.Ordinal)
    {
        "legacyMatte", "legacyPlastic", "legacyMetal", "legacyWireframe",
        "matte", "plastic", "metal", "warmMatte", "translucentPowder",
        "powder", "dkEdge", "softEdge", "flatten",
    };
    private static readonly HashSet<string> CameraPresets = new(StringComparer.Ordinal)
    {
        "legacyObliqueTopLeft", "legacyObliqueTop", "legacyObliqueTopRight",
        "legacyObliqueLeft", "legacyObliqueFront", "legacyObliqueRight",
        "legacyObliqueBottomLeft", "legacyObliqueBottom", "legacyObliqueBottomRight",
        "legacyPerspectiveTopLeft", "legacyPerspectiveTop", "legacyPerspectiveTopRight",
        "legacyPerspectiveLeft", "legacyPerspectiveFront", "legacyPerspectiveRight",
        "legacyPerspectiveBottomLeft", "legacyPerspectiveBottom", "legacyPerspectiveBottomRight",
        "orthographicFront",
        "isometricTopUp", "isometricTopDown", "isometricBottomUp", "isometricBottomDown",
        "isometricLeftUp", "isometricLeftDown", "isometricRightUp", "isometricRightDown",
        "isometricOffAxis1Left", "isometricOffAxis1Right", "isometricOffAxis1Top",
        "isometricOffAxis2Left", "isometricOffAxis2Right", "isometricOffAxis2Top",
        "isometricOffAxis3Left", "isometricOffAxis3Right", "isometricOffAxis3Bottom",
        "isometricOffAxis4Left", "isometricOffAxis4Right", "isometricOffAxis4Bottom",
        "obliqueTopLeft", "obliqueTop", "obliqueTopRight", "obliqueLeft", "obliqueRight",
        "obliqueBottomLeft", "obliqueBottom", "obliqueBottomRight",
        "perspectiveFront", "perspectiveLeft", "perspectiveRight", "perspectiveAbove", "perspectiveBelow",
        "perspectiveAboveLeftFacing", "perspectiveAboveRightFacing",
        "perspectiveContrastingLeftFacing", "perspectiveContrastingRightFacing",
        "perspectiveHeroicLeftFacing", "perspectiveHeroicRightFacing",
        "perspectiveHeroicExtremeLeftFacing", "perspectiveHeroicExtremeRightFacing",
        "perspectiveRelaxed", "perspectiveRelaxedModerately",
    };
    private static readonly HashSet<string> LightRigPresets = new(StringComparer.Ordinal)
    {
        "legacyFlat1", "legacyFlat2", "legacyFlat3", "legacyFlat4",
        "legacyNormal1", "legacyNormal2", "legacyNormal3", "legacyNormal4",
        "legacyHarsh1", "legacyHarsh2", "legacyHarsh3", "legacyHarsh4",
        "threePt", "balanced", "soft", "harsh", "flood", "contrasting",
        "morning", "sunrise", "sunset", "chilly", "freezing", "flat", "twoPt", "glow", "brightRoom",
    };
    private static readonly HashSet<string> LightRigDirections = new(StringComparer.Ordinal)
    {
        "tl", "t", "tr", "l", "r", "bl", "b", "br",
    };
    private static readonly HashSet<string> BevelPresets = new(StringComparer.Ordinal)
    {
        "relaxedInset", "circle", "slope", "cross", "angle", "softRound",
        "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco",
    };

    internal static bool TryReadExtrusionHeight(A.Shape3DType? source, out long value) =>
        TryReadCoordinate(source, "extrusionH", allowNegative: false, out value);

    internal static bool TryReadDepth(A.Shape3DType? source, out long value) =>
        TryReadCoordinate(source, "z", allowNegative: true, out value);

    internal static bool TryReadContourWidth(A.Shape3DType? source, out long value) =>
        TryReadCoordinate(source, "contourW", allowNegative: false, out value);

    internal static bool TryReadPresetMaterial(A.Shape3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadRawAttribute(source, "prstMaterial", out var token) ||
            !PresetMaterials.Contains(token))
            return false;
        value = token;
        return true;
    }

    internal static bool TryReadBevelTopWidth(A.Shape3DType? source, out long value) =>
        TryReadBevelTopCoordinate(source, "w", out value);

    internal static bool TryReadBevelTopHeight(A.Shape3DType? source, out long value) =>
        TryReadBevelTopCoordinate(source, "h", out value);

    internal static bool TryReadBevelTopPreset(A.Shape3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadBevelTopAttribute(source, "prst", out var token) ||
            !BevelPresets.Contains(token))
            return false;
        value = token;
        return true;
    }

    internal static bool TryReadBevelBottomWidth(A.Shape3DType? source, out long value) =>
        TryReadBevelBottomCoordinate(source, "w", out value);

    internal static bool TryReadBevelBottomHeight(A.Shape3DType? source, out long value) =>
        TryReadBevelBottomCoordinate(source, "h", out value);

    internal static bool TryReadBevelBottomPreset(A.Shape3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadBevelBottomAttribute(source, "prst", out var token) ||
            !BevelPresets.Contains(token))
            return false;
        value = token;
        return true;
    }

    internal static bool TryReadContourRgb(A.Shape3DType? source, out string value) =>
        TryReadShape3dRgb(source, "contourClr", out value);

    internal static bool TryReadExtrusionRgb(A.Shape3DType? source, out string value) =>
        TryReadShape3dRgb(source, "extrusionClr", out value);

    internal static bool TryReadContourColorScheme(A.Shape3DType? source, out string value) =>
        TryReadShape3dScheme(source, "contourClr", out value);

    internal static bool TryReadExtrusionColorScheme(A.Shape3DType? source, out string value) =>
        TryReadShape3dScheme(source, "extrusionClr", out value);

    internal static bool TryReadSceneCameraPreset(A.Scene3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadScenePair(source, out var cameraPreset, out _, out _, out _, out _))
            return false;
        value = cameraPreset;
        return true;
    }

    internal static bool TryReadSceneBackdropAnchorX(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropAnchorCoordinate(source, "x", out value);

    internal static bool TryReadSceneBackdropAnchorY(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropAnchorCoordinate(source, "y", out value);

    internal static bool TryReadSceneBackdropAnchorZ(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropAnchorCoordinate(source, "z", out value);

    internal static bool TryReadSceneBackdropNormalDx(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropNormalCoordinate(source, "dx", out value);

    internal static bool TryReadSceneBackdropNormalDy(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropNormalCoordinate(source, "dy", out value);

    internal static bool TryReadSceneBackdropNormalDz(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropNormalCoordinate(source, "dz", out value);

    internal static bool TryReadSceneBackdropUpDx(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropUpCoordinate(source, "dx", out value);

    internal static bool TryReadSceneBackdropUpDy(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropUpCoordinate(source, "dy", out value);

    internal static bool TryReadSceneBackdropUpDz(A.Scene3DType? source, out long value) =>
        TryReadSceneBackdropUpCoordinate(source, "dz", out value);

    internal static bool TryReadSceneCameraZoom(A.Scene3DType? source, out long value)
    {
        value = 0;
        if (!TryReadScenePair(source, out _, out _, out _, out var cameraZoom, out _) ||
            !TryLiteralCoordinate(cameraZoom, allowNegative: false, out value))
            return false;
        return true;
    }

    internal static bool TryReadSceneCameraFov(A.Scene3DType? source, out long value)
    {
        value = 0;
        if (!TryReadScenePair(source, out _, out _, out _, out _, out var cameraFov) ||
            !TryLiteralCameraFov(cameraFov, out value))
            return false;
        return true;
    }

    internal static bool TryReadSceneLightRigPreset(A.Scene3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadScenePair(source, out _, out var lightRigPreset, out _, out _, out _))
            return false;
        value = lightRigPreset;
        return true;
    }

    internal static bool TryReadSceneLightRigDirection(A.Scene3DType? source, out string value)
    {
        value = string.Empty;
        if (!TryReadScenePair(source, out _, out _, out var lightRigDirection, out _, out _))
            return false;
        value = lightRigDirection;
        return true;
    }

    internal static bool TryReadSceneLightRigRotationLatitude(A.Scene3DType? source, out long value) =>
        TryReadSceneLightRigRotation(source, "lat", out value);

    internal static bool TryReadSceneLightRigRotationLongitude(A.Scene3DType? source, out long value) =>
        TryReadSceneLightRigRotation(source, "lon", out value);

    internal static bool TryReadSceneLightRigRotationRevolution(A.Scene3DType? source, out long value) =>
        TryReadSceneLightRigRotation(source, "rev", out value);

    internal static bool TryReadSceneCameraRotationLatitude(A.Scene3DType? source, out long value) =>
        TryReadSceneCameraRotation(source, "lat", out value);

    internal static bool TryReadSceneCameraRotationLongitude(A.Scene3DType? source, out long value) =>
        TryReadSceneCameraRotation(source, "lon", out value);

    internal static bool TryReadSceneCameraRotationRevolution(A.Scene3DType? source, out long value) =>
        TryReadSceneCameraRotation(source, "rev", out value);

    private static bool TryReadSceneCameraRotation(
        A.Scene3DType? source,
        string angleName,
        out long value)
    {
        value = 0;
        if (source is null ||
            source.GetAttributes().Any(_ => true) ||
            source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Camera camera ||
            source.ChildElements[1] is not A.LightRig lightRig ||
            camera.ChildElements.Count != 1 ||
            camera.FirstChild is not A.Rotation rotation ||
            lightRig.ChildElements.Count != 0)
            return false;

        var cameraAttributes = camera.GetAttributes().ToArray();
        if (cameraAttributes.Length != 1 || cameraAttributes[0].NamespaceUri.Length != 0 ||
            cameraAttributes[0].LocalName != "prst" ||
            !CameraPresets.Contains(cameraAttributes[0].Value ?? string.Empty))
            return false;

        var lightAttributes = lightRig.GetAttributes().ToArray();
        if (lightAttributes.Length != 2 || lightAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("rig" or "dir")))
            return false;
        var rig = lightAttributes.Where(attribute => attribute.LocalName == "rig").ToArray();
        var direction = lightAttributes.Where(attribute => attribute.LocalName == "dir").ToArray();
        if (rig.Length != 1 || direction.Length != 1 ||
            !LightRigPresets.Contains(rig[0].Value ?? string.Empty) ||
            !LightRigDirections.Contains(direction[0].Value ?? string.Empty))
            return false;

        return TryReadSceneRotationValue(rotation, angleName, out value);
    }

    private static bool TryReadSceneLightRigRotation(
        A.Scene3DType? source,
        string angleName,
        out long value)
    {
        value = 0;
        if (source is null ||
            source.GetAttributes().Any(_ => true) ||
            source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Camera camera ||
            source.ChildElements[1] is not A.LightRig lightRig ||
            camera.ChildElements.Count != 0 ||
            lightRig.ChildElements.Count != 1 ||
            lightRig.FirstChild is not A.Rotation rotation ||
            rotation.ChildElements.Count != 0)
            return false;

        var cameraAttributes = camera.GetAttributes().ToArray();
        if (cameraAttributes.Length != 1 || cameraAttributes[0].NamespaceUri.Length != 0 ||
            cameraAttributes[0].LocalName != "prst" ||
            !CameraPresets.Contains(cameraAttributes[0].Value ?? string.Empty))
            return false;

        var lightAttributes = lightRig.GetAttributes().ToArray();
        if (lightAttributes.Length != 2 || lightAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("rig" or "dir")))
            return false;
        var rig = lightAttributes.Where(attribute => attribute.LocalName == "rig").ToArray();
        var direction = lightAttributes.Where(attribute => attribute.LocalName == "dir").ToArray();
        if (rig.Length != 1 || direction.Length != 1 ||
            !LightRigPresets.Contains(rig[0].Value ?? string.Empty) ||
            !LightRigDirections.Contains(direction[0].Value ?? string.Empty))
            return false;

        return TryReadSceneRotationValue(rotation, angleName, out value);
    }

    private static bool TryReadSceneRotationValue(
        A.Rotation rotation,
        string angleName,
        out long value)
    {
        value = 0;
        var rotationAttributes = rotation.GetAttributes().ToArray();
        if (rotationAttributes.Length != 3 || rotationAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("lat" or "lon" or "rev")))
            return false;
        var angle = rotationAttributes.Where(attribute => attribute.LocalName == angleName).ToArray();
        var longitude = rotationAttributes.Where(attribute => attribute.LocalName == "lon").ToArray();
        var revolution = rotationAttributes.Where(attribute => attribute.LocalName == "rev").ToArray();
        if (angle.Length != 1 || longitude.Length != 1 || revolution.Length != 1 ||
            !TryLiteralSceneRotationAngle(angle[0].Value, out value) ||
            !TryLiteralSceneRotationAngle(longitude[0].Value, out _) ||
            !TryLiteralSceneRotationAngle(revolution[0].Value, out _))
            return false;
        return true;
    }

    private static bool TryReadScenePair(
        A.Scene3DType? source,
        out string cameraPreset,
        out string lightRigPreset,
        out string lightRigDirection,
        out string cameraZoom,
        out string cameraFov)
    {
        cameraPreset = string.Empty;
        lightRigPreset = string.Empty;
        lightRigDirection = string.Empty;
        cameraZoom = string.Empty;
        cameraFov = string.Empty;
        if (source is null ||
            source.GetAttributes().Any(_ => true) ||
            source.ChildElements.Count != 2 ||
            source.ChildElements[0] is not A.Camera camera ||
            source.ChildElements[1] is not A.LightRig lightRig ||
            camera.ChildElements.Count != 0 ||
            lightRig.ChildElements.Count != 0)
            return false;

        var cameraAttributes = camera.GetAttributes().ToArray();
        if (cameraAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("prst" or "zoom" or "fov")))
            return false;
        var preset = cameraAttributes.Where(attribute => attribute.LocalName == "prst").ToArray();
        var zoom = cameraAttributes.Where(attribute => attribute.LocalName == "zoom").ToArray();
        var fov = cameraAttributes.Where(attribute => attribute.LocalName == "fov").ToArray();
        if (preset.Length != 1 || zoom.Length > 1 || fov.Length > 1 ||
            !CameraPresets.Contains(preset[0].Value ?? string.Empty) ||
            zoom is [{ } explicitZoom] &&
            !TryLiteralCoordinate(explicitZoom.Value, allowNegative: false, out _) ||
            fov is [{ } explicitFov] &&
            !TryLiteralCameraFov(explicitFov.Value, out _))
            return false;

        var lightAttributes = lightRig.GetAttributes().ToArray();
        if (lightAttributes.Length != 2 || lightAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("rig" or "dir")))
            return false;
        var rig = lightAttributes.Where(attribute => attribute.LocalName == "rig").ToArray();
        var direction = lightAttributes.Where(attribute => attribute.LocalName == "dir").ToArray();
        if (rig.Length != 1 || direction.Length != 1 ||
            !LightRigPresets.Contains(rig[0].Value ?? string.Empty) ||
            !LightRigDirections.Contains(direction[0].Value ?? string.Empty))
            return false;

        cameraPreset = preset[0].Value ?? string.Empty;
        lightRigPreset = rig[0].Value ?? string.Empty;
        lightRigDirection = direction[0].Value ?? string.Empty;
        cameraZoom = zoom.Length == 1 ? zoom[0].Value ?? string.Empty : string.Empty;
        cameraFov = fov.Length == 1 ? fov[0].Value ?? string.Empty : string.Empty;
        return true;
    }

    private static bool TryReadShape3dRgb(A.Shape3DType? source, string ownerName, out string value)
    {
        value = string.Empty;
        if (source is null || source.ChildElements.Count != 1 ||
            source.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("z" or "extrusionH" or "contourW" or "prstMaterial")))
            return false;

        if (source.FirstChild is not { } colorOwner || colorOwner.NamespaceUri != DrawingNamespace || colorOwner.LocalName != ownerName ||
            colorOwner.HasAttributes || colorOwner.ChildElements.Count != 1)
            return false;
        if (colorOwner.FirstChild is not { } color || color.NamespaceUri != DrawingNamespace || color.LocalName != "srgbClr" ||
            color.ChildElements.Count != 0)
            return false;
        var attributes = color.GetAttributes().ToArray();
        if (attributes.Length != 1 || attributes[0].NamespaceUri.Length != 0 ||
            attributes[0].LocalName != "val" || attributes[0].Value is not { Length: 6 } token ||
            !token.All(Uri.IsHexDigit))
            return false;
        value = token.ToUpperInvariant();
        return true;
    }

    private static bool TryReadShape3dScheme(A.Shape3DType? source, string ownerName, out string value)
    {
        value = string.Empty;
        if (source is null || source.ChildElements.Count != 1 ||
            source.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("z" or "extrusionH" or "contourW" or "prstMaterial")))
            return false;

        if (source.FirstChild is not { } colorOwner || colorOwner.NamespaceUri != DrawingNamespace || colorOwner.LocalName != ownerName ||
            colorOwner.HasAttributes || colorOwner.ChildElements.Count != 1)
            return false;
        if (colorOwner.FirstChild is not A.SchemeColor color || color.ChildElements.Count != 0)
            return false;
        var attributes = color.GetAttributes().ToArray();
        if (attributes.Length != 1 || attributes[0].NamespaceUri.Length != 0 ||
            attributes[0].LocalName != "val" || color.Val?.Value is not { } raw ||
            !PptxColor.TrySchemeToken(raw, out var canonical))
            return false;
        value = canonical;
        return true;
    }

    private static bool TryReadBevelBottomCoordinate(
        A.Shape3DType? source,
        string attributeName,
        out long value)
    {
        value = 0;
        if (!TryReadBevelBottomAttribute(source, attributeName, out var token))
            return false;
        return TryLiteralCoordinate(token, allowNegative: false, out value);
    }

    private static bool TryReadBevelBottomAttribute(
        A.Shape3DType? source,
        string attributeName,
        out string value)
    {
        value = string.Empty;
        if (source is null || source.ChildElements.Count != 1 || source.FirstChild is not A.BevelBottom bevel ||
            source.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("z" or "extrusionH" or "contourW" or "prstMaterial")) ||
            bevel.ChildElements.Count != 0)
            return false;

        var attributes = bevel.GetAttributes().ToArray();
        if (attributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("w" or "h" or "prst")))
            return false;

        var selected = attributes.Where(attribute => attribute.LocalName == attributeName).ToArray();
        if (selected.Length != 1)
            return false;
        value = selected[0].Value ?? string.Empty;
        return true;
    }

    private static bool TryReadBevelTopCoordinate(
        A.Shape3DType? source,
        string attributeName,
        out long value)
    {
        value = 0;
        if (!TryReadBevelTopAttribute(source, attributeName, out var token))
            return false;
        return TryLiteralCoordinate(token, allowNegative: false, out value);
    }

    private static bool TryReadBevelTopAttribute(
        A.Shape3DType? source,
        string attributeName,
        out string value)
    {
        value = string.Empty;
        if (source is null || source.ChildElements.Count != 1 || source.FirstChild is not A.BevelTop bevel ||
            source.GetAttributes().Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("z" or "extrusionH" or "contourW" or "prstMaterial")) ||
            bevel.ChildElements.Count != 0)
            return false;

        var attributes = bevel.GetAttributes().ToArray();
        if (attributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("w" or "h" or "prst")))
            return false;

        var selected = attributes.Where(attribute => attribute.LocalName == attributeName).ToArray();
        if (selected.Length != 1)
            return false;
        value = selected[0].Value ?? string.Empty;
        return true;
    }

    internal static bool IsPresetMaterialToken(string token) => PresetMaterials.Contains(token);

    internal static bool IsSceneCameraPresetToken(string token) => CameraPresets.Contains(token);

    internal static bool IsSceneLightRigPresetToken(string token) => LightRigPresets.Contains(token);

    internal static bool IsSceneLightRigDirectionToken(string token) => LightRigDirections.Contains(token);

    internal static bool IsBevelTopPresetToken(string token) => BevelPresets.Contains(token);

    internal static bool IsBevelBottomPresetToken(string token) => BevelPresets.Contains(token);

    private static bool TryReadCoordinate(
        A.Shape3DType? source,
        string attributeName,
        bool allowNegative,
        out long value)
    {
        value = 0;
        if (!TryReadRawAttribute(source, attributeName, out var token) ||
            !TryLiteralCoordinate(token, allowNegative, out value))
        {
            value = 0;
            return false;
        }

        return true;
    }

    private static bool TryReadRawAttribute(A.Shape3DType? source, string attributeName, out string value)
    {
        value = string.Empty;
        if (source is null || source.ChildElements.Count != 0)
            return false;

        var attributes = source.GetAttributes().ToArray();
        if (attributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("z" or "extrusionH" or "contourW" or "prstMaterial")))
            return false;

        var selected = attributes.Where(attribute => attribute.LocalName == attributeName).ToArray();
        if (selected.Length != 1)
            return false;
        value = selected[0].Value ?? string.Empty;
        return true;
    }

    internal static bool TryRequestedExtrusionHeight(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, out var requested) ||
            !TryReadExtrusionHeight(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedDepth(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: true, out var requested) ||
            !TryReadDepth(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedContourWidth(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadContourWidth(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedPresetMaterial(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!PresetMaterials.Contains(token) ||
            !TryReadPresetMaterial(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedBevelTopWidth(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadBevelTopWidth(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedBevelTopHeight(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadBevelTopHeight(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedBevelTopPreset(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!BevelPresets.Contains(token) ||
            !TryReadBevelTopPreset(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedBevelBottomWidth(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadBevelBottomWidth(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedBevelBottomHeight(
        A.Shape3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadBevelBottomHeight(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedBevelBottomPreset(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!BevelPresets.Contains(token) ||
            !TryReadBevelBottomPreset(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedContourRgb(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        return TryRequestedShape3dRgb(source, "contourClr", token, out value);
    }

    internal static bool TryRequestedExtrusionRgb(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        return TryRequestedShape3dRgb(source, "extrusionClr", token, out value);
    }

    internal static bool TryRequestedContourColorScheme(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        return TryRequestedShape3dScheme(source, "contourClr", token, out value);
    }

    internal static bool TryRequestedExtrusionColorScheme(
        A.Shape3DType? source,
        string token,
        out string value)
    {
        return TryRequestedShape3dScheme(source, "extrusionClr", token, out value);
    }

    internal static bool TryRequestedSceneCameraPreset(
        A.Scene3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!CameraPresets.Contains(token) ||
            !TryReadSceneCameraPreset(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedSceneBackdropAnchorX(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropAnchorCoordinate(source, "x", token, out value);
    }

    internal static bool TryRequestedSceneBackdropAnchorY(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropAnchorCoordinate(source, "y", token, out value);
    }

    internal static bool TryRequestedSceneBackdropAnchorZ(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropAnchorCoordinate(source, "z", token, out value);
    }

    internal static bool TryRequestedSceneBackdropNormalDx(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropNormalCoordinate(source, "dx", token, out value);
    }

    internal static bool TryRequestedSceneBackdropNormalDy(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropNormalCoordinate(source, "dy", token, out value);
    }

    internal static bool TryRequestedSceneBackdropNormalDz(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropNormalCoordinate(source, "dz", token, out value);
    }

    internal static bool TryRequestedSceneBackdropUpDx(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropUpCoordinate(source, "dx", token, out value);
    }

    internal static bool TryRequestedSceneBackdropUpDy(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropUpCoordinate(source, "dy", token, out value);
    }

    internal static bool TryRequestedSceneBackdropUpDz(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        return TryRequestedSceneBackdropUpCoordinate(source, "dz", token, out value);
    }

    internal static bool TryRequestedSceneCameraZoom(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var requested) ||
            !TryReadSceneCameraZoom(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneCameraFov(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralCameraFov(token, out var requested) ||
            !TryReadSceneCameraFov(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneLightRigPreset(
        A.Scene3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!LightRigPresets.Contains(token) ||
            !TryReadSceneLightRigPreset(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedSceneLightRigDirection(
        A.Scene3DType? source,
        string token,
        out string value)
    {
        value = string.Empty;
        if (!LightRigDirections.Contains(token) ||
            !TryReadSceneLightRigDirection(source, out var current) ||
            string.Equals(token, current, StringComparison.Ordinal))
            return false;
        value = token;
        return true;
    }

    internal static bool TryRequestedSceneLightRigRotationLatitude(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneLightRigRotationLatitude(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneLightRigRotationLongitude(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneLightRigRotationLongitude(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneLightRigRotationRevolution(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneLightRigRotationRevolution(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneCameraRotationLatitude(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneCameraRotationLatitude(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneCameraRotationLongitude(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneCameraRotationLongitude(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    internal static bool TryRequestedSceneCameraRotationRevolution(
        A.Scene3DType? source,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralSceneRotationAngle(token, out var requested) ||
            !TryReadSceneCameraRotationRevolution(source, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    private static bool TryRequestedShape3dRgb(
        A.Shape3DType? source,
        string ownerName,
        string token,
        out string value)
    {
        value = string.Empty;
        var requested = token.Trim().TrimStart('#').ToUpperInvariant();
        if (requested.Length != 6 || !requested.All(Uri.IsHexDigit) ||
            !TryReadShape3dRgb(source, ownerName, out var current) || requested == current)
            return false;
        value = requested;
        return true;
    }

    private static bool TryRequestedShape3dScheme(
        A.Shape3DType? source,
        string ownerName,
        string token,
        out string value)
    {
        value = string.Empty;
        string requested;
        try
        {
            requested = PptxColor.NormalizeScheme(token);
        }
        catch (CodecException)
        {
            return false;
        }

        if (!TryReadShape3dScheme(source, ownerName, out var current) ||
            string.Equals(requested, current, StringComparison.Ordinal))
            return false;
        value = requested;
        return true;
    }

    internal static bool TryLiteralCoordinate(string? token, out long value) =>
        TryLiteralCoordinate(token, allowNegative: false, out value);

    internal static bool TryLiteralBackdropCoordinate(string? token, out long value)
    {
        value = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed < MinBackdropCoordinate || parsed > MaxBackdropCoordinate ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed;
        return true;
    }

    internal static bool TryLiteralCameraFov(string? token, out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var parsed) ||
            parsed <= 0 || parsed >= 180 * 60_000)
            return false;
        value = parsed;
        return true;
    }

    internal static bool TryLiteralSceneRotationAngle(string? token, out long value)
    {
        value = 0;
        if (!TryLiteralCoordinate(token, allowNegative: false, out var parsed) ||
            parsed > 360 * 60_000)
            return false;
        value = parsed;
        return true;
    }

    private static bool TryLiteralCoordinate(string? token, bool allowNegative, out long value)
    {
        value = 0;
        if (!long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed < (allowNegative ? int.MinValue : 0) || parsed > MaxCoordinate ||
            parsed.ToString(CultureInfo.InvariantCulture) != token)
            return false;
        value = parsed;
        return true;
    }

    private static bool TryReadSceneBackdropAnchorCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        out long value)
    {
        value = 0;
        if (!TryReadSceneBackdrop(source, out var anchor))
            return false;

        var attributes = anchor.GetAttributes().ToArray();
        var selected = attributes.Where(attribute => attribute.LocalName == coordinateName).ToArray();
        if (selected.Length != 1 || !TryLiteralBackdropCoordinate(selected[0].Value, out value))
        {
            value = 0;
            return false;
        }

        return true;
    }

    private static bool TryReadSceneBackdropNormalCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        out long value)
    {
        value = 0;
        if (!TryReadSceneBackdrop(source, out _, out var normal, out _))
            return false;

        var attributes = normal.GetAttributes().ToArray();
        var selected = attributes.Where(attribute => attribute.LocalName == coordinateName).ToArray();
        if (selected.Length != 1 || !TryLiteralBackdropCoordinate(selected[0].Value, out value))
        {
            value = 0;
            return false;
        }

        return true;
    }

    private static bool TryReadSceneBackdropUpCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        out long value)
    {
        value = 0;
        if (!TryReadSceneBackdrop(source, out _, out _, out var upVector))
            return false;

        var attributes = upVector.GetAttributes().ToArray();
        var selected = attributes.Where(attribute => attribute.LocalName == coordinateName).ToArray();
        if (selected.Length != 1 || !TryLiteralBackdropCoordinate(selected[0].Value, out value))
        {
            value = 0;
            return false;
        }

        return true;
    }

    private static bool TryRequestedSceneBackdropAnchorCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralBackdropCoordinate(token, out var requested) ||
            !TryReadSceneBackdropAnchorCoordinate(source, coordinateName, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    private static bool TryRequestedSceneBackdropNormalCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralBackdropCoordinate(token, out var requested) ||
            !TryReadSceneBackdropNormalCoordinate(source, coordinateName, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    private static bool TryRequestedSceneBackdropUpCoordinate(
        A.Scene3DType? source,
        string coordinateName,
        string token,
        out long value)
    {
        value = 0;
        if (!TryLiteralBackdropCoordinate(token, out var requested) ||
            !TryReadSceneBackdropUpCoordinate(source, coordinateName, out var current) ||
            requested == current)
            return false;
        value = requested;
        return true;
    }

    private static bool TryReadSceneBackdrop(A.Scene3DType? source, out A.Anchor anchor)
    {
        return TryReadSceneBackdrop(source, out anchor, out _, out _);
    }

    private static bool TryReadSceneBackdrop(
        A.Scene3DType? source,
        out A.Anchor anchor,
        out A.Normal backdropNormal,
        out A.UpVector backdropUpVector)
    {
        anchor = null!;
        backdropNormal = null!;
        backdropUpVector = null!;
        if (source is null ||
            source.GetAttributes().Any(_ => true) ||
            source.ChildElements.Count != 3 ||
            source.ChildElements[0] is not A.Camera camera ||
            source.ChildElements[1] is not A.LightRig lightRig ||
            source.ChildElements[2] is not A.Backdrop backdrop ||
            camera.ChildElements.Count != 0 ||
            lightRig.ChildElements.Count != 0 ||
            backdrop.GetAttributes().Any(_ => true) ||
            backdrop.ChildElements.Count != 3 ||
            backdrop.ChildElements[0] is not A.Anchor backdropAnchor ||
            backdrop.ChildElements[1] is not A.Normal normal ||
            backdrop.ChildElements[2] is not A.UpVector upVector)
            return false;

        var cameraAttributes = camera.GetAttributes().ToArray();
        if (cameraAttributes.Length != 1 || cameraAttributes[0].NamespaceUri.Length != 0 ||
            cameraAttributes[0].LocalName != "prst" ||
            !CameraPresets.Contains(cameraAttributes[0].Value ?? string.Empty))
            return false;

        var lightAttributes = lightRig.GetAttributes().ToArray();
        if (lightAttributes.Length != 2 || lightAttributes.Any(attribute =>
                attribute.NamespaceUri.Length != 0 ||
                attribute.LocalName is not ("rig" or "dir")))
            return false;
        var rig = lightAttributes.Where(attribute => attribute.LocalName == "rig").ToArray();
        var direction = lightAttributes.Where(attribute => attribute.LocalName == "dir").ToArray();
        if (rig.Length != 1 || direction.Length != 1 ||
            !LightRigPresets.Contains(rig[0].Value ?? string.Empty) ||
            !LightRigDirections.Contains(direction[0].Value ?? string.Empty))
            return false;

        if (!HasBackdropCoordinateAttributes(backdropAnchor, "x", "y", "z") ||
            !HasBackdropCoordinateAttributes(normal, "dx", "dy", "dz") ||
            !HasBackdropCoordinateAttributes(upVector, "dx", "dy", "dz"))
            return false;

        anchor = backdropAnchor;
        backdropNormal = normal;
        backdropUpVector = upVector;
        return true;
    }

    private static bool HasBackdropCoordinateAttributes(OpenXmlElement element, params string[] names)
    {
        var attributes = element.GetAttributes().ToArray();
        return attributes.Length == names.Length &&
            attributes.All(attribute => attribute.NamespaceUri.Length == 0) &&
            names.All(name => attributes.Count(attribute => attribute.LocalName == name) == 1) &&
            attributes.All(attribute => TryLiteralBackdropCoordinate(attribute.Value, out _));
    }
}
