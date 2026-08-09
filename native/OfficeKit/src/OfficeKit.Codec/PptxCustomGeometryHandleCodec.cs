using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Owns the bounded a:ahLst protocol. The parent custom-geometry codec decides
// package topology; this leaf owns handle grammar, evaluated bounds, wire
// presence, native serialization, and source-bound handle identity.
internal static class PptxCustomGeometryHandleCodec
{
    private const int MaxHandles = 1_024;
    private const long MaxCoordinate = int.MaxValue;
    private const int FullTurnAngle = 21_600_000;

    private readonly record struct ParsedValue(bool Present, bool IsLiteral, long Literal, string Reference, double Resolved);

    internal static bool TryRead(
        A.AdjustHandleList? source,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        long widthEmu,
        long heightEmu,
        out IReadOnlyList<PresentationCustomGeometryAdjustmentHandle> handles)
    {
        handles = [];
        if (source is null) return true;
        if (source.HasAttributes || source.ChildElements.Count > MaxHandles) return false;
        var result = new List<PresentationCustomGeometryAdjustmentHandle>(source.ChildElements.Count);
        foreach (var nativeHandle in source.ChildElements)
        {
            PresentationCustomGeometryAdjustmentHandle? handle = nativeHandle switch
            {
                A.AdjustHandleXY xy when TryRead(xy, formulas, widthEmu, heightEmu, out var modeled) =>
                    new PresentationCustomGeometryAdjustmentHandle { Xy = modeled },
                A.AdjustHandlePolar polar when TryRead(polar, formulas, widthEmu, heightEmu, out var modeled) =>
                    new PresentationCustomGeometryAdjustmentHandle { Polar = modeled },
                _ => null,
            };
            if (handle is null) return false;
            result.Add(handle);
        }
        handles = result;
        return true;
    }

    internal static void Validate(
        PresentationShape shape,
        string shapeId,
        PptxCustomGeometryFormulaCodec.Graph formulas)
    {
        if (shape.CustomAdjustmentHandles.Count > MaxHandles)
            throw new CodecException("presentation_item_budget_exceeded", $"Presentation shape {shapeId} custom geometry exceeds the {MaxHandles}-adjustment-handle budget.");
        foreach (var handle in shape.CustomAdjustmentHandles)
        {
            var valid = handle.HandleCase switch
            {
                PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy => Validate(handle.Xy, shape, formulas),
                PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar => Validate(handle.Polar, shape, formulas),
                _ => false,
            };
            if (!valid)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid bounded custom adjustment handle, range, position, or formula reference.");
        }
    }

    internal static A.AdjustHandleList Build(IEnumerable<PresentationCustomGeometryAdjustmentHandle> handles) =>
        new(handles.Select(Write));

    internal static bool TopologyEquals(
        IReadOnlyList<PresentationCustomGeometryAdjustmentHandle> left,
        IEnumerable<PresentationCustomGeometryAdjustmentHandle> rightSource)
    {
        var right = rightSource.ToArray();
        if (left.Count != right.Length) return false;
        for (var index = 0; index < left.Count; index++)
        {
            if (left[index].HandleCase != right[index].HandleCase) return false;
            switch (left[index].HandleCase)
            {
                case PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy:
                    if (left[index].Xy.XAdjustment != right[index].Xy.XAdjustment ||
                        left[index].Xy.YAdjustment != right[index].Xy.YAdjustment)
                        return false;
                    break;
                case PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar:
                    if (left[index].Polar.RadialAdjustment != right[index].Polar.RadialAdjustment ||
                        left[index].Polar.AngleAdjustment != right[index].Polar.AngleAdjustment)
                        return false;
                    break;
                default:
                    return false;
            }
        }
        return true;
    }

    private static bool TryRead(
        A.AdjustHandleXY source,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        long widthEmu,
        long heightEmu,
        out PresentationCustomGeometryXyAdjustmentHandle handle)
    {
        handle = null!;
        if (!HasOnlyAttributes(source, "gdRefX", "minX", "maxX", "gdRefY", "minY", "maxY") ||
            source.ChildElements.Count != 1 || source.ChildElements[0] is not A.Position position ||
            !SupportsPosition(position, formulas, widthEmu, heightEmu) ||
            !TryReadRange(source.XAdjustmentGuide?.Value, source.MinX?.Value, source.MaxX?.Value, formulas, false, false, out var minX, out var maxX) ||
            !TryReadRange(source.YAdjustmentGuide?.Value, source.MinY?.Value, source.MaxY?.Value, formulas, false, false, out var minY, out var maxY) ||
            (string.IsNullOrEmpty(source.XAdjustmentGuide?.Value) && string.IsNullOrEmpty(source.YAdjustmentGuide?.Value)))
            return false;
        var result = new PresentationCustomGeometryXyAdjustmentHandle
        {
            XAdjustment = source.XAdjustmentGuide?.Value ?? string.Empty,
            YAdjustment = source.YAdjustmentGuide?.Value ?? string.Empty,
            Position = ReadPosition(position),
        };
        SetRange(result, minX, maxX, true);
        SetRange(result, minY, maxY, false);
        handle = result;
        return true;
    }

    private static bool TryRead(
        A.AdjustHandlePolar source,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        long widthEmu,
        long heightEmu,
        out PresentationCustomGeometryPolarAdjustmentHandle handle)
    {
        handle = null!;
        if (!HasOnlyAttributes(source, "gdRefR", "minR", "maxR", "gdRefAng", "minAng", "maxAng") ||
            source.ChildElements.Count != 1 || source.ChildElements[0] is not A.Position position ||
            !SupportsPosition(position, formulas, widthEmu, heightEmu) ||
            !TryReadRange(source.RadialAdjustmentGuide?.Value, source.MinRadial?.Value, source.MaxRadial?.Value, formulas, false, true, out var minRadius, out var maxRadius) ||
            !TryReadRange(source.AngleAdjustmentGuide?.Value, source.MinAngle?.Value, source.MaxAngle?.Value, formulas, true, false, out var minAngle, out var maxAngle) ||
            (string.IsNullOrEmpty(source.RadialAdjustmentGuide?.Value) && string.IsNullOrEmpty(source.AngleAdjustmentGuide?.Value)))
            return false;
        var result = new PresentationCustomGeometryPolarAdjustmentHandle
        {
            RadialAdjustment = source.RadialAdjustmentGuide?.Value ?? string.Empty,
            AngleAdjustment = source.AngleAdjustmentGuide?.Value ?? string.Empty,
            Position = ReadPosition(position),
        };
        SetRange(result, minRadius, maxRadius, true);
        SetRange(result, minAngle, maxAngle, false);
        handle = result;
        return true;
    }

    private static bool TryReadRange(
        string? adjustment,
        string? minimum,
        string? maximum,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        bool angle,
        bool nonNegative,
        out ParsedValue parsedMinimum,
        out ParsedValue parsedMaximum)
    {
        parsedMinimum = default;
        parsedMaximum = default;
        var hasMinimum = minimum is not null;
        var hasMaximum = maximum is not null;
        if (string.IsNullOrEmpty(adjustment)) return !hasMinimum && !hasMaximum;
        if (!formulas.Adjustments.Any(item => item.Name == adjustment) ||
            !formulas.TryResolveReference(adjustment, out var current))
            return false;
        if ((angle && Math.Abs(current) > FullTurnAngle) || (nonNegative && current < 0) || hasMinimum != hasMaximum)
            return false;
        if (!hasMinimum) return true;
        if (!TryReadValue(minimum, formulas, angle, out parsedMinimum) ||
            !TryReadValue(maximum, formulas, angle, out parsedMaximum) ||
            (nonNegative && (parsedMinimum.Resolved < 0 || parsedMaximum.Resolved < 0)) ||
            parsedMinimum.Resolved > parsedMaximum.Resolved ||
            current < parsedMinimum.Resolved || current > parsedMaximum.Resolved)
            return false;
        return true;
    }

    private static bool TryReadValue(
        string? source,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        bool angle,
        out ParsedValue value)
    {
        value = default;
        if (source is null) return false;
        if (angle && TryAngle(source, out var angleLiteral))
        {
            if (Math.Abs(angleLiteral) > FullTurnAngle) return false;
            value = new ParsedValue(true, true, angleLiteral, string.Empty, angleLiteral);
            return true;
        }
        if (!angle && TryCoordinate(source, out var coordinateLiteral))
        {
            value = new ParsedValue(true, true, coordinateLiteral, string.Empty, coordinateLiteral);
            return true;
        }
        if (!formulas.TryResolveReference(source, out var resolved) ||
            (angle && Math.Abs(resolved) > FullTurnAngle))
            return false;
        value = new ParsedValue(true, false, 0, source, resolved);
        return true;
    }

    private static bool Validate(
        PresentationCustomGeometryXyAdjustmentHandle handle,
        PresentationShape shape,
        PptxCustomGeometryFormulaCodec.Graph formulas) =>
        handle.Position is not null && ValidatePosition(handle.Position, shape, formulas) &&
        ValidateRange(handle.XAdjustment,
            handle.HasMinX, handle.MinX, handle.HasMinXReference, handle.MinXReference,
            handle.HasMaxX, handle.MaxX, handle.HasMaxXReference, handle.MaxXReference,
            formulas, false, false) &&
        ValidateRange(handle.YAdjustment,
            handle.HasMinY, handle.MinY, handle.HasMinYReference, handle.MinYReference,
            handle.HasMaxY, handle.MaxY, handle.HasMaxYReference, handle.MaxYReference,
            formulas, false, false) &&
        (!string.IsNullOrEmpty(handle.XAdjustment) || !string.IsNullOrEmpty(handle.YAdjustment));

    private static bool Validate(
        PresentationCustomGeometryPolarAdjustmentHandle handle,
        PresentationShape shape,
        PptxCustomGeometryFormulaCodec.Graph formulas) =>
        handle.Position is not null && ValidatePosition(handle.Position, shape, formulas) &&
        ValidateRange(handle.RadialAdjustment,
            handle.HasMinRadius, handle.MinRadius, handle.HasMinRadiusReference, handle.MinRadiusReference,
            handle.HasMaxRadius, handle.MaxRadius, handle.HasMaxRadiusReference, handle.MaxRadiusReference,
            formulas, false, true) &&
        ValidateRange(handle.AngleAdjustment,
            handle.HasMinAngle60000, handle.MinAngle60000, handle.HasMinAngleReference, handle.MinAngleReference,
            handle.HasMaxAngle60000, handle.MaxAngle60000, handle.HasMaxAngleReference, handle.MaxAngleReference,
            formulas, true, false) &&
        (!string.IsNullOrEmpty(handle.RadialAdjustment) || !string.IsNullOrEmpty(handle.AngleAdjustment));

    private static bool ValidateRange(
        string adjustment,
        bool hasMinimumLiteral,
        long minimumLiteral,
        bool hasMinimumReference,
        string minimumReference,
        bool hasMaximumLiteral,
        long maximumLiteral,
        bool hasMaximumReference,
        string maximumReference,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        bool angle,
        bool nonNegative)
    {
        var hasMinimum = hasMinimumLiteral || hasMinimumReference;
        var hasMaximum = hasMaximumLiteral || hasMaximumReference;
        if (string.IsNullOrEmpty(adjustment)) return !hasMinimum && !hasMaximum;
        if (!formulas.Adjustments.Any(item => item.Name == adjustment) ||
            !formulas.TryResolveReference(adjustment, out var current))
            return false;
        if ((angle && Math.Abs(current) > FullTurnAngle) || (nonNegative && current < 0) || hasMinimum != hasMaximum)
            return false;
        if (!hasMinimum) return true;
        return TryWireValue(hasMinimumLiteral, minimumLiteral, hasMinimumReference, minimumReference, formulas, angle, out var minimum) &&
            TryWireValue(hasMaximumLiteral, maximumLiteral, hasMaximumReference, maximumReference, formulas, angle, out var maximum) &&
            (!nonNegative || (minimum >= 0 && maximum >= 0)) &&
            minimum <= maximum && current >= minimum && current <= maximum;
    }

    private static bool TryWireValue(
        bool hasLiteral,
        long literal,
        bool hasReference,
        string reference,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        bool angle,
        out double value)
    {
        value = 0;
        if (hasLiteral == hasReference) return false;
        if (hasLiteral)
        {
            value = literal;
            return !angle || Math.Abs(value) <= FullTurnAngle;
        }
        return formulas.TryResolveReference(reference, out value) && (!angle || Math.Abs(value) <= FullTurnAngle);
    }

    private static bool SupportsPosition(
        A.Position position,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        long widthEmu,
        long heightEmu) =>
        HasNoInnerXml(position) && HasOnlyAttributes(position, "x", "y") &&
        TryNativeValue(position.X?.Value, formulas, out var x) && x >= 0 && x <= widthEmu &&
        TryNativeValue(position.Y?.Value, formulas, out var y) && y >= 0 && y <= heightEmu;

    private static bool ValidatePosition(
        PresentationCustomGeometryPoint point,
        PresentationShape shape,
        PptxCustomGeometryFormulaCodec.Graph formulas) =>
        TryWirePointValue(point.HasXReference, point.XReference, point.X, formulas, out var x) && x >= 0 && x <= shape.WidthEmu &&
        TryWirePointValue(point.HasYReference, point.YReference, point.Y, formulas, out var y) && y >= 0 && y <= shape.HeightEmu;

    private static bool TryWirePointValue(
        bool hasReference,
        string reference,
        long literal,
        PptxCustomGeometryFormulaCodec.Graph formulas,
        out double value)
    {
        value = 0;
        if (hasReference && literal != 0) return false;
        return formulas.TryResolve(hasReference ? reference : null, literal, out value);
    }

    private static bool TryNativeValue(string? source, PptxCustomGeometryFormulaCodec.Graph formulas, out double value)
    {
        if (TryCoordinate(source, out var coordinate))
        {
            value = coordinate;
            return true;
        }
        return formulas.TryResolveReference(source, out value);
    }

    private static PresentationCustomGeometryPoint ReadPosition(A.Position position)
    {
        var result = new PresentationCustomGeometryPoint();
        if (TryCoordinate(position.X?.Value, out var x)) result.X = x;
        else result.XReference = position.X!.Value!;
        if (TryCoordinate(position.Y?.Value, out var y)) result.Y = y;
        else result.YReference = position.Y!.Value!;
        return result;
    }

    private static void SetRange(PresentationCustomGeometryXyAdjustmentHandle target, ParsedValue minimum, ParsedValue maximum, bool x)
    {
        if (!minimum.Present) return;
        if (x)
        {
            if (minimum.IsLiteral) target.MinX = minimum.Literal; else target.MinXReference = minimum.Reference;
            if (maximum.IsLiteral) target.MaxX = maximum.Literal; else target.MaxXReference = maximum.Reference;
        }
        else
        {
            if (minimum.IsLiteral) target.MinY = minimum.Literal; else target.MinYReference = minimum.Reference;
            if (maximum.IsLiteral) target.MaxY = maximum.Literal; else target.MaxYReference = maximum.Reference;
        }
    }

    private static void SetRange(PresentationCustomGeometryPolarAdjustmentHandle target, ParsedValue minimum, ParsedValue maximum, bool radial)
    {
        if (!minimum.Present) return;
        if (radial)
        {
            if (minimum.IsLiteral) target.MinRadius = minimum.Literal; else target.MinRadiusReference = minimum.Reference;
            if (maximum.IsLiteral) target.MaxRadius = maximum.Literal; else target.MaxRadiusReference = maximum.Reference;
        }
        else
        {
            if (minimum.IsLiteral) target.MinAngle60000 = checked((int)minimum.Literal); else target.MinAngleReference = minimum.Reference;
            if (maximum.IsLiteral) target.MaxAngle60000 = checked((int)maximum.Literal); else target.MaxAngleReference = maximum.Reference;
        }
    }

    private static OpenXmlElement Write(PresentationCustomGeometryAdjustmentHandle source) => source.HandleCase switch
    {
        PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy => Write(source.Xy),
        PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar => Write(source.Polar),
        _ => throw new CodecException("invalid_presentation_geometry", "Presentation custom geometry contains an empty adjustment handle."),
    };

    private static A.AdjustHandleXY Write(PresentationCustomGeometryXyAdjustmentHandle source)
    {
        var handle = new A.AdjustHandleXY(Position(source.Position));
        if (!string.IsNullOrEmpty(source.XAdjustment)) handle.XAdjustmentGuide = source.XAdjustment;
        if (source.HasMinX) handle.MinX = source.MinX.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMinXReference) handle.MinX = source.MinXReference;
        if (source.HasMaxX) handle.MaxX = source.MaxX.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMaxXReference) handle.MaxX = source.MaxXReference;
        if (!string.IsNullOrEmpty(source.YAdjustment)) handle.YAdjustmentGuide = source.YAdjustment;
        if (source.HasMinY) handle.MinY = source.MinY.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMinYReference) handle.MinY = source.MinYReference;
        if (source.HasMaxY) handle.MaxY = source.MaxY.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMaxYReference) handle.MaxY = source.MaxYReference;
        return handle;
    }

    private static A.AdjustHandlePolar Write(PresentationCustomGeometryPolarAdjustmentHandle source)
    {
        var handle = new A.AdjustHandlePolar(Position(source.Position));
        if (!string.IsNullOrEmpty(source.RadialAdjustment)) handle.RadialAdjustmentGuide = source.RadialAdjustment;
        if (source.HasMinRadius) handle.MinRadial = source.MinRadius.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMinRadiusReference) handle.MinRadial = source.MinRadiusReference;
        if (source.HasMaxRadius) handle.MaxRadial = source.MaxRadius.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMaxRadiusReference) handle.MaxRadial = source.MaxRadiusReference;
        if (!string.IsNullOrEmpty(source.AngleAdjustment)) handle.AngleAdjustmentGuide = source.AngleAdjustment;
        if (source.HasMinAngle60000) handle.MinAngle = source.MinAngle60000.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMinAngleReference) handle.MinAngle = source.MinAngleReference;
        if (source.HasMaxAngle60000) handle.MaxAngle = source.MaxAngle60000.ToString(CultureInfo.InvariantCulture);
        else if (source.HasMaxAngleReference) handle.MaxAngle = source.MaxAngleReference;
        return handle;
    }

    private static A.Position Position(PresentationCustomGeometryPoint source) => new()
    {
        X = source.HasXReference ? source.XReference : source.X.ToString(CultureInfo.InvariantCulture),
        Y = source.HasYReference ? source.YReference : source.Y.ToString(CultureInfo.InvariantCulture),
    };

    private static bool TryCoordinate(string? value, out long coordinate) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out coordinate) &&
        coordinate >= -MaxCoordinate && coordinate <= MaxCoordinate;

    private static bool TryAngle(string? value, out int angle) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out angle);

    private static bool HasNoInnerXml(OpenXmlElement element) => string.IsNullOrEmpty(element.InnerXml);

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => string.IsNullOrEmpty(attribute.NamespaceUri) && allowed.Contains(attribute.LocalName));
    }
}
