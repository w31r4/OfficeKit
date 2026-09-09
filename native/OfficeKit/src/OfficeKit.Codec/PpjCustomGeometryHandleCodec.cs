using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// PPJ local-point/degree values map to the existing native handle graph.
internal static class PpjCustomGeometryHandleCodec
{
    internal static PresentationCustomGeometryAdjustmentHandle Read(JsonElement source)
    {
        var position = new PresentationCustomGeometryPoint();
        ReadValue(source.GetProperty("position"), "x", v => position.X = v, v => position.XReference = v);
        ReadValue(source.GetProperty("position"), "y", v => position.Y = v, v => position.YReference = v);
        if (source.GetProperty("kind").GetString() == "xy")
        {
            var target = new PresentationCustomGeometryXyAdjustmentHandle { Position = position };
            if (source.TryGetProperty("xAdjustment", out var xAdjustment)) target.XAdjustment = xAdjustment.GetString()!;
            if (source.TryGetProperty("yAdjustment", out var yAdjustment)) target.YAdjustment = yAdjustment.GetString()!;
            ReadValue(source, "minX", v => target.MinX = v, v => target.MinXReference = v);
            ReadValue(source, "maxX", v => target.MaxX = v, v => target.MaxXReference = v);
            ReadValue(source, "minY", v => target.MinY = v, v => target.MinYReference = v);
            ReadValue(source, "maxY", v => target.MaxY = v, v => target.MaxYReference = v);
            return new PresentationCustomGeometryAdjustmentHandle { Xy = target };
        }
        if (source.GetProperty("kind").GetString() == "polar")
        {
            var target = new PresentationCustomGeometryPolarAdjustmentHandle { Position = position };
            if (source.TryGetProperty("radialAdjustment", out var radialAdjustment)) target.RadialAdjustment = radialAdjustment.GetString()!;
            if (source.TryGetProperty("angleAdjustment", out var angleAdjustment)) target.AngleAdjustment = angleAdjustment.GetString()!;
            ReadValue(source, "minRadius", v => target.MinRadius = v, v => target.MinRadiusReference = v);
            ReadValue(source, "maxRadius", v => target.MaxRadius = v, v => target.MaxRadiusReference = v);
            ReadValue(source, "minAngle", v => target.MinAngle60000 = checked((int)v), v => target.MinAngleReference = v, angle: true);
            ReadValue(source, "maxAngle", v => target.MaxAngle60000 = checked((int)v), v => target.MaxAngleReference = v, angle: true);
            return new PresentationCustomGeometryAdjustmentHandle { Polar = target };
        }
        throw new CodecException("invalid_presentation_geometry", "Unknown PPJ adjustment handle kind.");
    }

    internal static JsonObject Project(PresentationCustomGeometryAdjustmentHandle source)
    {
        var result = new JsonObject();
        if (source.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Xy)
        {
            var handle = source.Xy;
            result["kind"] = "xy";
            result["position"] = Position(handle.Position);
            if (handle.XAdjustment.Length > 0) result["xAdjustment"] = handle.XAdjustment;
            if (handle.YAdjustment.Length > 0) result["yAdjustment"] = handle.YAdjustment;
            PutValue(result, "minX", handle.HasMinX, handle.MinX, handle.HasMinXReference, handle.MinXReference);
            PutValue(result, "maxX", handle.HasMaxX, handle.MaxX, handle.HasMaxXReference, handle.MaxXReference);
            PutValue(result, "minY", handle.HasMinY, handle.MinY, handle.HasMinYReference, handle.MinYReference);
            PutValue(result, "maxY", handle.HasMaxY, handle.MaxY, handle.HasMaxYReference, handle.MaxYReference);
        }
        else if (source.HandleCase == PresentationCustomGeometryAdjustmentHandle.HandleOneofCase.Polar)
        {
            var handle = source.Polar;
            result["kind"] = "polar";
            result["position"] = Position(handle.Position);
            if (handle.RadialAdjustment.Length > 0) result["radialAdjustment"] = handle.RadialAdjustment;
            if (handle.AngleAdjustment.Length > 0) result["angleAdjustment"] = handle.AngleAdjustment;
            PutValue(result, "minRadius", handle.HasMinRadius, handle.MinRadius, handle.HasMinRadiusReference, handle.MinRadiusReference);
            PutValue(result, "maxRadius", handle.HasMaxRadius, handle.MaxRadius, handle.HasMaxRadiusReference, handle.MaxRadiusReference);
            PutValue(result, "minAngle", handle.HasMinAngle60000, handle.MinAngle60000, handle.HasMinAngleReference, handle.MinAngleReference, angle: true);
            PutValue(result, "maxAngle", handle.HasMaxAngle60000, handle.MaxAngle60000, handle.HasMaxAngleReference, handle.MaxAngleReference, angle: true);
        }
        else throw new CodecException("invalid_presentation_geometry", "Unknown native adjustment handle kind.");
        return result;
    }

    private static JsonObject Position(PresentationCustomGeometryPoint point) => new()
    {
        ["x"] = point.HasXReference ? JsonValue.Create(point.XReference) : JsonValue.Create(point.X / 12_700d),
        ["y"] = point.HasYReference ? JsonValue.Create(point.YReference) : JsonValue.Create(point.Y / 12_700d),
    };

    private static void ReadValue(JsonElement source, string name, Action<long> literal, Action<string> reference, bool angle = false)
    {
        if (!source.TryGetProperty(name, out var value)) return;
        if (value.ValueKind == JsonValueKind.String) reference(value.GetString()!);
        else literal(checked((long)Math.Round(value.GetDouble() * (angle ? 60_000d : 12_700d))));
    }

    private static void PutValue(JsonObject target, string name, bool present, long literal, bool hasReference, string reference, bool angle = false)
    {
        if (hasReference) target[name] = reference;
        else if (present) target[name] = literal / (angle ? 60_000d : 12_700d);
    }
}
