namespace OfficeKit.Codec;

// PPJ frame anchors are layout references, not geometry connection-site indexes.
// Resolve in slide space so nested groups and both-auto endpoints use one metric.
internal sealed class PpjConnectorEndpointResolver
{
    internal readonly record struct Endpoints(long StartX, long StartY, long EndX, long EndY);
    private readonly record struct Point(double X, double Y);
    private sealed record Entry(PpjElementModel Element, Matrix Parent);
    private readonly Dictionary<string, Entry> _entries = new(StringComparer.Ordinal);
    private readonly IReadOnlyDictionary<string, PpjFrameModel> _frames;
    private readonly IReadOnlyDictionary<string, PpjFrameModel> _childFrames;
    private static readonly Point[] Sides = [new(.5, 0), new(1, .5), new(.5, 1), new(0, .5)];

    internal PpjConnectorEndpointResolver(IReadOnlyList<PpjElementModel> elements,
        IReadOnlyDictionary<string, PpjFrameModel>? frames = null,
        IReadOnlyDictionary<string, PpjFrameModel>? childFrames = null)
    {
        _frames = frames ?? new Dictionary<string, PpjFrameModel>();
        _childFrames = childFrames ?? new Dictionary<string, PpjFrameModel>();
        Index(elements, Matrix.Identity);
    }

    private void Index(IReadOnlyList<PpjElementModel> elements, Matrix parent)
    {
        foreach (var element in elements)
        {
            if (!_entries.TryAdd(element.Id, new(element, parent)))
                throw Invalid(element.Id, "duplicate target ID");
            if (element is not PpjGroupElementModel group) continue;
            var child = _childFrames.GetValueOrDefault(group.Id, group.ChildFrame);
            RequireFrame(child, element.Id);
            var normalize = new Matrix(1 / child.Width, 0, 0, 1 / child.Height,
                -child.X / child.Width, -child.Y / child.Height);
            Index(group.Elements, parent.Then(Frame(_frames.GetValueOrDefault(group.Id, group.Frame), group.Id)).Then(normalize));
        }
    }

    internal Endpoints Resolve(PpjConnectorElementModel connector)
    {
        if (!_entries.TryGetValue(connector.Id, out var entry))
            throw Invalid(connector.Id, "connector is absent from the resolved page");
        var start = Candidates(connector.From, entry.Parent, connector.Id, "from");
        var end = Candidates(connector.To, entry.Parent, connector.Id, "to");
        var selectedStart = start[0];
        var selectedEnd = end[0];
        var best = double.PositiveInfinity;
        foreach (var from in start)
        foreach (var to in end)
        {
            var dx = from.X - to.X;
            var dy = from.Y - to.Y;
            var distance = dx * dx + dy * dy;
            if (!double.IsFinite(distance)) throw Invalid(connector.Id, "endpoint distance is not finite");
            if (distance >= best) continue;
            best = distance;
            selectedStart = from;
            selectedEnd = to;
        }
        var inverse = entry.Parent.Inverse(connector.Id);
        return Coordinates(
            connector.From.ElementId is null ? Literal(connector.From, connector.Id) : inverse.Apply(selectedStart),
            connector.To.ElementId is null ? Literal(connector.To, connector.Id) : inverse.Apply(selectedEnd), connector.Id);
    }

    // Source overlays have no page target context until their candidate page is
    // resolved. Never silently use the connector frame in that path either.
    internal static Endpoints ResolveLiteral(PpjConnectorElementModel connector)
    {
        if (connector.From.ElementId is not null || connector.To.ElementId is not null)
            throw Invalid(connector.Id, "object endpoints require a resolved page target context");
        return Coordinates(Literal(connector.From, connector.Id), Literal(connector.To, connector.Id), connector.Id);
    }

    private Point[] Candidates(PpjConnectorEndpointModel endpoint, Matrix parent, string id, string field)
    {
        if (endpoint.ElementId is null) return [parent.Apply(Literal(endpoint, id))];
        if (!_entries.TryGetValue(endpoint.ElementId, out var target))
            throw Invalid(id, $"{field}.element target {endpoint.ElementId} is absent from the expanded page");
        if (target.Element is PpjConnectorElementModel or PpjComponentElementModel or PpjSlotElementModel or PpjOpaqueElementModel)
            throw Invalid(id, $"{field}.element target {endpoint.ElementId} has unsupported type {target.Element.Type}");
        var transform = target.Parent.Then(Frame(_frames.GetValueOrDefault(target.Element.Id, target.Element.Frame), target.Element.Id));
        Point[] points = endpoint.Anchor switch
        {
            "auto" => Sides,
            "top" => [Sides[0]],
            "right" => [Sides[1]],
            "bottom" => [Sides[2]],
            "left" => [Sides[3]],
            "center" => [new(.5, .5)],
            _ => throw Invalid(id, $"{field}.anchor is not supported"),
        };
        return points.Select(transform.Apply).ToArray();
    }

    private static Point Literal(PpjConnectorEndpointModel endpoint, string id) =>
        endpoint.X is { } x && endpoint.Y is { } y && double.IsFinite(x) && double.IsFinite(y)
            ? new(x, y) : throw Invalid(id, "literal endpoint requires finite x and y");

    private static Endpoints Coordinates(Point start, Point end, string id) =>
        new(Emu(start.X, id), Emu(start.Y, id), Emu(end.X, id), Emu(end.Y, id));

    private static long Emu(double value, string id)
    {
        var rounded = Math.Round(value * 12_700, MidpointRounding.ToEven);
        // The native connector profile requires nonnegative local coordinates.
        // Small round-off around zero is settled by EMU quantization, not clamp.
        if (!double.IsFinite(rounded) || rounded < 0 || rounded >= long.MaxValue)
            throw Invalid(id, "resolved endpoint is outside the native coordinate profile");
        return checked((long)rounded);
    }

    private static void RequireFrame(PpjFrameModel frame, string id)
    {
        if (!double.IsFinite(frame.X) || !double.IsFinite(frame.Y) ||
            !double.IsFinite(frame.Width) || !double.IsFinite(frame.Height) ||
            !double.IsFinite(frame.Rotation) || frame.Width <= 0 || frame.Height <= 0)
            throw Invalid(id, "target or group frame is not finite and invertible");
    }

    private static Matrix Frame(PpjFrameModel frame, string id)
    {
        RequireFrame(frame, id);
        var radians = frame.Rotation * Math.PI / 180;
        var cos = Math.Cos(radians);
        var sin = Math.Sin(radians);
        var sx = frame.FlipH ? -frame.Width : frame.Width;
        var sy = frame.FlipV ? -frame.Height : frame.Height;
        var a = cos * sx;
        var b = sin * sx;
        var c = -sin * sy;
        var d = cos * sy;
        return new(a, b, c, d, frame.X + frame.Width / 2 - (a + c) / 2,
            frame.Y + frame.Height / 2 - (b + d) / 2);
    }

    private static CodecException Invalid(string id, string reason) =>
        new("ppj.connector.endpoint", $"Connector/object {id}: {reason}.");

    private readonly record struct Matrix(double A, double B, double C, double D, double X, double Y)
    {
        internal static Matrix Identity => new(1, 0, 0, 1, 0, 0);
        internal Point Apply(Point point) => new(A * point.X + C * point.Y + X, B * point.X + D * point.Y + Y);
        // this(inner(point)): preserve the outer-to-inner ancestry order.
        internal Matrix Then(Matrix inner) => new(
            A * inner.A + C * inner.B, B * inner.A + D * inner.B,
            A * inner.C + C * inner.D, B * inner.C + D * inner.D,
            A * inner.X + C * inner.Y + X, B * inner.X + D * inner.Y + Y);
        internal Matrix Inverse(string id)
        {
            var determinant = A * D - B * C;
            if (!double.IsFinite(determinant) || determinant == 0)
                throw Invalid(id, "connector parent transform is not invertible");
            return new(D / determinant, -B / determinant, -C / determinant, A / determinant,
                (C * Y - D * X) / determinant, (B * X - A * Y) / determinant);
        }
    }
}
