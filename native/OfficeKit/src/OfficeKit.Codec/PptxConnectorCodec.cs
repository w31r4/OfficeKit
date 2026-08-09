using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal static class PptxConnectorCodec
{
    private const int RotationUnitsPerDegree = 60_000;
    private static readonly IReadOnlySet<string> ConnectorTypes = new HashSet<string>(StringComparer.Ordinal) { "straight", "elbow", "curved" };
    private static readonly IReadOnlySet<string> LineStyles = new HashSet<string>(StringComparer.Ordinal) { "", "solid", "dashed", "none" };
    private static readonly IReadOnlySet<string> ArrowTypes = new HashSet<string>(StringComparer.Ordinal) { "", "triangle", "stealth", "diamond", "oval", "arrow" };
    private static readonly IReadOnlySet<string> EndSizes = new HashSet<string>(StringComparer.Ordinal) { "", "sm", "med", "lg" };
    private static readonly IReadOnlySet<string> LineCaps = new HashSet<string>(StringComparer.Ordinal) { "", "flat", "round", "square" };
    private static readonly IReadOnlySet<string> LineJoins = new HashSet<string>(StringComparer.Ordinal) { "", "round", "bevel", "miter" };

    internal static bool TryRead(
        P.ConnectionShape source,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId,
        out PresentationConnector connector)
    {
        connector = new PresentationConnector();
        var properties = source.ShapeProperties;
        var transform = properties?.Transform2D;
        var geometry = properties?.GetFirstChild<A.PresetGeometry>();
        var outline = properties?.GetFirstChild<A.Outline>();
        if (properties is null || transform is null || geometry is null || outline is null ||
            !TryGeometryType(geometry, out var connectorType) ||
            !TryTransformEndpoints(transform, out var startX, out var startY, out var endX, out var endY) ||
            !TryOutline(outline, out var lineRgb, out var lineWidth, out var lineStyle, out var cap, out var join,
                out var startArrow, out var startArrowWidth, out var startArrowLength,
                out var endArrow, out var endArrowWidth, out var endArrowLength) ||
            properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.Outline)) return false;

        var nonVisual = source.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
        if (nonVisual is null ||
            !TryConnectionTargets(nonVisual, elementIdsByNativeId,
                out var startTargetId, out var startSiteIndex,
                out var endTargetId, out var endSiteIndex)) return false;

        connector = new PresentationConnector
        {
            ConnectorType = connectorType,
            StartXEmu = startX,
            StartYEmu = startY,
            EndXEmu = endX,
            EndYEmu = endY,
            LineRgb = lineRgb,
            LineWidthEmu = lineWidth,
            StartArrow = startArrow,
            EndArrow = endArrow,
            StartTargetId = startTargetId,
            EndTargetId = endTargetId,
            StartConnectionSiteIndex = startSiteIndex,
            EndConnectionSiteIndex = endSiteIndex,
            LineStyle = lineStyle,
            StartArrowWidth = startArrowWidth,
            StartArrowLength = startArrowLength,
            EndArrowWidth = endArrowWidth,
            EndArrowLength = endArrowLength,
            LineCap = cap,
            LineJoin = join,
        };
        return true;
    }

    internal static P.ConnectionShape Build(
        PresentationElement source,
        uint nativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        Validate(source.Connector, source.Id, source.Name, nativeIdsByElementId);
        var semantic = source.Connector;
        var drawingProperties = new P.NonVisualConnectorShapeDrawingProperties();
        ApplyConnectionTargets(drawingProperties, semantic, nativeIdsByElementId);
        var properties = new P.ShapeProperties(
            ConnectorTransform(semantic),
            CanonicalGeometry(semantic.ConnectorType),
            ConnectorOutline(semantic));
        return new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name },
                drawingProperties,
                new P.ApplicationNonVisualDrawingProperties()),
            properties);
    }

    internal static void Apply(
        P.ConnectionShape source,
        PresentationElement requested,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        Validate(requested.Connector, requested.Id, requested.Name, nativeIdsByElementId);
        source.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!.Name = requested.Name;
        var drawingProperties = source.NonVisualConnectionShapeProperties.NonVisualConnectorShapeDrawingProperties ??= new P.NonVisualConnectorShapeDrawingProperties();
        ApplyConnectionTargets(drawingProperties, requested.Connector, nativeIdsByElementId);

        var properties = source.ShapeProperties ??= new P.ShapeProperties();
        properties.RemoveAllChildren<A.Transform2D>();
        properties.PrependChild(ConnectorTransform(requested.Connector));
        var geometry = properties.GetFirstChild<A.PresetGeometry>();
        if (geometry is null || !TryGeometryType(geometry, out var existingType) || existingType != requested.Connector.ConnectorType)
        {
            geometry?.Remove();
            properties.InsertAfter(CanonicalGeometry(requested.Connector.ConnectorType), properties.Transform2D);
        }
        properties.GetFirstChild<A.Outline>()?.Remove();
        properties.Append(ConnectorOutline(requested.Connector));
    }

    internal static void Validate(
        PresentationConnector? source,
        string elementId,
        string name,
        IReadOnlyDictionary<string, uint>? nativeIdsByElementId = null)
    {
        if (source is null) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} payload is missing.");
        if (name.Length > 1_024) throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} name exceeds 1024 characters.");
        if (!ConnectorTypes.Contains(source.ConnectorType)) throw new CodecException("unsupported_presentation_connector", $"Presentation connector {elementId} uses unsupported type {source.ConnectorType}.");
        if (source.StartXEmu < 0 || source.StartYEmu < 0 || source.EndXEmu < 0 || source.EndYEmu < 0 ||
            source.LineWidthEmu < 0 || source.LineWidthEmu > int.MaxValue)
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} has invalid endpoints or line width.");
        if (!string.IsNullOrWhiteSpace(source.LineRgb)) PptxColor.Normalize(source.LineRgb);
        if (!LineStyles.Contains(source.LineStyle) || !ArrowTypes.Contains(source.StartArrow) || !ArrowTypes.Contains(source.EndArrow) ||
            !EndSizes.Contains(source.StartArrowWidth) || !EndSizes.Contains(source.StartArrowLength) ||
            !EndSizes.Contains(source.EndArrowWidth) || !EndSizes.Contains(source.EndArrowLength) ||
            !LineCaps.Contains(source.LineCap) || !LineJoins.Contains(source.LineJoin))
            throw new CodecException("unsupported_presentation_connector", $"Presentation connector {elementId} uses unsupported line styling.");
        if (source.StartArrow.Length == 0 && (source.StartArrowWidth.Length > 0 || source.StartArrowLength.Length > 0) ||
            source.EndArrow.Length == 0 && (source.EndArrowWidth.Length > 0 || source.EndArrowLength.Length > 0) ||
            source.StartTargetId.Length == 0 && source.StartConnectionSiteIndex != 0 ||
            source.EndTargetId.Length == 0 && source.EndConnectionSiteIndex != 0)
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} has incomplete endpoint or line-end state.");
        if (nativeIdsByElementId is null) return;
        if (source.StartTargetId.Length > 0 && !nativeIdsByElementId.ContainsKey(source.StartTargetId))
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} references missing start target {source.StartTargetId}.");
        if (source.EndTargetId.Length > 0 && !nativeIdsByElementId.ContainsKey(source.EndTargetId))
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} references missing end target {source.EndTargetId}.");
    }

    private static bool TryGeometryType(A.PresetGeometry geometry, out string connectorType)
    {
        connectorType = string.Empty;
        var preset = geometry.Preset?.Value;
        if (preset is null || geometry.ExtendedAttributes.Any()) return false;
        if (preset.Value.Equals(A.ShapeTypeValues.Line) || preset.Value.Equals(A.ShapeTypeValues.StraightConnector1)) connectorType = "straight";
        else if (preset.Value.Equals(A.ShapeTypeValues.BentConnector3)) connectorType = "elbow";
        else if (preset.Value.Equals(A.ShapeTypeValues.CurvedConnector3)) connectorType = "curved";
        else return false;

        var lists = geometry.Elements<A.AdjustValueList>().ToArray();
        if (geometry.ChildElements.Any(child => child is not A.AdjustValueList) || lists.Length > 1) return false;
        var guides = lists.SingleOrDefault()?.Elements<A.ShapeGuide>().ToArray() ?? [];
        if (lists.SingleOrDefault()?.ChildElements.Any(child => child is not A.ShapeGuide) == true) return false;
        if (connectorType != "curved") return guides.Length == 0;
        return guides.Length == 0 || guides.Length == 1 && guides[0].Name?.Value == "adj1" && guides[0].Formula?.Value == "val 50000" &&
            !guides[0].ChildElements.Any() && !guides[0].ExtendedAttributes.Any();
    }

    private static A.PresetGeometry CanonicalGeometry(string connectorType)
    {
        var adjustments = new A.AdjustValueList();
        A.ShapeTypeValues preset;
        if (connectorType == "curved")
        {
            preset = A.ShapeTypeValues.CurvedConnector3;
            adjustments.Append(new A.ShapeGuide { Name = "adj1", Formula = "val 50000" });
        }
        else preset = connectorType == "elbow" ? A.ShapeTypeValues.BentConnector3 : A.ShapeTypeValues.StraightConnector1;
        return new A.PresetGeometry(adjustments) { Preset = preset };
    }

    private static bool TryTransformEndpoints(A.Transform2D transform, out long startX, out long startY, out long endX, out long endY)
    {
        startX = startY = endX = endY = 0;
        var left = transform.Offset?.X?.Value;
        var top = transform.Offset?.Y?.Value;
        var width = transform.Extents?.Cx?.Value;
        var height = transform.Extents?.Cy?.Value;
        if (left is null || top is null || width is null or < 0 || height is null or < 0 ||
            transform.ChildElements.Any(child => child is not A.Offset and not A.Extents) || transform.ExtendedAttributes.Any()) return false;

        var localStartX = transform.HorizontalFlip?.Value == true ? left.Value + width.Value : left.Value;
        var localStartY = transform.VerticalFlip?.Value == true ? top.Value + height.Value : top.Value;
        var localEndX = transform.HorizontalFlip?.Value == true ? left.Value : left.Value + width.Value;
        var localEndY = transform.VerticalFlip?.Value == true ? top.Value : top.Value + height.Value;
        var rotation = transform.Rotation?.Value ?? 0;
        var centerX = left.Value + width.Value / 2d;
        var centerY = top.Value + height.Value / 2d;
        (startX, startY) = RotatePoint(localStartX, localStartY, centerX, centerY, rotation);
        (endX, endY) = RotatePoint(localEndX, localEndY, centerX, centerY, rotation);
        return startX >= 0 && startY >= 0 && endX >= 0 && endY >= 0;
    }

    private static (long X, long Y) RotatePoint(long x, long y, double centerX, double centerY, int rotation)
    {
        if (rotation == 0) return (x, y);
        var radians = rotation / (double)RotationUnitsPerDegree * Math.PI / 180d;
        var dx = x - centerX;
        var dy = y - centerY;
        return (
            checked((long)Math.Round(centerX + dx * Math.Cos(radians) - dy * Math.Sin(radians))),
            checked((long)Math.Round(centerY + dx * Math.Sin(radians) + dy * Math.Cos(radians))));
    }

    private static bool TryOutline(
        A.Outline outline,
        out string lineRgb,
        out long lineWidth,
        out string lineStyle,
        out string cap,
        out string join,
        out string startArrow,
        out string startArrowWidth,
        out string startArrowLength,
        out string endArrow,
        out string endArrowWidth,
        out string endArrowLength)
    {
        lineRgb = lineStyle = cap = join = startArrow = startArrowWidth = startArrowLength = endArrow = endArrowWidth = endArrowLength = string.Empty;
        lineWidth = outline.Width?.Value ?? 0;
        if (lineWidth < 0 || lineWidth > int.MaxValue || outline.ExtendedAttributes.Any() ||
            outline.CompoundLineType is not null || outline.Alignment is not null) return false;

        var noFills = outline.Elements<A.NoFill>().ToArray();
        var solidFills = outline.Elements<A.SolidFill>().ToArray();
        if (noFills.Length + solidFills.Length != 1) return false;
        if (solidFills.Length == 1)
        {
            var fill = solidFills[0];
            var rgb = fill.GetFirstChild<A.RgbColorModelHex>();
            if (rgb?.Val?.Value is not string value || value.Length != 6 || fill.ChildElements.Count != 1 || fill.ExtendedAttributes.Any() || rgb.ChildElements.Any() || rgb.ExtendedAttributes.Any()) return false;
            lineRgb = value.ToUpperInvariant();
        }

        var dashes = outline.Elements<A.PresetDash>().ToArray();
        if (dashes.Length > 1 || dashes.SingleOrDefault()?.ChildElements.Any() == true || dashes.SingleOrDefault()?.ExtendedAttributes.Any() == true) return false;
        var dash = dashes.SingleOrDefault()?.Val?.Value;
        if (noFills.Length == 1)
        {
            if (dash is not null) return false;
            lineStyle = "none";
        }
        else if (dash is null || dash.Value.Equals(A.PresetLineDashValues.Solid)) lineStyle = "solid";
        else if (dash.Value.Equals(A.PresetLineDashValues.Dash)) lineStyle = "dashed";
        else return false;

        if (!TryCap(outline.CapType?.Value, out cap) || !TryJoin(outline, out join)) return false;
        var heads = outline.Elements<A.HeadEnd>().ToArray();
        var tails = outline.Elements<A.TailEnd>().ToArray();
        if (heads.Length > 1 || tails.Length > 1 ||
            !TryLineEnd(heads.SingleOrDefault(), out startArrow, out startArrowWidth, out startArrowLength) ||
            !TryLineEnd(tails.SingleOrDefault(), out endArrow, out endArrowWidth, out endArrowLength)) return false;

        return outline.ChildElements.All(child => child is A.NoFill or A.SolidFill or A.PresetDash or A.Round or A.LineJoinBevel or A.Miter or A.HeadEnd or A.TailEnd);
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
            if (round.ChildElements.Any() || round.ExtendedAttributes.Any()) return false;
            join = "round";
        }
        else if (joins.SingleOrDefault() is A.LineJoinBevel bevel)
        {
            if (bevel.ChildElements.Any() || bevel.ExtendedAttributes.Any()) return false;
            join = "bevel";
        }
        else if (joins.SingleOrDefault() is A.Miter miter)
        {
            if (miter.ChildElements.Any() || miter.ExtendedAttributes.Any() || miter.Limit is not null) return false;
            join = "miter";
        }
        return true;
    }

    private static bool TryLineEnd(A.HeadEnd? source, out string type, out string width, out string length) =>
        TryLineEnd(source, source?.Type?.Value, source?.Width?.Value, source?.Length?.Value, out type, out width, out length);

    private static bool TryLineEnd(A.TailEnd? source, out string type, out string width, out string length) =>
        TryLineEnd(source, source?.Type?.Value, source?.Width?.Value, source?.Length?.Value, out type, out width, out length);

    private static bool TryLineEnd(
        OpenXmlElement? source,
        A.LineEndValues? sourceType,
        A.LineEndWidthValues? sourceWidth,
        A.LineEndLengthValues? sourceLength,
        out string type,
        out string width,
        out string length)
    {
        type = width = length = string.Empty;
        if (source is null) return true;
        if (source.ChildElements.Any() || source.ExtendedAttributes.Any() || !TryArrow(sourceType, out type) ||
            !TryEndWidth(sourceWidth, out width) || !TryEndLength(sourceLength, out length)) return false;
        return type.Length > 0 || width.Length == 0 && length.Length == 0;
    }

    private static bool TryArrow(A.LineEndValues? source, out string arrow)
    {
        arrow = string.Empty;
        if (source is null || source.Value.Equals(A.LineEndValues.None)) return true;
        if (source.Value.Equals(A.LineEndValues.Triangle)) arrow = "triangle";
        else if (source.Value.Equals(A.LineEndValues.Stealth)) arrow = "stealth";
        else if (source.Value.Equals(A.LineEndValues.Diamond)) arrow = "diamond";
        else if (source.Value.Equals(A.LineEndValues.Oval)) arrow = "oval";
        else if (source.Value.Equals(A.LineEndValues.Arrow)) arrow = "arrow";
        else return false;
        return true;
    }

    private static bool TryEndWidth(A.LineEndWidthValues? source, out string width)
    {
        width = string.Empty;
        if (source is null) return true;
        if (source.Value.Equals(A.LineEndWidthValues.Small)) width = "sm";
        else if (source.Value.Equals(A.LineEndWidthValues.Medium)) width = "med";
        else if (source.Value.Equals(A.LineEndWidthValues.Large)) width = "lg";
        else return false;
        return true;
    }

    private static bool TryEndLength(A.LineEndLengthValues? source, out string length)
    {
        length = string.Empty;
        if (source is null) return true;
        if (source.Value.Equals(A.LineEndLengthValues.Small)) length = "sm";
        else if (source.Value.Equals(A.LineEndLengthValues.Medium)) length = "med";
        else if (source.Value.Equals(A.LineEndLengthValues.Large)) length = "lg";
        else return false;
        return true;
    }

    private static bool TryConnectionTargets(
        P.NonVisualConnectorShapeDrawingProperties source,
        IReadOnlyDictionary<uint, string>? ids,
        out string startTargetId,
        out uint startSiteIndex,
        out string endTargetId,
        out uint endSiteIndex)
    {
        startTargetId = endTargetId = string.Empty;
        startSiteIndex = endSiteIndex = 0;
        var starts = source.Elements<A.StartConnection>().ToArray();
        var ends = source.Elements<A.EndConnection>().ToArray();
        return starts.Length <= 1 && ends.Length <= 1 &&
            TryConnectionTarget(starts.SingleOrDefault(), ids, out startTargetId, out startSiteIndex) &&
            TryConnectionTarget(ends.SingleOrDefault(), ids, out endTargetId, out endSiteIndex);
    }

    private static bool TryConnectionTarget(A.StartConnection? source, IReadOnlyDictionary<uint, string>? ids, out string targetId, out uint siteIndex) =>
        TryConnectionTarget(source, source?.Id?.Value, source?.Index?.Value, ids, out targetId, out siteIndex);

    private static bool TryConnectionTarget(A.EndConnection? source, IReadOnlyDictionary<uint, string>? ids, out string targetId, out uint siteIndex) =>
        TryConnectionTarget(source, source?.Id?.Value, source?.Index?.Value, ids, out targetId, out siteIndex);

    private static bool TryConnectionTarget(OpenXmlElement? source, uint? nativeId, uint? index, IReadOnlyDictionary<uint, string>? ids, out string targetId, out uint siteIndex)
    {
        targetId = string.Empty;
        siteIndex = 0;
        if (source is null) return true;
        if (nativeId is null || index is null || source.ChildElements.Any() || source.ExtendedAttributes.Any()) return false;
        siteIndex = index.Value;
        return ids is not null && ids.TryGetValue(nativeId.Value, out targetId!);
    }

    private static void ApplyConnectionTargets(
        P.NonVisualConnectorShapeDrawingProperties properties,
        PresentationConnector source,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId)
    {
        properties.RemoveAllChildren<A.StartConnection>();
        properties.RemoveAllChildren<A.EndConnection>();
        var extensionList = properties.GetFirstChild<A.ExtensionList>();
        void Insert(OpenXmlElement connection)
        {
            if (extensionList is null) properties.Append(connection);
            else properties.InsertBefore(connection, extensionList);
        }
        if (source.StartTargetId.Length > 0) Insert(new A.StartConnection
        {
            Id = nativeIdsByElementId[source.StartTargetId],
            Index = source.StartConnectionSiteIndex,
        });
        if (source.EndTargetId.Length > 0) Insert(new A.EndConnection
        {
            Id = nativeIdsByElementId[source.EndTargetId],
            Index = source.EndConnectionSiteIndex,
        });
    }

    private static A.Transform2D ConnectorTransform(PresentationConnector source)
    {
        var left = Math.Min(source.StartXEmu, source.EndXEmu);
        var top = Math.Min(source.StartYEmu, source.EndYEmu);
        return new A.Transform2D(
            new A.Offset { X = left, Y = top },
            new A.Extents { Cx = Math.Abs(source.EndXEmu - source.StartXEmu), Cy = Math.Abs(source.EndYEmu - source.StartYEmu) })
        {
            HorizontalFlip = source.EndXEmu < source.StartXEmu,
            VerticalFlip = source.EndYEmu < source.StartYEmu,
        };
    }

    private static A.Outline ConnectorOutline(PresentationConnector source)
    {
        var outline = new A.Outline { Width = checked((int)source.LineWidthEmu) };
        if (source.LineCap.Length > 0) outline.CapType = source.LineCap switch
        {
            "round" => A.LineCapValues.Round,
            "square" => A.LineCapValues.Square,
            _ => A.LineCapValues.Flat,
        };
        outline.Append(string.IsNullOrWhiteSpace(source.LineRgb) || source.LineStyle == "none"
            ? new A.NoFill()
            : new A.SolidFill(new A.RgbColorModelHex { Val = PptxColor.Normalize(source.LineRgb) }));
        if (source.LineStyle == "dashed") outline.Append(new A.PresetDash { Val = A.PresetLineDashValues.Dash });
        if (source.LineJoin.Length > 0) outline.Append(source.LineJoin switch
        {
            "round" => new A.Round(),
            "bevel" => new A.LineJoinBevel(),
            _ => new A.Miter(),
        });
        if (source.StartArrow.Length > 0) outline.Append(HeadEnd(source.StartArrow, source.StartArrowWidth, source.StartArrowLength));
        if (source.EndArrow.Length > 0) outline.Append(TailEnd(source.EndArrow, source.EndArrowWidth, source.EndArrowLength));
        return outline;
    }

    private static A.HeadEnd HeadEnd(string type, string width, string length)
    {
        var output = new A.HeadEnd { Type = LineEndType(type) };
        if (width.Length > 0) output.Width = LineEndWidth(width);
        if (length.Length > 0) output.Length = LineEndLength(length);
        return output;
    }

    private static A.TailEnd TailEnd(string type, string width, string length)
    {
        var output = new A.TailEnd { Type = LineEndType(type) };
        if (width.Length > 0) output.Width = LineEndWidth(width);
        if (length.Length > 0) output.Length = LineEndLength(length);
        return output;
    }

    private static A.LineEndValues LineEndType(string type) => type switch
    {
        "stealth" => A.LineEndValues.Stealth,
        "diamond" => A.LineEndValues.Diamond,
        "oval" => A.LineEndValues.Oval,
        "arrow" => A.LineEndValues.Arrow,
        _ => A.LineEndValues.Triangle,
    };

    private static A.LineEndWidthValues LineEndWidth(string width) => width switch
    {
        "sm" => A.LineEndWidthValues.Small,
        "lg" => A.LineEndWidthValues.Large,
        _ => A.LineEndWidthValues.Medium,
    };

    private static A.LineEndLengthValues LineEndLength(string length) => length switch
    {
        "sm" => A.LineEndLengthValues.Small,
        "lg" => A.LineEndLengthValues.Large,
        _ => A.LineEndLengthValues.Medium,
    };
}
