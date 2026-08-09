using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal static class PptxConnectorCodec
{
    private const int RotationUnitsPerDegree = 60_000;
    private static readonly IReadOnlySet<string> ConnectorTypes = new HashSet<string>(StringComparer.Ordinal) { "straight", "elbow", "curved" };

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
            !PptxLineStyleCodec.TryRead(outline, out var lineStyle) ||
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
            StartTargetId = startTargetId,
            EndTargetId = endTargetId,
            StartConnectionSiteIndex = startSiteIndex,
            EndConnectionSiteIndex = endSiteIndex,
        };
        PptxLineStyleCodec.CopyTo(lineStyle, connector);
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
            PptxLineStyleCodec.Build(semantic));
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
        properties.Append(PptxLineStyleCodec.Build(requested.Connector));
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
        if (source.StartXEmu < 0 || source.StartYEmu < 0 || source.EndXEmu < 0 || source.EndYEmu < 0)
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} has invalid endpoints.");
        PptxLineStyleCodec.Validate(source, elementId);
        if (source.StartTargetId.Length == 0 && source.StartConnectionSiteIndex != 0 ||
            source.EndTargetId.Length == 0 && source.EndConnectionSiteIndex != 0)
            throw new CodecException("invalid_presentation_connector", $"Presentation connector {elementId} has incomplete connection-target state.");
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

}
