using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using OD = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the source-free, clean-room SmartArt package graph. PPJ and the wire
// remain semantic; part paths, relationship IDs, model GUIDs, and the cached
// drawing are all derived here at the PPTX boundary.
internal static class PptxSmartArtCodec
{
    private const string DiagramUri = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string DrawingUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string Drawing2010Uri = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    private const string RelationshipUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string OfficeKitUri = "urn:officekit:smartart:v1";
    private const string OfficeKitExtension = "{4D1D90A1-58F8-4E34-9D03-0A2C74E285B0}";
    private const string DrawingExtension = "{C28D6A17-3A09-4B91-8F6C-3E9D4F6A8C12}";
    private const string MinVersion = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    private static readonly XNamespace DgmNs = DiagramUri;
    private static readonly XNamespace ANs = DrawingUri;
    private static readonly XNamespace DspNs = Drawing2010Uri;
    private static readonly XNamespace RNs = RelationshipUri;
    private static readonly XNamespace OkNs = OfficeKitUri;

    internal static P.GraphicFrame Build(
        PresentationElement element,
        uint nativeId,
        PptxPartContext slideContext,
        SlidePart slidePart)
    {
        var diagram = element.Diagram;
        var relationshipSuffix = nativeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var dataPart = slidePart.AddNewPart<DiagramDataPart>(slideContext.NextRelationshipId($"rIdOfficeKitSmartArtData{relationshipSuffix}"));
        var layoutPart = slidePart.AddNewPart<DiagramLayoutDefinitionPart>(slideContext.NextRelationshipId($"rIdOfficeKitSmartArtLayout{relationshipSuffix}"));
        var stylePart = slidePart.AddNewPart<DiagramStylePart>(slideContext.NextRelationshipId($"rIdOfficeKitSmartArtStyle{relationshipSuffix}"));
        var colorsPart = slidePart.AddNewPart<DiagramColorsPart>(slideContext.NextRelationshipId($"rIdOfficeKitSmartArtColors{relationshipSuffix}"));
        var drawingPart = slidePart.AddNewPart<DiagramPersistLayoutPart>(slideContext.NextRelationshipId($"rIdOfficeKitSmartArtDrawing{relationshipSuffix}"));
        slideContext.TrackAddedPart(dataPart);
        slideContext.TrackAddedPart(layoutPart);
        slideContext.TrackAddedPart(stylePart);
        slideContext.TrackAddedPart(colorsPart);
        slideContext.TrackAddedPart(drawingPart);
        var drawingRelationshipId = slidePart.GetIdOfPart(drawingPart);
        var imageRelationshipIds = AddCachedDrawingImages(diagram, drawingPart, slideContext);

        var definition = string.IsNullOrWhiteSpace(diagram.DefinitionAssetId)
            ? null
            : (slideContext.Assets ?? throw new CodecException(
                "invalid_presentation_asset",
                "Presentation SmartArt custom definitions require an asset catalog."))
                .GetSmartArtDefinition(diagram.DefinitionAssetId);
        var parsedDefinition = definition is null
            ? null
            : PpjSmartArtDefinitionCodec.Parse(definition, "presentation.diagram.definition");
        var styleId = parsedDefinition?.Root.GetProperty("style").GetProperty("id").GetString() ?? "basic";
        var colorsId = parsedDefinition?.Root.GetProperty("colors").GetProperty("id").GetString() ?? "accent";

        dataPart.DataModelRoot = new Dgm.DataModelRoot(BuildDataModel(
            element.Id,
            diagram,
            drawingRelationshipId,
            styleId,
            colorsId));
        layoutPart.LayoutDefinition = new Dgm.LayoutDefinition(BuildLayoutDefinition(diagram.Layout, definition));
        stylePart.StyleDefinition = new Dgm.StyleDefinition(BuildStyleDefinition(styleId));
        colorsPart.ColorsDefinition = new Dgm.ColorsDefinition(BuildColorsDefinition(colorsId));
        drawingPart.Drawing = new OD.Drawing(BuildCachedDrawing(element.Id, diagram, imageRelationshipIds));
        dataPart.DataModelRoot.Save();
        layoutPart.LayoutDefinition.Save();
        stylePart.StyleDefinition.Save();
        colorsPart.ColorsDefinition.Save();
        drawingPart.Drawing.Save();

        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = element.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, diagram.Accessibility);
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                nonVisual,
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = diagram.LeftEmu, Y = diagram.TopEmu },
                new A.Extents { Cx = diagram.WidthEmu, Cy = diagram.HeightEmu }),
            new A.Graphic(
                new A.GraphicData(
                    new Dgm.RelationshipIds
                    {
                        DataPart = slidePart.GetIdOfPart(dataPart),
                        LayoutPart = slidePart.GetIdOfPart(layoutPart),
                        StylePart = slidePart.GetIdOfPart(stylePart),
                        ColorPart = slidePart.GetIdOfPart(colorsPart),
                    })
                { Uri = DiagramUri }));
    }

    internal static bool TryRead(
        P.GraphicFrame frame,
        SlidePart slidePart,
        PptxAssetCatalog assets,
        out PresentationDiagram diagram)
    {
        diagram = null!;
        var relationshipIds = frame.Descendants<Dgm.RelationshipIds>().ToArray();
        if (relationshipIds.Length != 1 || frame.Transform?.Offset is not { } offset || frame.Transform.Extents is not { } extents)
            return false;
        DiagramDataPart dataPart;
        DiagramLayoutDefinitionPart layoutPart;
        DiagramPersistLayoutPart? drawingPart;
        try
        {
            dataPart = slidePart.GetPartById(relationshipIds[0].DataPart?.Value ?? string.Empty) as DiagramDataPart ?? throw new InvalidDataException();
            layoutPart = slidePart.GetPartById(relationshipIds[0].LayoutPart?.Value ?? string.Empty) as DiagramLayoutDefinitionPart ?? throw new InvalidDataException();
            drawingPart = slidePart.Parts.Select(pair => pair.OpenXmlPart).OfType<DiagramPersistLayoutPart>().SingleOrDefault(part =>
                DataModelDrawingRelationshipId(dataPart) is { } id && ReferenceEquals(slidePart.GetPartById(id), part));
        }
        catch (Exception exception) when (exception is ArgumentOutOfRangeException or InvalidDataException or InvalidOperationException)
        {
            return false;
        }

        XDocument data;
        try
        {
            data = ReadXml(dataPart);
        }
        catch (Exception exception) when (exception is IOException or System.Xml.XmlException)
        {
            return false;
        }
        var root = data.Root;
        if (root?.Name != DgmNs + "dataModel") return false;
        var docPoints = root.Element(DgmNs + "ptLst")?.Elements(DgmNs + "pt")
            .Where(point => point.Attribute("type")?.Value == "doc")
            .Take(2)
            .ToArray() ?? [];
        if (docPoints.Length != 1) return false;
        var docPoint = docPoints[0];
        var layoutId = docPoint?.Element(DgmNs + "prSet")?.Attribute("loTypeId")?.Value;
        const string prefix = "urn:officekit:smartart:v1:layout:";
        if (layoutId is null || !layoutId.StartsWith(prefix, StringComparison.Ordinal)) return false;

        var output = new PresentationDiagram
        {
            Layout = layoutId[prefix.Length..],
            LeftEmu = offset.X?.Value ?? 0,
            TopEmu = offset.Y?.Value ?? 0,
            WidthEmu = extents.Cx?.Value ?? 0,
            HeightEmu = extents.Cy?.Value ?? 0,
        };
        if (!TryReadDefinitionAsset(layoutPart, assets, out var definitionAssetId)) return false;
        output.DefinitionAssetId = definitionAssetId;
        var modelIdsByNodeId = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var point in root.Element(DgmNs + "ptLst")?.Elements(DgmNs + "pt") ?? [])
        {
            var metadata = point.Descendants(OkNs + "node").SingleOrDefault();
            var id = metadata?.Attribute("id")?.Value;
            var modelId = point.Attribute("modelId")?.Value;
            var textElement = point.Element(DgmNs + "t");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(modelId) || textElement is null || !TryReadTextBody(textElement, out var textBody))
                continue;
            if (!modelIdsByNodeId.TryAdd(id, modelId)) return false;
            output.Nodes.Add(new PresentationDiagramNode
            {
                Id = id,
                TextBody = textBody,
                AssetId = metadata?.Attribute("asset")?.Value ?? string.Empty,
            });
        }
        if (output.Nodes.Count == 0) return false;

        foreach (var connection in root.Element(DgmNs + "cxnLst")?.Elements(DgmNs + "cxn") ?? [])
        {
            var metadata = connection.Descendants(OkNs + "connection").SingleOrDefault();
            if (metadata is null) continue;
            var fromId = metadata.Attribute("from")?.Value ?? string.Empty;
            var toId = metadata.Attribute("to")?.Value ?? string.Empty;
            if (!modelIdsByNodeId.ContainsKey(fromId) || !modelIdsByNodeId.ContainsKey(toId)) return false;
            output.Connections.Add(new PresentationDiagramConnection
            {
                Id = metadata.Attribute("id")?.Value ?? string.Empty,
                FromId = fromId,
                ToId = toId,
                Role = metadata.Attribute("role")?.Value ?? string.Empty,
                Order = uint.TryParse(metadata.Attribute("order")?.Value, out var order) ? order : 0,
            });
        }
        output.Drawing = ReadCachedDrawing(output, drawingPart, assets, out var drawingCacheVerified);
        output.DrawingCacheVerified = drawingCacheVerified;
        diagram = output;
        return true;
    }

    internal static void Validate(PresentationDiagram diagram, string elementId, PptxAssetCatalog assets)
    {
        if (string.IsNullOrWhiteSpace(diagram.Layout) || diagram.LeftEmu < 0 || diagram.TopEmu < 0 ||
            diagram.WidthEmu <= 0 || diagram.HeightEmu <= 0)
            throw new CodecException("invalid_presentation_diagram", $"Presentation SmartArt {elementId} requires a layout and positive frame.");
        if (diagram.Nodes.Count is < 1 or > 64)
            throw new CodecException("invalid_presentation_diagram", $"Presentation SmartArt {elementId} requires between 1 and 64 nodes.");
        var nodeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in diagram.Nodes)
        {
            if (string.IsNullOrWhiteSpace(node.Id) || !nodeIds.Add(node.Id))
                throw new CodecException("invalid_presentation_diagram", $"Presentation SmartArt {elementId} contains an empty or duplicate node ID.");
            var shape = new PresentationShape { Text = PptxTextCodec.Flatten(node.TextBody), TextBody = node.TextBody?.Clone() };
            PptxTextCodec.Validate(shape);
            if (!string.IsNullOrWhiteSpace(node.AssetId)) _ = assets.Get(node.AssetId);
        }
        var connectionIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var connection in diagram.Connections)
        {
            if (string.IsNullOrWhiteSpace(connection.Id) || !connectionIds.Add(connection.Id) ||
                !nodeIds.Contains(connection.FromId) || !nodeIds.Contains(connection.ToId) || connection.FromId == connection.ToId ||
                connection.Role is not ("parent" or "sequence" or "association"))
                throw new CodecException("invalid_presentation_diagram", $"Presentation SmartArt {elementId} contains an invalid connection.");
        }
        if (diagram.Drawing is null || diagram.Drawing.Children.Count == 0)
            throw new CodecException("invalid_presentation_diagram", $"Presentation SmartArt {elementId} requires a non-empty cached drawing.");
        if (!string.IsNullOrWhiteSpace(diagram.DefinitionAssetId))
            _ = assets.GetSmartArtDefinition(diagram.DefinitionAssetId);
        PptxNonVisualAccessibilityCodec.Validate(diagram.Accessibility, elementId, "SmartArt");
    }

    private static string BuildDataModel(
        string elementId,
        PresentationDiagram diagram,
        string drawingRelationshipId,
        string styleId,
        string colorsId)
    {
        var layoutId = LayoutId(diagram.Layout);
        var documentId = StableModelId(elementId, "document");
        var presentationRootId = StableModelId(elementId, "presentation-root");
        var dataIds = diagram.Nodes.ToDictionary(node => node.Id, node => StableModelId(elementId, $"data:{node.Id}"), StringComparer.Ordinal);
        var presentationIds = diagram.Nodes.ToDictionary(node => node.Id, node => StableModelId(elementId, $"presentation:{node.Id}"), StringComparer.Ordinal);
        var pointList = new XElement(DgmNs + "ptLst",
            new XElement(DgmNs + "pt",
                new XAttribute("modelId", documentId),
                new XAttribute("type", "doc"),
                new XElement(DgmNs + "prSet",
                    new XAttribute("loTypeId", layoutId),
                    new XAttribute("loCatId", diagram.Layout),
                    new XAttribute("qsTypeId", $"urn:officekit:smartart:v1:quickstyle:{styleId}"),
                    new XAttribute("qsCatId", "simple"),
                    new XAttribute("csTypeId", $"urn:officekit:smartart:v1:colors:{colorsId}"),
                    new XAttribute("csCatId", "accent")),
                new XElement(DgmNs + "spPr"),
                EmptyDiagramText()));
        foreach (var node in diagram.Nodes)
        {
            var metadata = new XElement(OkNs + "node", new XAttribute("id", node.Id));
            if (!string.IsNullOrWhiteSpace(node.AssetId)) metadata.Add(new XAttribute("asset", node.AssetId));
            pointList.Add(new XElement(DgmNs + "pt",
                new XAttribute("modelId", dataIds[node.Id]),
                new XElement(DgmNs + "prSet"),
                new XElement(DgmNs + "spPr"),
                DiagramText(node.TextBody),
                Extension(metadata)));
        }
        pointList.Add(new XElement(DgmNs + "pt",
            new XAttribute("modelId", presentationRootId),
            new XAttribute("type", "pres"),
            new XElement(DgmNs + "prSet",
                new XAttribute("presAssocID", documentId),
                new XAttribute("presName", "diagram"),
                new XAttribute("presStyleCnt", "0")),
            new XElement(DgmNs + "spPr")));
        foreach (var (node, index) in diagram.Nodes.Select((node, index) => (node, index)))
        {
            pointList.Add(new XElement(DgmNs + "pt",
                new XAttribute("modelId", presentationIds[node.Id]),
                new XAttribute("type", "pres"),
                new XElement(DgmNs + "prSet",
                    new XAttribute("presAssocID", dataIds[node.Id]),
                    new XAttribute("presName", "node"),
                    new XAttribute("presStyleLbl", "node1"),
                    new XAttribute("presStyleIdx", index),
                    new XAttribute("presStyleCnt", diagram.Nodes.Count)),
                new XElement(DgmNs + "spPr")));
        }

        var connectionList = new XElement(DgmNs + "cxnLst");
        foreach (var (node, index) in diagram.Nodes.Select((node, index) => (node, index)))
        {
            connectionList.Add(Connection(elementId, $"document:{node.Id}", documentId, dataIds[node.Id], index));
            connectionList.Add(Connection(elementId, $"presentation-of:{node.Id}", dataIds[node.Id], presentationIds[node.Id], 0, "presOf", layoutId));
            connectionList.Add(Connection(elementId, $"presentation-parent:{node.Id}", presentationRootId, presentationIds[node.Id], index, "presParOf", layoutId));
        }
        foreach (var connection in diagram.Connections)
        {
            connectionList.Add(Connection(
                elementId,
                $"semantic:{connection.Id}",
                dataIds[connection.FromId],
                dataIds[connection.ToId],
                checked((int)connection.Order),
                "unknownRelationship",
                extension: Extension(new XElement(OkNs + "connection",
                    new XAttribute("id", connection.Id),
                    new XAttribute("from", connection.FromId),
                    new XAttribute("to", connection.ToId),
                    new XAttribute("role", connection.Role),
                    new XAttribute("order", connection.Order)))));
        }
        var root = new XElement(DgmNs + "dataModel",
            new XAttribute(XNamespace.Xmlns + "dgm", DgmNs),
            new XAttribute(XNamespace.Xmlns + "a", ANs),
            new XAttribute(XNamespace.Xmlns + "dsp", DspNs),
            new XAttribute(XNamespace.Xmlns + "r", RNs),
            new XAttribute(XNamespace.Xmlns + "ok", OkNs),
            pointList,
            connectionList,
            new XElement(DgmNs + "bg"),
            new XElement(DgmNs + "whole"),
            new XElement(DgmNs + "extLst",
                new XElement(ANs + "ext",
                    new XAttribute("uri", DrawingExtension),
                    new XElement(DspNs + "dataModelExt",
                        new XAttribute("relId", drawingRelationshipId),
                        new XAttribute("minVer", MinVersion)))));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Connection(
        string elementId,
        string identity,
        string sourceId,
        string destinationId,
        int sourceOrder,
        string? type = null,
        string? presentationId = null,
        XElement? extension = null)
    {
        var output = new XElement(DgmNs + "cxn",
            new XAttribute("modelId", StableModelId(elementId, $"connection:{identity}")),
            new XAttribute("srcId", sourceId),
            new XAttribute("destId", destinationId),
            new XAttribute("srcOrd", sourceOrder),
            new XAttribute("destOrd", 0));
        if (type is not null) output.Add(new XAttribute("type", type));
        if (presentationId is not null) output.Add(new XAttribute("presId", presentationId));
        if (extension is not null) output.Add(extension);
        return output;
    }

    private static string BuildLayoutDefinition(string layout, Asset? definition)
    {
        var root = new XElement(DgmNs + "layoutDef",
            new XAttribute(XNamespace.Xmlns + "dgm", DgmNs),
            new XAttribute(XNamespace.Xmlns + "a", ANs),
            new XAttribute(XNamespace.Xmlns + "r", RNs),
            new XAttribute("uniqueId", LayoutId(layout)),
            new XElement(DgmNs + "title", new XAttribute("val", $"OfficeKit {layout}")),
            new XElement(DgmNs + "desc", new XAttribute("val", "Deterministic OfficeKit SmartArt layout")),
            new XElement(DgmNs + "catLst", new XElement(DgmNs + "cat", new XAttribute("type", layout), new XAttribute("pri", 1))),
            new XElement(DgmNs + "layoutNode",
                new XAttribute("name", "diagram"),
                new XElement(DgmNs + "alg", new XAttribute("type", "composite")),
                new XElement(DgmNs + "shape", new XAttribute(RNs + "blip", string.Empty), new XElement(DgmNs + "adjLst")),
                new XElement(DgmNs + "presOf"),
                new XElement(DgmNs + "forEach",
                    new XAttribute("name", "nodes"),
                    new XAttribute("axis", "ch"),
                    new XAttribute("ptType", "node"),
                    new XElement(DgmNs + "layoutNode",
                        new XAttribute("name", "node"),
                        new XElement(DgmNs + "alg", new XAttribute("type", "tx")),
                        new XElement(DgmNs + "shape", new XAttribute("type", "rect"), new XAttribute(RNs + "blip", string.Empty), new XElement(DgmNs + "adjLst")),
                        new XElement(DgmNs + "presOf", new XAttribute("axis", "desOrSelf"), new XAttribute("ptType", "node"))))));
        if (definition is not null)
        {
            root.Add(new XElement(DgmNs + "extLst",
                new XElement(DgmNs + "ext",
                    new XAttribute("uri", OfficeKitExtension),
                    new XElement(OkNs + "definition",
                        new XAttribute("schema", "office-kit/smartart-definition/v1"),
                        new XAttribute("sha256", definition.Sha256),
                        Convert.ToBase64String(definition.Data.ToByteArray())))));
        }
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static bool TryReadDefinitionAsset(
        DiagramLayoutDefinitionPart part,
        PptxAssetCatalog assets,
        out string definitionAssetId)
    {
        definitionAssetId = string.Empty;
        XDocument layout;
        try { layout = ReadXml(part); }
        catch (Exception exception) when (exception is IOException or System.Xml.XmlException) { return false; }
        var definitions = layout.Descendants(OkNs + "definition").Take(2).ToArray();
        if (definitions.Length == 0) return true;
        if (definitions.Length != 1 ||
            definitions[0].Attribute("schema")?.Value != "office-kit/smartart-definition/v1" ||
            definitions[0].Attribute("sha256")?.Value is not { Length: 64 } sha256)
            return false;
        byte[] data;
        try { data = Convert.FromBase64String(definitions[0].Value); }
        catch (FormatException) { return false; }
        try
        {
            definitionAssetId = assets.ImportSmartArtDefinition(data, sha256).Id;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    private static string BuildStyleDefinition(string styleId)
    {
        var root = new XElement(DgmNs + "styleDef",
            new XAttribute(XNamespace.Xmlns + "dgm", DgmNs),
            new XAttribute(XNamespace.Xmlns + "a", ANs),
            new XAttribute("uniqueId", $"urn:officekit:smartart:v1:quickstyle:{styleId}"),
            new XElement(DgmNs + "title", new XAttribute("val", $"OfficeKit {styleId}")),
            new XElement(DgmNs + "desc", new XAttribute("val", "OfficeKit clean-room SmartArt style")),
            new XElement(DgmNs + "catLst", new XElement(DgmNs + "cat", new XAttribute("type", "simple"), new XAttribute("pri", 1))),
            Scene3D(),
            StyleLabel());
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static string BuildColorsDefinition(string colorsId)
    {
        var root = new XElement(DgmNs + "colorsDef",
            new XAttribute(XNamespace.Xmlns + "dgm", DgmNs),
            new XAttribute(XNamespace.Xmlns + "a", ANs),
            new XAttribute("uniqueId", $"urn:officekit:smartart:v1:colors:{colorsId}"),
            new XElement(DgmNs + "title", new XAttribute("val", $"OfficeKit {colorsId}")),
            new XElement(DgmNs + "desc", new XAttribute("val", "OfficeKit clean-room SmartArt colors")),
            new XElement(DgmNs + "catLst", new XElement(DgmNs + "cat", new XAttribute("type", "accent"), new XAttribute("pri", 1))),
            new XElement(DgmNs + "styleLbl",
                new XAttribute("name", "node1"),
                new XElement(DgmNs + "fillClrLst", new XAttribute("meth", "repeat"), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(DgmNs + "linClrLst", new XAttribute("meth", "repeat"), new XElement(ANs + "schemeClr", new XAttribute("val", "lt1"))),
                new XElement(DgmNs + "effectClrLst"),
                new XElement(DgmNs + "txLinClrLst"),
                new XElement(DgmNs + "txFillClrLst", new XAttribute("meth", "repeat"), new XElement(ANs + "schemeClr", new XAttribute("val", "lt1"))),
                new XElement(DgmNs + "txEffectClrLst")));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Scene3D() => new(DgmNs + "scene3d",
        new XElement(ANs + "camera", new XAttribute("prst", "orthographicFront")),
        new XElement(ANs + "lightRig", new XAttribute("rig", "threePt"), new XAttribute("dir", "t")));

    private static XElement StyleLabel() => new(DgmNs + "styleLbl",
        new XAttribute("name", "node1"),
        Scene3D(),
        new XElement(DgmNs + "sp3d"),
        new XElement(DgmNs + "txPr"),
        new XElement(DgmNs + "style",
            new XElement(ANs + "lnRef", new XAttribute("idx", 2), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
            new XElement(ANs + "fillRef", new XAttribute("idx", 1), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
            new XElement(ANs + "effectRef", new XAttribute("idx", 0), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
            new XElement(ANs + "fontRef", new XAttribute("idx", "minor"), new XElement(ANs + "schemeClr", new XAttribute("val", "lt1")))));

    private static string BuildCachedDrawing(
        string elementId,
        PresentationDiagram diagram,
        IReadOnlyDictionary<string, string> imageRelationshipIds)
    {
        var shapeByName = diagram.Drawing.Children
            .Where(element => element.ContentCase == PresentationElement.ContentOneofCase.Shape)
            .GroupBy(element => element.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var shapeTree = new XElement(DspNs + "spTree",
            new XElement(DspNs + "nvGrpSpPr",
                new XElement(DspNs + "cNvPr", new XAttribute("id", 0), new XAttribute("name", string.Empty)),
                new XElement(DspNs + "cNvGrpSpPr")),
            new XElement(DspNs + "grpSpPr"));
        foreach (var node in diagram.Nodes)
        {
            if (!shapeByName.TryGetValue(node.Id, out var element)) continue;
            shapeTree.Add(CachedShape(
                elementId,
                node,
                element.Shape,
                diagram,
                imageRelationshipIds.GetValueOrDefault(node.Id)));
        }
        var root = new XElement(DspNs + "drawing",
            new XAttribute(XNamespace.Xmlns + "dsp", DspNs),
            new XAttribute(XNamespace.Xmlns + "a", ANs),
            new XAttribute(XNamespace.Xmlns + "dgm", DgmNs),
            new XAttribute(XNamespace.Xmlns + "r", RNs),
            shapeTree);
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement CachedShape(
        string elementId,
        PresentationDiagramNode node,
        PresentationShape shape,
        PresentationDiagram diagram,
        string? imageRelationshipId)
    {
        var x = shape.LeftEmu - diagram.LeftEmu;
        var y = shape.TopEmu - diagram.TopEmu;
        var geometry = string.IsNullOrWhiteSpace(shape.Geometry) || shape.Geometry == "custom" ? "rect" : shape.Geometry;
        var transform = new XElement(ANs + "xfrm",
            new XElement(ANs + "off", new XAttribute("x", x), new XAttribute("y", y)),
            new XElement(ANs + "ext", new XAttribute("cx", shape.WidthEmu), new XAttribute("cy", shape.HeightEmu)));
        var properties = new XElement(DspNs + "spPr",
            new XElement(transform),
            new XElement(ANs + "prstGeom", new XAttribute("prst", geometry), new XElement(ANs + "avLst")),
            imageRelationshipId is null ? SolidFill(shape) : ImageFill(imageRelationshipId, shape.ImageFill),
            new XElement(ANs + "ln", new XAttribute("w", Math.Max(1, shape.LineWidthEmu)),
                new XElement(ANs + "solidFill", new XElement(ANs + "schemeClr", new XAttribute("val", "lt1"))),
                new XElement(ANs + "prstDash", new XAttribute("val", "solid"))));
        var body = PptxTextCodec.BuildDrawingTextBody(node.TextBody ?? new PresentationTextBody());
        var textBody = new XElement(DspNs + "txBody", XElement.Parse(body.OuterXml).Elements());
        return new XElement(DspNs + "sp",
            new XAttribute("modelId", StableModelId(elementId, $"presentation:{node.Id}")),
            new XElement(DspNs + "nvSpPr",
                new XElement(DspNs + "cNvPr", new XAttribute("id", 0), new XAttribute("name", node.Id)),
                new XElement(DspNs + "cNvSpPr")),
            properties,
            new XElement(DspNs + "style",
                new XElement(ANs + "lnRef", new XAttribute("idx", 2), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(ANs + "fillRef", new XAttribute("idx", 1), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(ANs + "effectRef", new XAttribute("idx", 0), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(ANs + "fontRef", new XAttribute("idx", "minor"), new XElement(ANs + "schemeClr", new XAttribute("val", "lt1")))),
            textBody,
            new XElement(DspNs + "txXfrm", transform.Elements()));
    }

    private static XElement SolidFill(PresentationShape shape)
    {
        if (!string.IsNullOrWhiteSpace(shape.FillRgb))
            return new XElement(ANs + "solidFill", new XElement(ANs + "srgbClr", new XAttribute("val", shape.FillRgb)));
        return new XElement(ANs + "solidFill", new XElement(ANs + "schemeClr", new XAttribute("val", string.IsNullOrWhiteSpace(shape.FillScheme) ? "accent1" : shape.FillScheme)));
    }

    private static XElement ImageFill(string relationshipId, PresentationImagePaint? paint = null)
    {
        var blip = new XElement(ANs + "blip", new XAttribute(RNs + "embed", relationshipId));
        if (paint?.HasOpacityThousandthPercent == true)
            blip.Add(new XElement(ANs + "alphaModFix", new XAttribute("amt", paint.OpacityThousandthPercent)));
        var fill = new XElement(ANs + "blipFill", blip);
        if (paint?.Crop is { } crop)
            fill.Add(new XElement(ANs + "srcRect",
                new XAttribute("l", crop.LeftThousandthPercent),
                new XAttribute("t", crop.TopThousandthPercent),
                new XAttribute("r", crop.RightThousandthPercent),
                new XAttribute("b", crop.BottomThousandthPercent)));
        fill.Add(paint?.Mode == PresentationImagePaint.Types.Mode.Tile
            ? new XElement(ANs + "tile")
            : new XElement(ANs + "stretch", new XElement(ANs + "fillRect")));
        return fill;
    }

    private static IReadOnlyDictionary<string, string> AddCachedDrawingImages(
        PresentationDiagram diagram,
        DiagramPersistLayoutPart drawingPart,
        PptxPartContext slideContext)
    {
        var output = new Dictionary<string, string>(StringComparer.Ordinal);
        var relationshipByAsset = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var node in diagram.Nodes.Where(node => !string.IsNullOrWhiteSpace(node.AssetId)))
        {
            if (slideContext.Assets is null)
                throw new CodecException("invalid_presentation_asset", "SmartArt picture nodes require the presentation asset catalog.");
            if (!relationshipByAsset.TryGetValue(node.AssetId, out var relationshipId))
            {
                var asset = slideContext.Assets.Get(node.AssetId);
                relationshipId = $"rIdOfficeKitSmartArtImage{relationshipByAsset.Count + 1}";
                var part = drawingPart.AddImagePart(PptxAssetCatalog.ImagePartTypeFor(asset.ContentType), relationshipId);
                using var source = new MemoryStream(asset.Data.ToByteArray(), writable: false);
                part.FeedData(source);
                slideContext.TrackAddedPart(drawingPart, part);
                relationshipByAsset.Add(node.AssetId, relationshipId);
            }
            output.Add(node.Id, relationshipId);
        }
        return output;
    }

    private static PresentationGroup ReadCachedDrawing(
        PresentationDiagram diagram,
        DiagramPersistLayoutPart? part,
        PptxAssetCatalog assets,
        out bool verified)
    {
        var group = new PresentationGroup
        {
            LeftEmu = diagram.LeftEmu,
            TopEmu = diagram.TopEmu,
            WidthEmu = diagram.WidthEmu,
            HeightEmu = diagram.HeightEmu,
            ChildLeftEmu = diagram.LeftEmu,
            ChildTopEmu = diagram.TopEmu,
            ChildWidthEmu = diagram.WidthEmu,
            ChildHeightEmu = diagram.HeightEmu,
        };
        XDocument? drawing = null;
        try { if (part is not null) drawing = ReadXml(part); }
        catch (Exception exception) when (exception is IOException or System.Xml.XmlException) { }
        verified = drawing is not null;
        var drawingContext = part is null
            ? null
            : new PptxPartContext(
                part,
                new Dictionary<string, string>(StringComparer.Ordinal),
                assets: assets);
        foreach (var node in diagram.Nodes)
        {
            var native = drawing?.Descendants(DspNs + "sp").SingleOrDefault(shape =>
                shape.Element(DspNs + "nvSpPr")?.Element(DspNs + "cNvPr")?.Attribute("name")?.Value == node.Id);
            var blips = native?.Descendants(ANs + "blip").ToArray() ?? [];
            var nodeAssetVerified = true;
            if (blips.Length > 1)
            {
                verified = false;
                nodeAssetVerified = false;
                node.AssetId = string.Empty;
            }
            if (blips.Length == 1)
            {
                var embed = blips[0].Attribute(RNs + "embed")?.Value ?? string.Empty;
                var link = blips[0].Attribute(RNs + "link")?.Value ?? string.Empty;
                if (embed.Length == 0 || link.Length != 0 || part is null)
                {
                    verified = false;
                    nodeAssetVerified = false;
                    node.AssetId = string.Empty;
                }
                else
                {
                    try
                    {
                        if (part.GetPartById(embed) is not ImagePart imagePart)
                        {
                            verified = false;
                            nodeAssetVerified = false;
                            node.AssetId = string.Empty;
                        }
                        else
                        {
                            var importedAssetId = assets.Import(imagePart).Id;
                            if (!string.IsNullOrWhiteSpace(node.AssetId) &&
                                !node.AssetId.Equals(importedAssetId, StringComparison.Ordinal))
                            {
                                verified = false;
                                nodeAssetVerified = false;
                            }
                            else
                                node.AssetId = importedAssetId;
                            if (!nodeAssetVerified) node.AssetId = string.Empty;
                        }
                    }
                    catch (Exception exception) when (exception is ArgumentOutOfRangeException or CodecException)
                    {
                        verified = false;
                        nodeAssetVerified = false;
                        node.AssetId = string.Empty;
                    }
                }
            }
            else if (blips.Length == 0 && !string.IsNullOrWhiteSpace(node.AssetId))
            {
                verified = false;
                nodeAssetVerified = false;
                node.AssetId = string.Empty;
            }
            PresentationImagePaint? imagePaint = null;
            var blipFill = native?.Element(DspNs + "spPr")?.Element(ANs + "blipFill");
            if (blipFill is not null && drawingContext is not null && nodeAssetVerified)
            {
                try
                {
                    var typedBlipFill = new A.BlipFill(blipFill.ToString(SaveOptions.DisableFormatting));
                    if (PptxImagePaintCodec.TryRead(typedBlipFill, drawingContext, out var parsedPaint) &&
                        (string.IsNullOrWhiteSpace(node.AssetId) ||
                         node.AssetId.Equals(parsedPaint.AssetId, StringComparison.Ordinal)))
                    {
                        imagePaint = parsedPaint;
                        node.AssetId = parsedPaint.AssetId;
                    }
                }
                catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
                {
                    // The asset-only profile above may still be safe for a
                    // replacement, but a non-canonical blip graph must not
                    // receive the image-paint capability.
                }
            }
            var transform = native?.Element(DspNs + "spPr")?.Element(ANs + "xfrm");
            var nativeOffset = transform?.Element(ANs + "off");
            var nativeExtents = transform?.Element(ANs + "ext");
            if (native is null || nativeOffset is null || nativeExtents is null ||
                LongAttribute(nativeExtents, "cx") <= 0 || LongAttribute(nativeExtents, "cy") <= 0)
                verified = false;
            var shape = new PresentationShape
            {
                Geometry = native?.Element(DspNs + "spPr")?.Element(ANs + "prstGeom")?.Attribute("prst")?.Value ?? "rect",
                LeftEmu = diagram.LeftEmu + LongAttribute(nativeOffset, "x"),
                TopEmu = diagram.TopEmu + LongAttribute(nativeOffset, "y"),
                WidthEmu = Math.Max(1, LongAttribute(nativeExtents, "cx", diagram.WidthEmu)),
                HeightEmu = Math.Max(1, LongAttribute(nativeExtents, "cy", diagram.HeightEmu)),
                TextBody = node.TextBody?.Clone(),
                Text = PptxTextCodec.Flatten(node.TextBody),
                ImageFill = imagePaint,
            };
            group.Children.Add(new PresentationElement { Id = $"{node.Id}-cache", Name = node.Id, Shape = shape });
        }
        return group;
    }

    private static bool TryReadTextBody(XElement source, out PresentationTextBody body)
    {
        try
        {
            var drawingBody = new A.TextBody(new XElement(ANs + "txBody", source.Elements()).ToString(SaveOptions.DisableFormatting));
            body = PptxTextCodec.ReadDrawingTextBody(drawingBody);
            return true;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            body = null!;
            return false;
        }
    }

    private static XElement DiagramText(PresentationTextBody? body)
    {
        var native = PptxTextCodec.BuildDrawingTextBody(body ?? new PresentationTextBody());
        return new XElement(DgmNs + "t", XElement.Parse(native.OuterXml).Elements());
    }

    private static XElement EmptyDiagramText() => new(DgmNs + "t",
        new XElement(ANs + "bodyPr"),
        new XElement(ANs + "lstStyle"),
        new XElement(ANs + "p", new XElement(ANs + "endParaRPr", new XAttribute("lang", "en-US"))));

    private static XElement Extension(XElement payload) => new(DgmNs + "extLst",
        new XElement(ANs + "ext", new XAttribute("uri", OfficeKitExtension), payload));

    private static string? DataModelDrawingRelationshipId(DiagramDataPart part)
    {
        try
        {
            return ReadXml(part).Descendants(DspNs + "dataModelExt").SingleOrDefault()?.Attribute("relId")?.Value;
        }
        catch (Exception exception) when (exception is IOException or System.Xml.XmlException)
        {
            return null;
        }
    }

    private static XDocument ReadXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static long LongAttribute(XElement? element, string name, long fallback = 0) =>
        long.TryParse(element?.Attribute(name)?.Value, out var value) ? value : fallback;

    private static string LayoutId(string layout) => $"urn:officekit:smartart:v1:layout:{layout}";

    private static string StableModelId(string owner, string identity)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes($"officekit-smartart:{owner}:{identity}"))[..16];
        bytes[6] = (byte)((bytes[6] & 0x0F) | 0x50);
        bytes[8] = (byte)((bytes[8] & 0x3F) | 0x80);
        return new Guid(bytes).ToString("B").ToUpperInvariant();
    }
}
