using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxDiagramTextReplacement(string PartPath, string Sha256, byte[] Data);

internal sealed record PptxDiagramTextEditLeaf(
    uint TextLeafIndex,
    uint RawTextOrdinal,
    string ModelId,
    uint RunIndex,
    string Text);

internal sealed record PptxDiagramTextEditResolution(
    PresentationDiagramText Binding,
    DiagramDataPart Part,
    IReadOnlyList<PptxDiagramTextEditLeaf> Leaves);

// Owns one deliberately small SmartArt edit boundary. It does not author a
// diagram, change the graph, or reinterpret layout/style/colors. It exposes
// a semantic slice only where an imported top-level p:graphicFrame proves it
// owns four closed directly referenced Diagram definition parts. An optional
// cached drawing side graph, including images, remains source-owned and is not
// interpreted here. Content points must have one or more DrawingML paragraphs
// made only of direct plain runs and fixed line breaks. Parent-of connections
// between those points are exposed, while the document root, presentation
// points/edges, layout program, style, and colors remain source-owned. A node
// may retain source-owned empty paragraphs, but must contain at least one text
// run overall. Paragraph, run, and break topology remain owned by the source
// XML; the wire projects only the source-ordered text leaves. Everything
// outside that profile stays opaque.
internal static class PptxDiagramTextCodec
{
    private const string DiagramDataContentType = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    private const int MaxModelIdLength = 1_024;
    private const int MaxLayoutDefinitionIdLength = 2_048;
    private const int MaxNodeTextLength = 32_767;
    private const int MaxNodeParagraphCount = 256;
    private const int MaxNodeInlineCount = 256;
    private const int MaxPointCount = 1_024;
    private const int MaxConnectionCount = 4_096;
    private const int MaxProjectedNodeCount = 64;
    private const int MaxProjectedConnectionCount = 256;

    private static readonly HashSet<string> DiagramNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "http://purl.oclc.org/ooxml/drawingml/diagram",
    };

    private static readonly HashSet<string> DrawingNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    };

    private sealed record DiagramRun(string Text, XElement TextElement);
    private sealed record DiagramNode(string ModelId, string PointType, string Text, IReadOnlyList<DiagramRun> Runs);
    private sealed record DiagramConnection(string ModelId, string FromModelId, string ToModelId, uint Order);

    private sealed record ResolvedDiagram(
        PresentationDiagramText Binding,
        DiagramDataPart Part,
        XDocument Document,
        IReadOnlyList<DiagramNode> Nodes,
        IReadOnlyList<DiagramConnection> Connections);
    private sealed record DiagramParts(
        DiagramDataPart Data,
        DiagramLayoutDefinitionPart Layout,
        string DataRelationshipId);

    internal static bool TryDescribe(OpenXmlElement source, OpenXmlPart owner, out PresentationDiagramText binding)
    {
        if (TryResolve(source, owner, out var resolved))
        {
            binding = resolved.Binding;
            return true;
        }
        binding = null!;
        if (!TryResolveParts(source, owner, out var parts)) return false;
        try
        {
            var sourceBytes = ReadPart(parts.Data);
            if (!TryReadLayoutDefinitionId(parts.Layout, out var layoutDefinitionId)) return false;
            binding = new PresentationDiagramText
            {
                PartPath = PartPath(parts.Data),
                ContentType = parts.Data.ContentType,
                SourceSha256 = Hash(sourceBytes),
                RelationshipId = parts.DataRelationshipId,
                LayoutDefinitionId = layoutDefinitionId,
            };
            return true;
        }
        catch (Exception exception) when (exception is XmlException or IOException or UnauthorizedAccessException)
        {
            binding = null!;
            return false;
        }
    }

    internal static bool TryResolveForEditPlan(
        OpenXmlElement source,
        OpenXmlPart owner,
        out PptxDiagramTextEditResolution resolved)
    {
        resolved = null!;
        if (!TryResolve(source, owner, out var diagram)) return false;
        var rawOrdinals = diagram.Document.Descendants()
            .Where(element => IsDrawing(element, "t"))
            .Select((element, index) => (Element: element, Ordinal: checked((uint)index)))
            .ToDictionary(item => item.Element, item => item.Ordinal);
        var leaves = new List<PptxDiagramTextEditLeaf>();
        uint textLeafIndex = 0;
        foreach (var node in diagram.Nodes)
        {
            for (var runIndex = 0; runIndex < node.Runs.Count; runIndex++)
            {
                var run = node.Runs[runIndex];
                if (!rawOrdinals.TryGetValue(run.TextElement, out var rawTextOrdinal)) return false;
                leaves.Add(new PptxDiagramTextEditLeaf(
                    textLeafIndex++,
                    rawTextOrdinal,
                    node.ModelId,
                    checked((uint)runIndex),
                    run.Text));
            }
        }
        if (leaves.Count == 0) return false;
        resolved = new PptxDiagramTextEditResolution(diagram.Binding, diagram.Part, leaves);
        return true;
    }

    internal static PptxDiagramTextReplacement? PrepareReplacement(
        SlidePart owner,
        OpenXmlElement source,
        PresentationOpaqueElement original,
        PresentationOpaqueElement requested)
    {
        if (original.DiagramText is null)
        {
            if (requested.DiagramText is not null)
                throw Unsupported("An unrecognized SmartArt graph cannot claim the bounded diagram-text edit capability.");
            return null;
        }
        if (requested.DiagramText is null)
            throw Unsupported("A source-bound SmartArt text binding cannot be removed.");
        if (IsReadOnlyDescription(original.DiagramText))
        {
            if (!original.DiagramText.Equals(requested.DiagramText))
                throw Unsupported("This SmartArt graph is typed for preservation but has no proven editable text or graph leaves.");
            return null;
        }
        if (!TryResolve(source, owner, out var resolved))
            throw BindingMismatch("The SmartArt source no longer proves the bounded diagram-text profile.", PartPath(owner));
        if (!SameBinding(original.DiagramText, resolved.Binding) ||
            !SameGraphIdentity(original.DiagramText, resolved.Binding) ||
            !SameNodes(original.DiagramText.Nodes, resolved.Nodes))
            throw BindingMismatch("The SmartArt diagram data no longer matches its source binding.", resolved.Binding.PartPath);
        ValidateRequestedGraph(original.DiagramText, requested.DiagramText, resolved.Binding.PartPath);

        var changed = false;
        for (var index = 0; index < resolved.Nodes.Count; index++)
        {
            var requestedRuns = RequestedRunTexts(
                original.DiagramText.Nodes[index],
                requested.DiagramText.Nodes[index],
                resolved.Binding.PartPath);
            if (!requested.DiagramText.Nodes[index].RunTexts.SequenceEqual(requestedRuns))
            {
                requested.DiagramText.Nodes[index].RunTexts.Clear();
                requested.DiagramText.Nodes[index].RunTexts.Add(requestedRuns);
            }
            for (var runIndex = 0; runIndex < resolved.Nodes[index].Runs.Count; runIndex++)
            {
                if (resolved.Nodes[index].Runs[runIndex].Text == requestedRuns[runIndex]) continue;
                SetText(resolved.Nodes[index].Runs[runIndex].TextElement, requestedRuns[runIndex]);
                changed = true;
            }
        }
        if (!changed) return null;

        var data = Serialize(resolved.Document);
        return new PptxDiagramTextReplacement(resolved.Binding.PartPath, Hash(data), data);
    }

    internal static void Apply(
        SlidePart owner,
        PresentationDiagramText binding,
        PptxDiagramTextReplacement replacement)
    {
        DiagramDataPart part;
        try
        {
            part = owner.GetPartById(binding.RelationshipId) as DiagramDataPart
                ?? throw BindingMismatch("The SmartArt data relationship no longer resolves to a DiagramDataPart.", PartPath(owner));
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw BindingMismatch("The SmartArt data relationship no longer resolves to a package part.", PartPath(owner), exception);
        }
        var partPath = PartPath(part);
        if (!partPath.Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !partPath.Equals(replacement.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(binding.ContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(DiagramDataContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/diagramData", StringComparison.Ordinal))
            throw BindingMismatch("The SmartArt data part path, content type, or relationship type no longer matches its source binding.", partPath);
        if (!Hash(ReadPart(part)).Equals(binding.SourceSha256, StringComparison.OrdinalIgnoreCase))
            throw BindingMismatch("The SmartArt data bytes no longer match their source digest.", partPath);

        using var output = part.GetStream(FileMode.Create, FileAccess.Write);
        output.Write(replacement.Data);
    }

    internal static void ValidateSourceBoundOutput(
        SlidePart sourceOwner,
        SlidePart outputOwner,
        OpenXmlElement source,
        OpenXmlElement output,
        PresentationOpaqueElement requested)
    {
        if (requested.DiagramText is null) return;
        if (IsReadOnlyDescription(requested.DiagramText))
        {
            if (!TryDescribe(source, sourceOwner, out var sourceDescription) ||
                !SameEditBinding(requested.DiagramText, sourceDescription) ||
                !TryDescribe(output, outputOwner, out var outputDescription) ||
                !SameEditBinding(requested.DiagramText, outputDescription))
                throw BindingMismatch("The preserved SmartArt graph no longer matches its typed read-only binding.", PartPath(outputOwner));
            return;
        }
        if (!TryResolve(source, sourceOwner, out var sourceResolved) ||
            !SameBinding(requested.DiagramText, sourceResolved.Binding) ||
            !SameGraphIdentity(requested.DiagramText, sourceResolved.Binding))
            throw BindingMismatch("The source SmartArt diagram text does not match the requested source binding.", PartPath(sourceOwner));
        if (!TryResolve(output, outputOwner, out var outputResolved))
            throw BindingMismatch("The exported SmartArt diagram no longer proves the bounded diagram-text profile.", PartPath(outputOwner));
        if (!outputResolved.Binding.PartPath.Equals(requested.DiagramText.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !outputResolved.Binding.ContentType.Equals(requested.DiagramText.ContentType, StringComparison.OrdinalIgnoreCase) ||
            outputResolved.Binding.RelationshipId != requested.DiagramText.RelationshipId ||
            !SameGraphIdentity(requested.DiagramText, outputResolved.Binding) ||
            !SameRequestedText(requested.DiagramText.Nodes, outputResolved.Nodes))
            throw BindingMismatch("The exported SmartArt node text does not match the requested bounded edit.", outputResolved.Binding.PartPath);
    }

    private static bool TryResolve(OpenXmlElement source, OpenXmlPart owner, out ResolvedDiagram resolved)
    {
        resolved = null!;
        if (!TryResolveParts(source, owner, out var parts)) return false;
        var dataPart = parts.Data;
        var layoutPart = parts.Layout;

        byte[] sourceBytes;
        XDocument document;
        IReadOnlyList<DiagramNode> nodes;
        IReadOnlyList<DiagramConnection> connections;
        string layoutDefinitionId;
        try
        {
            sourceBytes = ReadPart(dataPart);
            using var memory = new MemoryStream(sourceBytes, writable: false);
            using var reader = XmlReader.Create(memory, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                IgnoreWhitespace = false,
            });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            if (!TryReadGraph(document, out nodes, out connections) ||
                !TryReadLayoutDefinitionId(layoutPart, out layoutDefinitionId)) return false;
        }
        catch (Exception exception) when (exception is XmlException or IOException or UnauthorizedAccessException)
        {
            return false;
        }

        var binding = new PresentationDiagramText
        {
            PartPath = PartPath(dataPart),
            ContentType = dataPart.ContentType,
            SourceSha256 = Hash(sourceBytes),
            RelationshipId = parts.DataRelationshipId,
            LayoutDefinitionId = layoutDefinitionId,
        };
        binding.Nodes.Add(nodes.Select(ToWireNode));
        binding.Connections.Add(connections.Select(ToWireConnection));
        resolved = new ResolvedDiagram(binding, dataPart, document, nodes, connections);
        return true;
    }

    private static bool TryResolveParts(OpenXmlElement source, OpenXmlPart owner, out DiagramParts parts)
    {
        parts = null!;
        if (source is not P.GraphicFrame frame || owner is not SlidePart || source.Parent is not P.ShapeTree ||
            PptxNativeObjectCatalog.Classify(source) != "diagram" || !PptxNativeObjectCatalog.SupportsPlacementEditing(source))
            return false;
        var roots = frame.Descendants().Where(PptxNativeObjectCatalog.IsDiagramRelationshipIds).ToArray();
        if (roots.Length != 1) return false;
        var relationshipAttributes = new[] { (OpenXmlElement)frame }.Concat(frame.Descendants())
            .SelectMany(element => element.GetAttributes())
            .Where(attribute => PptxNativeObjectCatalog.IsRelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        var rootAttributes = roots[0].GetAttributes()
            .Where(attribute => PptxNativeObjectCatalog.IsRelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        if (relationshipAttributes.Length != 4 || rootAttributes.Length != 4) return false;

        var expected = new Dictionary<string, Type>(StringComparer.Ordinal)
        {
            ["dm"] = typeof(DiagramDataPart),
            ["lo"] = typeof(DiagramLayoutDefinitionPart),
            ["qs"] = typeof(DiagramStylePart),
            ["cs"] = typeof(DiagramColorsPart),
        };
        var relationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var partPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        DiagramDataPart? dataPart = null;
        DiagramLayoutDefinitionPart? layoutPart = null;
        string dataRelationshipId = string.Empty;
        foreach (var attribute in rootAttributes)
        {
            if (!expected.TryGetValue(attribute.LocalName, out var expectedType) || string.IsNullOrWhiteSpace(attribute.Value) ||
                !relationshipIds.Add(attribute.Value)) return false;
            OpenXmlPart part;
            try
            {
                part = owner.GetPartById(attribute.Value);
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
            if (part.GetType() != expectedType || !IsClosedDiagramPart(part) || !partPaths.Add(PartPath(part))) return false;
            if (attribute.LocalName == "dm")
            {
                dataPart = (DiagramDataPart)part;
                dataRelationshipId = attribute.Value;
            }
            else if (attribute.LocalName == "lo")
            {
                layoutPart = (DiagramLayoutDefinitionPart)part;
            }
        }
        if (dataPart is null || layoutPart is null || relationshipIds.Count != 4 || partPaths.Count != 4 ||
            expected.Count != rootAttributes.Select(attribute => attribute.LocalName).Distinct(StringComparer.Ordinal).Count())
            return false;

        parts = new(dataPart, layoutPart, dataRelationshipId);
        return true;
    }

    private static bool TryReadGraph(
        XDocument document,
        out IReadOnlyList<DiagramNode> nodes,
        out IReadOnlyList<DiagramConnection> connections)
    {
        nodes = [];
        connections = [];
        var root = document.Root;
        if (root is null || root.Name.LocalName != "dataModel" || !DiagramNamespaces.Contains(root.Name.NamespaceName)) return false;
        var pointLists = root.Elements().Where(element => IsDiagram(element, "ptLst")).ToArray();
        var connectionLists = root.Elements().Where(element => IsDiagram(element, "cxnLst")).ToArray();
        if (pointLists.Length != 1 || connectionLists.Length != 1) return false;
        var points = pointLists[0].Elements().Where(element => IsDiagram(element, "pt")).ToArray();
        var sourceConnections = connectionLists[0].Elements().Where(element => IsDiagram(element, "cxn")).ToArray();
        if (points.Length is < 1 or > MaxPointCount || sourceConnections.Length > MaxConnectionCount ||
            pointLists[0].Elements().Count() != points.Length || connectionLists[0].Elements().Count() != sourceConnections.Length)
            return false;

        var results = new List<DiagramNode>();
        var pointTypes = new Dictionary<string, string>(StringComparer.Ordinal);
        var modelIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var point in points)
        {
            var modelId = point.Attribute("modelId")?.Value ?? string.Empty;
            var pointType = NormalizePointType(point.Attribute("type")?.Value);
            if (!IsBoundedModelId(modelId) || pointType is null || !modelIds.Add(modelId) || !pointTypes.TryAdd(modelId, pointType)) return false;
            var textBodies = point.Elements().Where(element => IsDiagram(element, "t")).ToArray();
            if (pointType == "doc" && textBodies.Length == 0) continue;
            if (pointType is not "node" and not "asst" and not "doc") continue;
            if (textBodies.Length != 1 || !TryReadPlainRuns(textBodies[0], out var text, out var runs)) return false;
            results.Add(new DiagramNode(modelId, pointType, text, runs));
        }
        if (results.Count is < 1 or > MaxProjectedNodeCount) return false;

        var contentIds = results.Select(node => node.ModelId).ToHashSet(StringComparer.Ordinal);
        var documentIds = pointTypes.Where(pair => pair.Value == "doc").Select(pair => pair.Key).ToHashSet(StringComparer.Ordinal);
        var graph = new List<DiagramConnection>();
        foreach (var connection in sourceConnections)
        {
            var modelId = connection.Attribute("modelId")?.Value ?? string.Empty;
            var type = connection.Attribute("type")?.Value;
            var fromId = connection.Attribute("srcId")?.Value ?? string.Empty;
            var toId = connection.Attribute("destId")?.Value ?? string.Empty;
            if (!IsBoundedModelId(modelId) || !modelIds.Add(modelId) ||
                !pointTypes.ContainsKey(fromId) || !pointTypes.ContainsKey(toId)) return false;

            // Missing type is the schema default, parent-of. Presentation
            // relationships drive cached drawing/layout generation, not the
            // user's content graph, so they stay source-private.
            if (string.IsNullOrEmpty(type) || type == "parOf")
            {
                if (documentIds.Contains(fromId) && contentIds.Contains(toId)) continue;
                if (!contentIds.Contains(fromId) || !contentIds.Contains(toId) ||
                    !TryReadOrder(connection.Attribute("srcOrd")?.Value, out var order)) return false;
                graph.Add(new DiagramConnection(modelId, fromId, toId, order));
                if (graph.Count > MaxProjectedConnectionCount) return false;
                continue;
            }
            if (type is "presOf" or "presParOf") continue;
            return false;
        }
        nodes = results;
        connections = graph;
        return true;
    }

    private static string? NormalizePointType(string? value) => value switch
    {
        null or "" or "node" => "node",
        "asst" => "asst",
        "doc" => "doc",
        "pres" => "pres",
        "parTrans" => "parTrans",
        "sibTrans" => "sibTrans",
        _ => null,
    };

    private static bool TryReadOrder(string? value, out uint order) =>
        uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out order);

    private static bool TryReadLayoutDefinitionId(DiagramLayoutDefinitionPart part, out string layoutDefinitionId)
    {
        layoutDefinitionId = string.Empty;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = false,
            IgnoreProcessingInstructions = false,
            IgnoreWhitespace = false,
        });
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        var root = document.Root;
        if (root is null || !IsDiagram(root, "layoutDef")) return false;
        var value = root.Attribute("uniqueId")?.Value ?? string.Empty;
        if (!IsBoundedIdentifier(value, MaxLayoutDefinitionIdLength)) return false;
        layoutDefinitionId = value;
        return true;
    }

    private static bool TryReadPlainRuns(XElement body, out string text, out IReadOnlyList<DiagramRun> resolvedRuns)
    {
        text = string.Empty;
        resolvedRuns = [];
        var bodyChildren = body.Elements().ToArray();
        if (bodyChildren.Any(element => !IsDrawing(element, "bodyPr") && !IsDrawing(element, "lstStyle") && !IsDrawing(element, "p"))) return false;
        var paragraphs = bodyChildren.Where(element => IsDrawing(element, "p")).ToArray();
        if (paragraphs.Length is < 1 or > MaxNodeParagraphCount) return false;
        var results = new List<DiagramRun>();
        var combined = new StringBuilder();
        var inlineCount = 0;
        foreach (var paragraph in paragraphs)
        {
            var paragraphChildren = paragraph.Elements().ToArray();
            if (paragraphChildren.Any(element => !IsDrawing(element, "pPr") && !IsDrawing(element, "r") && !IsDrawing(element, "br") && !IsDrawing(element, "endParaRPr"))) return false;
            var runs = paragraphChildren.Where(element => IsDrawing(element, "r")).ToArray();
            var breaks = paragraphChildren.Where(element => IsDrawing(element, "br")).ToArray();
            if (breaks.Any(element => !IsCanonicalBreak(element)) ||
                inlineCount + runs.Length + breaks.Length > MaxNodeInlineCount) return false;
            inlineCount += runs.Length + breaks.Length;
            foreach (var run in runs)
            {
                var runChildren = run.Elements().ToArray();
                if (runChildren.Any(element => !IsDrawing(element, "rPr") && !IsDrawing(element, "t"))) return false;
                var textElements = runChildren.Where(element => IsDrawing(element, "t")).ToArray();
                // Replacing XElement.Value would discard comments or processing
                // instructions nested in a:t. They are not part of the plain-text
                // profile, so withhold the capability rather than silently erasing
                // source-owned markup.
                if (textElements.Length != 1 || textElements[0].HasElements || textElements[0].Nodes().Any(node => node is not XText)) return false;
                var runText = textElements[0].Value;
                if (!IsBoundedText(runText)) return false;
                combined.Append(runText);
                if (combined.Length > MaxNodeTextLength) return false;
                results.Add(new DiagramRun(runText, textElements[0]));
            }
        }
        if (results.Count == 0) return false;
        text = combined.ToString();
        resolvedRuns = results;
        return true;
    }

    private static bool IsCanonicalBreak(XElement element)
    {
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
        var children = element.Elements().ToArray();
        if (children.Length > 1 || children.Any(child => !IsDrawing(child, "rPr"))) return false;
        return element.Nodes().All(node => node switch
        {
            XElement child => IsDrawing(child, "rPr"),
            XText text => string.IsNullOrWhiteSpace(text.Value),
            _ => false,
        });
    }

    private static void ValidateRequestedGraph(
        PresentationDiagramText original,
        PresentationDiagramText requested,
        string partPath)
    {
        if (original.LayoutDefinitionId != requested.LayoutDefinitionId ||
            !SameConnections(original.Connections, requested.Connections))
            throw Unsupported("SmartArt layout identity and connection topology are source-bound and cannot be changed.", partPath);
        var originalNodes = original.Nodes;
        var requestedNodes = requested.Nodes;
        if (originalNodes.Count != requestedNodes.Count)
            throw Unsupported("SmartArt node topology is source-bound and cannot be changed.", partPath);
        for (var index = 0; index < originalNodes.Count; index++)
        {
            if (originalNodes[index].ModelId != requestedNodes[index].ModelId ||
                originalNodes[index].PointType != requestedNodes[index].PointType)
                throw Unsupported("SmartArt node identifiers are source-bound and cannot be changed.", partPath);
            var requestedRuns = RequestedRunTexts(originalNodes[index], requestedNodes[index], partPath);
            if (requestedRuns.Count != originalNodes[index].RunTexts.Count)
                throw Unsupported("SmartArt run topology is source-bound and cannot be changed.", partPath);
            if (requestedRuns.Any(text => !IsBoundedText(text)) ||
                !IsBoundedText(requestedNodes[index].Text) ||
                string.Concat(requestedRuns) != requestedNodes[index].Text)
                throw Unsupported($"SmartArt node {requestedNodes[index].ModelId} text is outside the bounded plain-run profile.", partPath);
        }
    }

    private static IReadOnlyList<string> RequestedRunTexts(
        PresentationDiagramTextNode original,
        PresentationDiagramTextNode requested,
        string partPath)
    {
        if (requested.RunTexts.Count > 0)
        {
            // Protocol-v2 callers predating run_texts mutate only text. Keep
            // that spelling valid for the old one-run profile without letting
            // it guess how a multi-run node should be divided.
            if (original.RunTexts.Count == 1 && requested.RunTexts.Count == 1 &&
                requested.RunTexts[0] == original.RunTexts[0] && requested.Text != original.Text)
                return [requested.Text];
            return requested.RunTexts;
        }
        if (original.RunTexts.Count == 1) return [requested.Text];
        throw Unsupported("A multi-run SmartArt node must retain its complete source-bound run topology.", partPath);
    }

    private static PresentationDiagramTextNode ToWireNode(DiagramNode node)
    {
        var result = new PresentationDiagramTextNode { ModelId = node.ModelId, PointType = node.PointType, Text = node.Text };
        result.RunTexts.Add(node.Runs.Select(run => run.Text));
        return result;
    }

    private static PresentationDiagramTextConnection ToWireConnection(DiagramConnection connection) => new()
    {
        ModelId = connection.ModelId,
        FromModelId = connection.FromModelId,
        ToModelId = connection.ToModelId,
        Order = connection.Order,
    };

    internal static bool SameEditBinding(PresentationDiagramText expected, PresentationDiagramText actual) =>
        SameBinding(expected, actual) &&
        SameGraphIdentity(expected, actual) &&
        expected.Nodes.Count == actual.Nodes.Count &&
        expected.Nodes.Select((node, index) =>
            node.ModelId == actual.Nodes[index].ModelId &&
            node.PointType == actual.Nodes[index].PointType &&
            node.Text == actual.Nodes[index].Text &&
            node.RunTexts.SequenceEqual(actual.Nodes[index].RunTexts)).All(match => match);

    private static bool SameBinding(PresentationDiagramText expected, PresentationDiagramText actual) =>
        expected.PartPath.Equals(actual.PartPath, StringComparison.OrdinalIgnoreCase) &&
        expected.ContentType.Equals(actual.ContentType, StringComparison.OrdinalIgnoreCase) &&
        expected.SourceSha256.Equals(actual.SourceSha256, StringComparison.OrdinalIgnoreCase) &&
        expected.RelationshipId == actual.RelationshipId;

    private static bool SameNodes(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> expected,
        IReadOnlyList<DiagramNode> actual) =>
        expected.Count == actual.Count && expected.Select((node, index) =>
            node.ModelId == actual[index].ModelId && node.Text == actual[index].Text &&
            node.PointType == actual[index].PointType &&
            SameRunTexts(node, actual[index])).All(match => match);

    private static bool SameRequestedText(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> expected,
        IReadOnlyList<DiagramNode> actual) =>
        expected.Count == actual.Count && expected.Select((node, index) =>
            node.ModelId == actual[index].ModelId && node.PointType == actual[index].PointType && node.Text == actual[index].Text &&
            SameRunTexts(node, actual[index])).All(match => match);

    private static bool SameGraphIdentity(PresentationDiagramText expected, PresentationDiagramText actual) =>
        expected.LayoutDefinitionId == actual.LayoutDefinitionId &&
        expected.Nodes.Count == actual.Nodes.Count &&
        expected.Nodes.Select((node, index) =>
            node.ModelId == actual.Nodes[index].ModelId &&
            node.PointType == actual.Nodes[index].PointType &&
            (node.RunTexts.Count == 0 ? actual.Nodes[index].RunTexts.Count == 1 : node.RunTexts.Count == actual.Nodes[index].RunTexts.Count)).All(match => match) &&
        SameConnections(expected.Connections, actual.Connections);

    private static bool SameConnections(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextConnection> expected,
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextConnection> actual) =>
        expected.Count == actual.Count && expected.Select((connection, index) =>
            connection.ModelId == actual[index].ModelId &&
            connection.FromModelId == actual[index].FromModelId &&
            connection.ToModelId == actual[index].ToModelId &&
            connection.Order == actual[index].Order).All(match => match);

    private static bool SameRunTexts(PresentationDiagramTextNode expected, DiagramNode actual)
    {
        if (expected.RunTexts.Count == 0)
            return actual.Runs.Count == 1 && expected.Text == actual.Runs[0].Text;
        return expected.RunTexts.Count == actual.Runs.Count &&
            expected.RunTexts.Select((text, index) => text == actual.Runs[index].Text).All(match => match);
    }

    private static bool IsReadOnlyDescription(PresentationDiagramText binding) =>
        binding.Nodes.Count == 0 && binding.Connections.Count == 0;

    private static bool IsClosedDiagramPart(OpenXmlPart part) =>
        part.Parts.Any() == false && !part.ExternalRelationships.Any() && !part.HyperlinkRelationships.Any() && !part.DataPartReferenceRelationships.Any();

    private static bool IsDiagram(XElement element, string localName) =>
        element.Name.LocalName == localName && DiagramNamespaces.Contains(element.Name.NamespaceName);

    private static bool IsDrawing(XElement element, string localName) =>
        element.Name.LocalName == localName && DrawingNamespaces.Contains(element.Name.NamespaceName);

    private static bool IsBoundedModelId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaxModelIdLength || value.Any(char.IsControl)) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            // ST_ModelId is the ECMA-376 union of xsd:int and a:ST_Guid.
            // Accepting arbitrary XML-safe strings here would expose a write
            // capability for source packages that the Open XML validator must
            // later reject. Keep the import capability at the same boundary.
            return int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _) ||
                Guid.TryParseExact(value, "B", out _);
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool IsBoundedIdentifier(string value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > maxLength || value.Any(char.IsControl)) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool IsBoundedText(string value)
    {
        if (value.Length > MaxNodeTextLength || value.Any(character => char.IsControl(character) && character is not '\t' and not '\n' and not '\r')) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static void SetText(XElement element, string value)
    {
        element.Value = value;
        var preserveWhitespace = value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        element.SetAttributeValue(XNamespace.Xml + "space", preserveWhitespace ? "preserve" : null);
    }

    private static byte[] Serialize(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = document.Declaration is null,
        }))
        {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    private static byte[] ReadPart(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    private static CodecException Unsupported(string message, string? partPath = null) => new("unsupported_presentation_edit", message, partPath);
    private static CodecException BindingMismatch(string message, string? partPath = null, Exception? innerException = null) =>
        new("presentation_diagram_text_binding_mismatch", message, partPath, innerException);
}
