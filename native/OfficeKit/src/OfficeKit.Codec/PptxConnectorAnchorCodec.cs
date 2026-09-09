using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Native geometry sites and PPJ frame references have different semantics.
// Keep frame references in their own bounded, optional extension instead of
// pretending that every shape has a matching site (especially for center).
internal static class PptxConnectorAnchorCodec
{
    internal const string Namespace = "urn:officekit:ppj:connector-frame-anchors:v1";
    private static readonly IReadOnlySet<string> Anchors = new HashSet<string>(StringComparer.Ordinal)
        { "auto", "top", "right", "bottom", "left", "center" };

    internal static bool TryRead(P.NonVisualConnectorShapeDrawingProperties source,
        IReadOnlyDictionary<uint, string>? ids, out PresentationConnectorFrameAnchor? start,
        out PresentationConnectorFrameAnchor? end)
    {
        start = end = null;
        var extensions = source.Elements<A.ExtensionList>().SelectMany(list => list.Elements<A.Extension>())
            .Where(extension => extension.Uri?.Value == Namespace).ToArray();
        if (extensions.Length == 0) return true;
        if (extensions.Length != 1 || source.Elements<A.ExtensionList>().Count() != 1) return false;
        var extension = extensions[0];
        if (extension.ExtendedAttributes.Any() || extension.ChildElements.Count != 1) return false;
        var root = extension.ChildElements[0];
        if (root.NamespaceUri != Namespace || root.LocalName != "anchors" || root.HasAttributes ||
            root.ChildElements.Count is < 1 or > 2) return false;
        foreach (var node in root.ChildElements)
        {
            if (node.NamespaceUri != Namespace || node.ChildElements.Count != 0 || node.InnerText.Length != 0 ||
                node.GetAttributes().Count != 2 || node.GetAttributes().Any(a => a.NamespaceUri.Length != 0 || a.LocalName is not ("target" or "anchor"))) return false;
            var nativeText = node.GetAttribute("target", "").Value;
            var anchor = node.GetAttribute("anchor", "").Value;
            if (!uint.TryParse(nativeText, NumberStyles.None, CultureInfo.InvariantCulture, out var nativeId) ||
                nativeText != nativeId.ToString(CultureInfo.InvariantCulture) ||
                !Anchors.Contains(anchor) || ids is null || !ids.TryGetValue(nativeId, out var target)) return false;
            var value = new PresentationConnectorFrameAnchor { TargetId = target, Anchor = anchor };
            if (node.LocalName == "start" && start is null && end is null) start = value;
            else if (node.LocalName == "end" && end is null) end = value;
            else return false;
        }
        return true;
    }

    internal static void Validate(PresentationConnectorFrameAnchor? anchor, string nativeTarget,
        string elementId, IReadOnlyDictionary<string, uint>? ids)
    {
        if (anchor is null) return;
        if (!Anchors.Contains(anchor.Anchor) || anchor.TargetId.Length is < 1 or > 1_024 ||
            anchor.TargetId == elementId || nativeTarget.Length != 0 || ids is not null && !ids.ContainsKey(anchor.TargetId))
            throw new CodecException("invalid_presentation_connector", $"Connector {elementId} has invalid or conflicting frame-anchor state.");
    }

    internal static void Apply(P.NonVisualConnectorShapeDrawingProperties properties, PresentationConnector connector,
        IReadOnlyDictionary<string, uint> ids)
    {
        var list = properties.GetFirstChild<A.ExtensionList>();
        if (list is not null)
        {
            foreach (var extension in list.Elements<A.Extension>().Where(e => e.Uri?.Value == Namespace).ToArray()) extension.Remove();
        }
        if (connector.StartFrameAnchor is null && connector.EndFrameAnchor is null)
        {
            if (list is { ChildElements.Count: 0 }) list.Remove();
            return;
        }
        if (list is null) properties.Append(list = new A.ExtensionList());
        var root = new OpenXmlUnknownElement("oka", "anchors", Namespace);
        foreach (var (name, value) in new[] { ("start", connector.StartFrameAnchor), ("end", connector.EndFrameAnchor) })
        {
            if (value is null) continue;
            var node = new OpenXmlUnknownElement("oka", name, Namespace);
            node.SetAttribute(new OpenXmlAttribute("target", "", ids[value.TargetId].ToString(CultureInfo.InvariantCulture)));
            node.SetAttribute(new OpenXmlAttribute("anchor", "", value.Anchor));
            root.Append(node);
        }
        list.Append(new A.Extension(root) { Uri = Namespace });
    }

    internal static bool References(P.ConnectionShape connector, uint nativeId)
    {
        // Also protect malformed/opaque instances of our extension from target
        // deletion. Unknown topology cannot establish that deletion is safe.
        foreach (var extension in connector.Descendants<A.Extension>().Where(e => e.Uri?.Value == Namespace))
        {
            var targets = extension.Descendants().Where(e => e.LocalName is "start" or "end").ToArray();
            if (targets.Length == 0) return true;
            foreach (var target in targets)
            {
                if (!uint.TryParse(target.GetAttribute("target", "").Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id == nativeId)
                    return true;
            }
        }
        return false;
    }
}
