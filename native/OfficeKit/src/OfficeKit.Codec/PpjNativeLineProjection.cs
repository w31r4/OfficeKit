using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OfficeKit.Codec;

/// <summary>
/// Projects the small, source-bound line style surface that the presentation
/// object model already knows how to prove and token-splice.  This is used
/// only when a connector could not be promoted to the typed PPJ connector;
/// the connector topology and every other native child remain opaque.
/// </summary>
internal static class PpjNativeLineProjection
{
    private const long MaxWidthEmu = 20_116_800;
    private const int MaxLeaves = 32;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    internal sealed record Leaf(string Kind, string Value);

    internal static bool TryRead(string rawXml, out IReadOnlyList<Leaf> leaves)
    {
        leaves = [];
        if (string.IsNullOrWhiteSpace(rawXml)) return false;

        XElement root;
        try
        {
            using var text = new StringReader(rawXml);
            using var reader = XmlReader.Create(text, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 16 * 1024 * 1024,
            });
            root = XElement.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is XmlException or InvalidOperationException)
        {
            return false;
        }

        if (root.Name != P + "cxnSp") return false;
        var shapeProperties = root.Elements(P + "spPr").ToArray();
        if (shapeProperties.Length != 1) return false;
        var outlines = shapeProperties[0].Elements(A + "ln").ToArray();
        if (outlines.Length != 1) return false;
        var line = outlines[0];

        var result = new List<Leaf>();
        if (TryAttribute(line, "w", out var width) && TryCanonicalWidth(width, out var canonicalWidth))
            result.Add(new("lineWidthEmu", canonicalWidth));

        var solidFills = line.Elements(A + "solidFill").ToArray();
        Leaf? colorLeaf = null;
        var hasSimplePaint = solidFills.Length == 1 && TryReadColor(solidFills[0], out colorLeaf);
        if (hasSimplePaint) result.Add(colorLeaf!);

        var dashes = line.Elements(A + "prstDash").ToArray();
        if (dashes.Length == 1 && hasSimplePaint &&
            TryAttribute(dashes[0], "val", out var dash) &&
            PptxLineStyleCodec.TryReadPresetDashValue(dash, out var style))
            result.Add(new("lineStyle", style));

        if (TryAttribute(line, "cap", out var cap) && hasSimplePaint &&
            PptxLineStyleCodec.TryReadCapValue(cap, out var capValue))
            result.Add(new("lineCap", capValue));

        var joins = line.Elements().Where(child => child.Name is { } name &&
            (name == A + "round" || name == A + "bevel" || name == A + "miter"))
            .ToArray();
        if (joins.Length == 1 && hasSimplePaint && IsBare(joins[0]))
        {
            var join = joins[0].Name.LocalName;
            result.Add(new("lineJoin", join));
        }

        var hasStyleReference = HasCanonicalStyleLineReference(root);
        if (hasSimplePaint || hasStyleReference)
        {
            if (TryReadArrow(line, A + "headEnd", "lineStartArrow", out var startArrow)) result.Add(startArrow!);
            if (TryReadArrow(line, A + "tailEnd", "lineEndArrow", out var endArrow)) result.Add(endArrow!);
        }

        if (result.Count == 0 || result.Count > MaxLeaves) return false;
        leaves = result;
        return true;
    }

    private static bool TryReadColor(XElement solidFill, out Leaf? leaf)
    {
        leaf = null;
        var colors = solidFill.Elements().Where(child => child.Name == A + "srgbClr" || child.Name == A + "schemeClr").ToArray();
        if (colors.Length != 1 || !IsBareColor(colors[0]) || !TryAttribute(colors[0], "val", out var value)) return false;
        if (colors[0].Name == A + "srgbClr" && value.Length == 6 && value.All(Uri.IsHexDigit))
        {
            leaf = new("lineRgb", value.ToUpperInvariant());
            return true;
        }
        if (colors[0].Name == A + "schemeClr" && PptxColor.TrySchemeToken(value, out var token))
        {
            leaf = new("lineScheme", token);
            return true;
        }
        return false;
    }

    private static bool TryReadArrow(XElement line, XName name, string kind, out Leaf? leaf)
    {
        leaf = null;
        var endpoints = line.Elements(name).ToArray();
        if (endpoints.Length != 1 || endpoints[0].HasElements) return false;
        var endpoint = endpoints[0];
        var attributes = endpoint.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Any(attribute => attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("type" or "w" or "len")) ||
            attributes.Count(attribute => attribute.Name.LocalName == "type") != 1 ||
            attributes.Count(attribute => attribute.Name.LocalName == "w") > 1 ||
            attributes.Count(attribute => attribute.Name.LocalName == "len") > 1 ||
            attributes.Where(attribute => attribute.Name.LocalName is "w" or "len").Any(attribute => attribute.Value is not ("sm" or "med" or "lg")) ||
            !TryAttribute(endpoint, "type", out var type)) return false;
        var parsed = type.Trim().ToLowerInvariant() switch
        {
            "none" => "none",
            "triangle" => "triangle",
            "stealth" => "stealth",
            "diamond" => "diamond",
            "oval" => "oval",
            "arrow" => "arrow",
            _ => string.Empty,
        };
        if (parsed.Length == 0) return false;
        leaf = new(kind, parsed);
        return true;
    }

    private static bool TryCanonicalWidth(string value, out string canonical)
    {
        canonical = string.Empty;
        if (value.Length == 0 || value.Any(character => character is < '0' or > '9') ||
            !ulong.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) || parsed > MaxWidthEmu)
            return false;
        canonical = parsed.ToString(CultureInfo.InvariantCulture);
        return value == canonical;
    }

    private static bool TryAttribute(XElement element, string name, out string value)
    {
        var attributes = element.Attributes().Where(attribute => attribute.Name.LocalName == name && attribute.Name.Namespace == XNamespace.None).ToArray();
        value = attributes.Length == 1 ? attributes[0].Value : string.Empty;
        return attributes.Length == 1;
    }

    private static bool IsBare(XElement element) => !element.HasElements && element.Attributes().All(attribute => attribute.IsNamespaceDeclaration == false);

    private static bool IsBareColor(XElement element) => IsBare(element);

    private static bool HasCanonicalStyleLineReference(XElement root)
    {
        var styles = root.Elements(P + "style").ToArray();
        if (styles.Length != 1 || styles[0].Attributes().Any() || styles[0].Elements().Count() != 4) return false;
        var names = styles[0].Elements().Select(element => element.Name).ToArray();
        var expected = new[] { A + "lnRef", A + "fillRef", A + "effectRef", A + "fontRef" };
        if (names.Distinct().Count() != 4 || expected.Any(name => !names.Contains(name))) return false;
        return styles[0].Elements().All(HasCanonicalStyleReference);
    }

    private static bool HasCanonicalStyleReference(XElement reference)
    {
        if (reference.Attributes().Any() || reference.Elements().Count() != 1) return false;
        var index = reference.Attribute("idx")?.Value;
        if (index is null) return false;
        if (reference.Name == A + "fontRef")
        {
            if (index is not ("minor" or "major")) return false;
        }
        else if (!uint.TryParse(index, NumberStyles.None, CultureInfo.InvariantCulture, out var numeric) || numeric > 32)
            return false;
        var color = reference.Elements().Single();
        return IsBareColor(color) && TryAttribute(color, "val", out var value) &&
            (color.Name == A + "schemeClr" && PptxColor.TrySchemeToken(value, out _) ||
             color.Name == A + "srgbClr" && value.Length == 6 && value.All(Uri.IsHexDigit));
    }
}
