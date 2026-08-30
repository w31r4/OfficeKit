using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using D14 = DocumentFormat.OpenXml.Office2010.Drawing;
using OM = DocumentFormat.OpenXml.Math;

namespace OfficeKit.Codec;

// Owns the finite canonical OMML profile emitted for PPJ formula runs. It is
// deliberately smaller than OMML: unknown native equations remain opaque and
// are never reverse-engineered into guessed LaTeX.
internal static class PptxMathCodec
{
    private const int MaxNodes = 2_048;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace Xml = XNamespace.Xml;

    internal static OpenXmlElement Build(PresentationTextRun source)
    {
        Validate(source);
        var paragraph = new XElement(
            M + "oMathPara",
            new XAttribute(XNamespace.Xmlns + "m", M),
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XElement(M + "oMathParaPr", new XElement(M + "jc", new XAttribute(M + "val", "left"))),
            new XElement(M + "oMath", BuildSequence(source.Formula.Expression, source)));
        var wrapper = new XElement(
            A14 + "m",
            new XAttribute(XNamespace.Xmlns + "a14", A14),
            paragraph);
        // The SDK models a14:m as a leaf even though Office stores an OMML
        // subtree inside it. Construct the known QName as an unknown composite
        // so the foreign OMML children survive serialization; a package reload
        // resolves the outer QName back to D14.TextMath without flattening XML.
        return Unknown(wrapper);
    }

    internal static bool IsMathElement(OpenXmlElement source) =>
        source.LocalName == "m" && source.NamespaceUri == A14.NamespaceName;

    internal static bool TryRead(OpenXmlElement source, out PresentationTextRun run)
    {
        if (!IsMathElement(source))
        {
            run = new PresentationTextRun();
            return false;
        }
        try
        {
            return TryRead(XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace), out run);
        }
        catch (System.Xml.XmlException)
        {
            run = new PresentationTextRun();
            return false;
        }
    }

    internal static bool IsCanonical(XElement source) => TryRead(source, out _);

    private static bool TryRead(XElement root, out PresentationTextRun run)
    {
        run = new PresentationTextRun();
        try
        {
            if (root.Name != A14 + "m") return false;
            var paragraph = root.Elements().SingleOrDefault();
            if (paragraph?.Name != M + "oMathPara") return false;
            var children = paragraph.Elements().ToArray();
            if (children.Length != 2 || children[0].Name != M + "oMathParaPr" || children[1].Name != M + "oMath") return false;
            if (!CanonicalParagraphProperties(children[0])) return false;

            var style = new StyleAccumulator();
            var budget = 0;
            var expression = ReadSequence(children[1].Elements(), style, ref budget);
            if (expression.Nodes.Count == 0) return false;
            run.Formula = new PresentationMathFormula
            {
                Expression = expression,
                PlainText = PpjLatexCompiler.PlainText(expression),
            };
            style.Apply(run);
            Validate(run);
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or FormatException or OverflowException or CodecException or System.Xml.XmlException)
        {
            run = new PresentationTextRun();
            return false;
        }
    }

    internal static bool SemanticallyEqual(OpenXmlElement source, PresentationTextRun requested)
    {
        if (!TryRead(source, out var current)) return false;
        var normalized = requested.Clone();
        if (normalized.Formula is not null) normalized.Formula.SourceLatex = string.Empty;
        return current.Equals(normalized);
    }

    internal static void Scrub(OpenXmlElement source)
    {
        if (!TryRead(source, out _)) return;
        foreach (var element in source.Descendants())
        {
            if (element is OM.Text text) text.Text = string.Empty;
            if (element.NamespaceUri != A.NamespaceName || element.LocalName != "rPr") continue;
            element.RemoveAttribute("sz", string.Empty);
            foreach (var fill in element.ChildElements.Where(child => child.LocalName == "solidFill").ToArray()) fill.Remove();
        }
    }

    internal static void Validate(PresentationTextRun source)
    {
        if (source.ContentCase != PresentationTextRun.ContentOneofCase.Formula || source.Formula?.Expression is null)
            throw new CodecException("invalid_presentation_formula", "Presentation formula inline requires a finite expression AST.");
        if (string.IsNullOrEmpty(source.Formula.PlainText))
            throw new CodecException("invalid_presentation_formula", "Presentation formula plain text must be non-empty.");
        if (source.HasBold || source.HasItalic || source.HasFontFamily || source.HasFontFamilyEastAsia ||
            source.HasFontKerningPoints || source.HasFontBaselinePercent || source.HasFontSpacingPoints || source.HasFontCaps ||
            source.HasUnderline || source.HasStrike || source.HasLanguage || source.HighlightCase != PresentationTextRun.HighlightOneofCase.None ||
            source.GradientFill is not null || source.Shadow is not null || source.HyperlinkCase != PresentationTextRun.HyperlinkOneofCase.None)
            throw new CodecException("invalid_presentation_formula_style", "Presentation formula runs support only direct size and solid color.");
        if (source.HasColorRgb && source.HasColorScheme)
            throw new CodecException("invalid_presentation_formula_style", "Presentation formula cannot specify both RGB and theme colors.");
        if (source.HasColorOpacityThousandthPercent && !source.HasColorRgb && !source.HasColorScheme)
            throw new CodecException("invalid_presentation_formula_style", "Presentation formula color opacity requires a direct color.");
        if (source.HasColorOpacityThousandthPercent && source.ColorOpacityThousandthPercent > 100_000)
            throw new CodecException("invalid_presentation_formula_style", "Presentation formula color opacity must be at most 100000 thousandths of a percent.");
        if (source.HasFontSizePoints && (!(source.FontSizePoints > 0) || source.FontSizePoints > 768 || !double.IsFinite(source.FontSizePoints)))
            throw new CodecException("invalid_presentation_formula_style", "Presentation formula size must be finite and between 0 and 768 points.");
        if (source.HasColorRgb) _ = PptxColor.Normalize(source.ColorRgb);
        if (source.HasColorScheme) _ = PptxColor.NormalizeScheme(source.ColorScheme);

        var budget = 0;
        ValidateSequence(source.Formula.Expression, ref budget);
        if (!PpjLatexCompiler.PlainText(source.Formula.Expression).Equals(source.Formula.PlainText, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_formula", "Presentation formula plain text does not match its expression AST.");
    }

    private static IEnumerable<XElement> BuildSequence(PresentationMathSequence sequence, PresentationTextRun style) =>
        sequence.Nodes.Select(node => BuildNode(node, style));

    private static XElement BuildNode(PresentationMathNode node, PresentationTextRun style) => node.KindCase switch
    {
        PresentationMathNode.KindOneofCase.Text => MathRun(node.Text, style),
        PresentationMathNode.KindOneofCase.Fraction => new XElement(
            M + "f",
            new XElement(M + "fPr", new XElement(M + "type", new XAttribute(M + "val", "bar"))),
            new XElement(M + "num", BuildSequence(node.Fraction.Numerator, style)),
            new XElement(M + "den", BuildSequence(node.Fraction.Denominator, style))),
        PresentationMathNode.KindOneofCase.Radical => new XElement(
            M + "rad",
            new XElement(M + "radPr", new XElement(M + "degHide", new XAttribute(M + "val", "1"))),
            new XElement(M + "deg"),
            new XElement(M + "e", BuildSequence(node.Radical.Radicand, style))),
        PresentationMathNode.KindOneofCase.Script => BuildScript(node.Script, style),
        _ => throw new CodecException("invalid_presentation_formula", "Presentation formula contains an unknown AST node."),
    };

    private static XElement BuildScript(PresentationMathScript script, PresentationTextRun style)
    {
        var hasSub = script.Subscript is not null;
        var hasSup = script.Superscript is not null;
        var kind = hasSub && hasSup ? "sSubSup" : hasSub ? "sSub" : "sSup";
        return new XElement(
            M + kind,
            new XElement(M + kind + "Pr", new XElement(M + "ctrlPr")),
            new XElement(M + "e", BuildSequence(script.Base, style)),
            hasSub ? new XElement(M + "sub", BuildSequence(script.Subscript!, style)) : null,
            hasSup ? new XElement(M + "sup", BuildSequence(script.Superscript!, style)) : null);
    }

    private static XElement MathRun(string text, PresentationTextRun style)
    {
        var properties = new XElement(A + "rPr", new XAttribute("lang", "en-US"));
        if (style.HasFontSizePoints)
            properties.Add(new XAttribute("sz", checked((int)Math.Round(style.FontSizePoints * 100)).ToString(CultureInfo.InvariantCulture)));
        if (style.HasColorRgb || style.HasColorScheme)
        {
            var color = style.HasColorRgb
                ? new XElement(A + "srgbClr", new XAttribute("val", PptxColor.Normalize(style.ColorRgb)))
                : new XElement(A + "schemeClr", new XAttribute("val", PptxColor.NormalizeScheme(style.ColorScheme)));
            if (style.HasColorOpacityThousandthPercent)
                color.Add(new XElement(A + "alpha", new XAttribute("val", style.ColorOpacityThousandthPercent.ToString(CultureInfo.InvariantCulture))));
            properties.Add(new XElement(A + "solidFill", color));
        }
        return new XElement(
            M + "r",
            properties,
            new XElement(M + "t", text.Any(char.IsWhiteSpace) ? new XAttribute(Xml + "space", "preserve") : null, text));
    }

    private static OpenXmlElement Unknown(XElement source)
    {
        if (source.Name == M + "t")
        {
            var text = new OM.Text(source.Value);
            if (source.Attribute(Xml + "space")?.Value == "preserve") text.Space = SpaceProcessingModeValues.Preserve;
            return text;
        }
        var prefix = source.GetPrefixOfNamespace(source.Name.Namespace) ?? string.Empty;
        var output = new OpenXmlUnknownElement(prefix, source.Name.LocalName, source.Name.NamespaceName);
        foreach (var attribute in source.Attributes())
        {
            if (attribute.IsNamespaceDeclaration)
            {
                output.AddNamespaceDeclaration(attribute.Name.LocalName == "xmlns" ? string.Empty : attribute.Name.LocalName, attribute.Value);
                continue;
            }
            var attributePrefix = source.GetPrefixOfNamespace(attribute.Name.Namespace) ?? string.Empty;
            output.SetAttribute(new OpenXmlAttribute(attributePrefix, attribute.Name.LocalName, attribute.Name.NamespaceName, attribute.Value));
        }
        foreach (var child in source.Elements()) output.Append(Unknown(child));
        return output;
    }

    private static PresentationMathSequence ReadSequence(IEnumerable<XElement> elements, StyleAccumulator style, ref int budget)
    {
        var sequence = new PresentationMathSequence();
        foreach (var element in elements)
        {
            budget++;
            if (budget > MaxNodes) throw new CodecException("presentation_formula_budget_exceeded", $"Presentation formula exceeds {MaxNodes} nodes.");
            sequence.Nodes.Add(ReadNode(element, style, ref budget));
        }
        return sequence;
    }

    private static PresentationMathNode ReadNode(XElement element, StyleAccumulator style, ref int budget)
    {
        if (element.Name == M + "r") return new PresentationMathNode { Text = ReadMathRun(element, style) };
        if (element.Name == M + "f")
        {
            var children = element.Elements().ToArray();
            if (children.Length != 3 || children[0].Name != M + "fPr" || !CanonicalFractionProperties(children[0]) ||
                children[1].Name != M + "num" || children[2].Name != M + "den") throw new FormatException();
            return new PresentationMathNode
            {
                Fraction = new PresentationMathFraction
                {
                    Numerator = ReadSequence(children[1].Elements(), style, ref budget),
                    Denominator = ReadSequence(children[2].Elements(), style, ref budget),
                },
            };
        }
        if (element.Name == M + "rad")
        {
            var children = element.Elements().ToArray();
            if (children.Length != 3 || children[0].Name != M + "radPr" || !CanonicalRadicalProperties(children[0]) ||
                children[1].Name != M + "deg" || children[1].HasElements || children[2].Name != M + "e") throw new FormatException();
            return new PresentationMathNode
            {
                Radical = new PresentationMathRadical { Radicand = ReadSequence(children[2].Elements(), style, ref budget) },
            };
        }
        if (element.Name.LocalName is "sSub" or "sSup" or "sSubSup" && element.Name.Namespace == M)
            return new PresentationMathNode { Script = ReadScript(element, style, ref budget) };
        throw new FormatException();
    }

    private static PresentationMathScript ReadScript(XElement element, StyleAccumulator style, ref int budget)
    {
        var kind = element.Name.LocalName;
        var children = element.Elements().ToArray();
        var expected = kind == "sSubSup" ? 4 : 3;
        if (children.Length != expected || children[0].Name != M + kind + "Pr" || !CanonicalScriptProperties(children[0]) || children[1].Name != M + "e")
            throw new FormatException();
        var output = new PresentationMathScript { Base = ReadSequence(children[1].Elements(), style, ref budget) };
        var cursor = 2;
        if (kind is "sSub" or "sSubSup")
        {
            if (children[cursor].Name != M + "sub") throw new FormatException();
            output.Subscript = ReadSequence(children[cursor++].Elements(), style, ref budget);
        }
        if (kind is "sSup" or "sSubSup")
        {
            if (children[cursor].Name != M + "sup") throw new FormatException();
            output.Superscript = ReadSequence(children[cursor].Elements(), style, ref budget);
        }
        return output;
    }

    private static string ReadMathRun(XElement element, StyleAccumulator style)
    {
        var children = element.Elements().ToArray();
        if (children.Length != 2 || children[0].Name != A + "rPr" || children[1].Name != M + "t") throw new FormatException();
        style.Read(children[0]);
        return children[1].Value;
    }

    private static bool CanonicalParagraphProperties(XElement element)
    {
        var children = element.Elements().ToArray();
        return children.Length == 1 && children[0].Name == M + "jc" && children[0].Attributes().Count() == 1 &&
               children[0].Attribute(M + "val")?.Value == "left";
    }

    private static bool CanonicalFractionProperties(XElement element)
    {
        var children = element.Elements().ToArray();
        return children.Length == 1 && children[0].Name == M + "type" && children[0].Attributes().Count() == 1 &&
               children[0].Attribute(M + "val")?.Value == "bar";
    }

    private static bool CanonicalRadicalProperties(XElement element)
    {
        var children = element.Elements().ToArray();
        return children.Length == 1 && children[0].Name == M + "degHide" && children[0].Attributes().Count() == 1 &&
               children[0].Attribute(M + "val")?.Value == "1";
    }

    private static bool CanonicalScriptProperties(XElement element)
    {
        var children = element.Elements().ToArray();
        return children.Length == 1 && children[0].Name == M + "ctrlPr" && !children[0].HasElements && !children[0].HasAttributes;
    }

    private static void ValidateSequence(PresentationMathSequence sequence, ref int budget)
    {
        if (sequence is null || sequence.Nodes.Count == 0)
            throw new CodecException("invalid_presentation_formula", "Presentation formula sequences must be non-empty.");
        foreach (var node in sequence.Nodes)
        {
            budget++;
            if (budget > MaxNodes) throw new CodecException("presentation_formula_budget_exceeded", $"Presentation formula exceeds {MaxNodes} nodes.");
            switch (node.KindCase)
            {
                case PresentationMathNode.KindOneofCase.Text:
                    if (string.IsNullOrEmpty(node.Text) || node.Text.Any(char.IsControl))
                        throw new CodecException("invalid_presentation_formula", "Presentation formula text nodes must be non-empty printable text.");
                    break;
                case PresentationMathNode.KindOneofCase.Fraction:
                    ValidateSequence(node.Fraction.Numerator, ref budget);
                    ValidateSequence(node.Fraction.Denominator, ref budget);
                    break;
                case PresentationMathNode.KindOneofCase.Radical:
                    ValidateSequence(node.Radical.Radicand, ref budget);
                    break;
                case PresentationMathNode.KindOneofCase.Script:
                    ValidateSequence(node.Script.Base, ref budget);
                    if (node.Script.Subscript is null && node.Script.Superscript is null)
                        throw new CodecException("invalid_presentation_formula", "Presentation formula script requires a subscript or superscript.");
                    if (node.Script.Subscript is not null) ValidateSequence(node.Script.Subscript, ref budget);
                    if (node.Script.Superscript is not null) ValidateSequence(node.Script.Superscript, ref budget);
                    break;
                default:
                    throw new CodecException("invalid_presentation_formula", "Presentation formula contains an unknown AST node.");
            }
        }
    }

    private sealed class StyleAccumulator
    {
        private bool initialized;
        private double? size;
        private string? rgb;
        private string? scheme;
        private uint? opacity;

        internal void Read(XElement properties)
        {
            if (properties.Attributes().Any(attribute => attribute.Name.LocalName is not ("lang" or "sz")) ||
                properties.Attribute("lang")?.Value != "en-US") throw new FormatException();
            double? currentSize = null;
            if (properties.Attribute("sz") is { } sizeAttribute)
                currentSize = int.Parse(sizeAttribute.Value, CultureInfo.InvariantCulture) / 100d;
            var fills = properties.Elements().ToArray();
            string? currentRgb = null;
            string? currentScheme = null;
            uint? currentOpacity = null;
            if (fills.Length > 1) throw new FormatException();
            if (fills.Length == 1)
            {
                if (fills[0].Name != A + "solidFill") throw new FormatException();
                var colors = fills[0].Elements().ToArray();
                if (colors.Length != 1) throw new FormatException();
                var color = colors[0];
                var value = color.Attribute("val")?.Value;
                if (value is null) throw new FormatException();
                if (color.Name == A + "srgbClr") currentRgb = PptxColor.Normalize(value);
                else if (color.Name == A + "schemeClr") currentScheme = PptxColor.NormalizeScheme(value);
                else throw new FormatException();
                var transforms = color.Elements().ToArray();
                if (transforms.Length > 1) throw new FormatException();
                if (transforms.Length == 1)
                {
                    if (transforms[0].Name != A + "alpha") throw new FormatException();
                    currentOpacity = uint.Parse(transforms[0].Attribute("val")?.Value ?? string.Empty, CultureInfo.InvariantCulture);
                    if (currentOpacity > 100_000) throw new FormatException();
                }
            }
            if (!initialized)
            {
                initialized = true;
                size = currentSize;
                rgb = currentRgb;
                scheme = currentScheme;
                opacity = currentOpacity;
            }
            else if (size != currentSize || rgb != currentRgb || scheme != currentScheme || opacity != currentOpacity)
            {
                throw new FormatException();
            }
        }

        internal void Apply(PresentationTextRun run)
        {
            if (size is { } fontSize) run.FontSizePoints = fontSize;
            if (rgb is not null) run.ColorRgb = rgb;
            if (scheme is not null) run.ColorScheme = scheme;
            if (opacity is { } colorOpacity) run.ColorOpacityThousandthPercent = colorOpacity;
        }
    }
}
