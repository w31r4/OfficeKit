using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class OpenXmlChartLayoutCodec
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly string[] Children = ["layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h"];

    internal static void Validate(SpreadsheetChartLayoutArtifact? layout)
    {
        if (layout?.Manual is not { } m) return;
        if (m.HasTarget && m.Target is not ("inner" or "outer") ||
            new[] { m.HasXMode ? m.XMode : null, m.HasYMode ? m.YMode : null,
                m.HasWidthMode ? m.WidthMode : null, m.HasHeightMode ? m.HeightMode : null }
                .Any(mode => mode is not (null or "edge" or "factor")) ||
            new[] { m.X, m.Y, m.Width, m.Height }.Any(value => !double.IsFinite(value)))
            throw new CodecException("invalid_spreadsheet_chart", "Chart manual layout requires inner/outer target, edge/factor modes and finite numbers.");
    }

    internal static bool TryRead(XElement source, out SpreadsheetChartLayoutArtifact layout)
    {
        layout = new SpreadsheetChartLayoutArtifact();
        if (!Container(source) || source.Elements().Count() > 1) return false;
        if (source.Elements().SingleOrDefault() is not { } manual) return true;
        if (manual.Name != C + "manualLayout" || !Container(manual)) return false;
        var m = new SpreadsheetChartManualLayoutArtifact();
        var previous = -1;
        foreach (var child in manual.Elements())
        {
            var index = child.Name.Namespace == C ? Array.IndexOf(Children, child.Name.LocalName) : -1;
            if (index <= previous || child.HasElements || UnexpectedNodes(child) ||
                child.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "val")) return false;
            previous = index;
            var literal = (string?)child.Attribute("val");
            if (index < 5)
            {
                literal ??= index == 0 ? "outer" : "factor";
                if (index == 0 ? literal is not ("inner" or "outer") : literal is not ("edge" or "factor")) return false;
                switch (index)
                {
                    case 0: m.Target = literal; break;
                    case 1: m.XMode = literal; break;
                    case 2: m.YMode = literal; break;
                    case 3: m.WidthMode = literal; break;
                    case 4: m.HeightMode = literal; break;
                }
            }
            else
            {
                if (!double.TryParse(literal, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || !double.IsFinite(value)) return false;
                switch (index)
                {
                    case 5: m.X = value; break;
                    case 6: m.Y = value; break;
                    case 7: m.Width = value; break;
                    case 8: m.Height = value; break;
                }
            }
        }
        layout.Manual = m;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartLayoutArtifact? layout)
    {
        if (layout is null) return null;
        Validate(layout);
        var result = new XElement(C + "layout");
        if (layout.Manual is not { } m) return result;
        result.Add(new XElement(C + "manualLayout",
            Leaf("layoutTarget", m.HasTarget, m.Target),
            Leaf("xMode", m.HasXMode, m.XMode), Leaf("yMode", m.HasYMode, m.YMode),
            Leaf("wMode", m.HasWidthMode, m.WidthMode), Leaf("hMode", m.HasHeightMode, m.HeightMode),
            Leaf("x", m.HasX, m.X), Leaf("y", m.HasY, m.Y),
            Leaf("w", m.HasWidth, m.Width), Leaf("h", m.HasHeight, m.Height)));
        return result;
    }

    internal static SpreadsheetChartLayoutArtifact FromPpj(JsonElement source)
    {
        var layout = new SpreadsheetChartLayoutArtifact();
        if (!source.TryGetProperty("manual", out var manual)) return layout;
        var m = new SpreadsheetChartManualLayoutArtifact();
        if (manual.TryGetProperty("target", out var target)) m.Target = target.GetString()!;
        if (manual.TryGetProperty("xMode", out var xMode)) m.XMode = xMode.GetString()!;
        if (manual.TryGetProperty("yMode", out var yMode)) m.YMode = yMode.GetString()!;
        if (manual.TryGetProperty("widthMode", out var widthMode)) m.WidthMode = widthMode.GetString()!;
        if (manual.TryGetProperty("heightMode", out var heightMode)) m.HeightMode = heightMode.GetString()!;
        if (manual.TryGetProperty("x", out var x)) m.X = x.GetDouble();
        if (manual.TryGetProperty("y", out var y)) m.Y = y.GetDouble();
        if (manual.TryGetProperty("width", out var width)) m.Width = width.GetDouble();
        if (manual.TryGetProperty("height", out var height)) m.Height = height.GetDouble();
        layout.Manual = m;
        Validate(layout);
        return layout;
    }

    internal static JsonObject Project(SpreadsheetChartLayoutArtifact layout)
    {
        var result = new JsonObject();
        if (layout.Manual is not { } m) return result;
        var manual = new JsonObject();
        if (m.HasTarget) manual["target"] = m.Target;
        if (m.HasXMode) manual["xMode"] = m.XMode;
        if (m.HasYMode) manual["yMode"] = m.YMode;
        if (m.HasWidthMode) manual["widthMode"] = m.WidthMode;
        if (m.HasHeightMode) manual["heightMode"] = m.HeightMode;
        if (m.HasX) manual["x"] = m.X;
        if (m.HasY) manual["y"] = m.Y;
        if (m.HasWidth) manual["width"] = m.Width;
        if (m.HasHeight) manual["height"] = m.Height;
        result["manual"] = manual;
        return result;
    }

    private static XElement? Leaf(string name, bool present, object value) => present ? new XElement(C + name, new XAttribute("val", value)) : null;
    private static bool Container(XElement element) => !element.Attributes().Any(a => !a.IsNamespaceDeclaration) && !UnexpectedNodes(element);
    private static bool UnexpectedNodes(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
}
