import { resolveColorToken } from "../shared/colors.mjs";
import { sampleChartTrendline } from "../shared/chart-trendlines.mjs";

function attrEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function svgNumber(value) {
  const rounded = Math.round(value * 1_000) / 1_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

export function presentationChartLineSvgAttributes(line) {
  if (!line) return "";
  const dash = {
    dot: "1 3",
    dash: "6 4",
    longDash: "10 4",
    dashDot: "6 3 1 3",
    longDashDot: "10 4 1 4",
    longDashDotDot: "10 3 1 3 1 3",
    systemDash: "4 3",
    systemDot: "1 2",
    systemDashDot: "4 2 1 2",
    systemDashDotDot: "4 2 1 2 1 2",
  }[line.style];
  return ` stroke="${attrEscape(resolveColorToken(line.fill, line.fill || "#0f172a"))}" stroke-width="${svgNumber(line.width)}"${dash ? ` stroke-dasharray="${dash}"` : ""}`;
}

function mapPoint(point, plot, max, domain, { horizontal, centered }) {
  const axisDomain = centered ? { start: domain.start - 0.5, end: domain.end + 0.5 } : domain;
  const domainWidth = axisDomain.end - axisDomain.start;
  if (!(domainWidth > 0) || !Number.isFinite(point.x) || !Number.isFinite(point.y) || point.y < 0 || point.y > max) return undefined;
  const categoryRatio = (point.x - axisDomain.start) / domainWidth;
  const valueRatio = point.y / max;
  if (horizontal) {
    return {
      x: plot.left + valueRatio * plot.width,
      y: plot.top + categoryRatio * plot.height,
    };
  }
  return {
    x: plot.left + categoryRatio * plot.width,
    y: plot.top + plot.height - valueRatio * plot.height,
  };
}

export function samplePresentationChartTrendline(values, trendline, options = {}) {
  return sampleChartTrendline(values, trendline, options);
}

function polylineSvg(points, trendline, attributes) {
  if (points.length < 2) return "";
  const encoded = points.map((point) => `${svgNumber(point.x)},${svgNumber(point.y)}`).join(" ");
  return `<polyline data-trendline-type="${trendline.type}" points="${encoded}" fill="none"${attributes}/>`;
}

export function presentationChartTrendlinesSvg(series, plot, max, categoryCount, { horizontal = false, centered = horizontal } = {}) {
  if (categoryCount < 2 || !(max > 0)) return "";
  return (series.trendlines || []).map((trendline) => {
    const line = trendline.line || { fill: series.color || "#475569", width: 1.5, style: "dash" };
    const attributes = presentationChartLineSvgAttributes(line);
    return sampleChartTrendline(series.values || [], trendline, { categoryCount })
      .flatMap(({ domain, points }) => {
        const segments = [];
        let current = [];
        for (const point of points) {
          const mapped = mapPoint(point, plot, max, domain, { horizontal, centered });
          if (mapped) current.push(mapped);
          else if (current.length) {
            if (current.length > 1) segments.push(current);
            current = [];
          }
        }
        if (current.length > 1) segments.push(current);
        return segments;
      })
      .map((points) => polylineSvg(points, trendline, attributes))
      .join("");
  }).join("");
}
