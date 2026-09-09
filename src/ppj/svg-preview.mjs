import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { compilePpjWorkspace, loadPpjWorkspace } from "./workspace.mjs";

const require = createRequire(import.meta.url);
const PREVIEW_CAPABILITIES = require("./svg-preview-capabilities.json");
export const SVG_PREVIEW_SUPPORTED_TYPES = new Set(PREVIEW_CAPABILITIES.supported.filter((value) => !value.includes(":")));
const SVG_PREVIEW_SUPPORTED_CHARTS = new Set(PREVIEW_CAPABILITIES.supported.filter((value) => value.startsWith("chart:")).map((value) => value.slice(6)));

const esc = (value) => String(value ?? "").replace(/[&<>"']/gu, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c]);
const num = (value, fallback = 0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
const frame = (e) => e?.frame || { x: 0, y: 0, width: 0, height: 0 };
const color = (value, fallback = "#D9E2F3") => typeof value === "string" && value ? value : fallback;
const paintColor = (paint, fallback = "#D9E2F3") => typeof paint === "string" ? paint : color(paint?.color?.hex || paint?.color?.value || paint?.hex, fallback);

function textValue(text) {
  if (typeof text === "string") return text;
  if (!text || typeof text !== "object") return "";
  if (typeof text.value === "string") return text.value;
  return (text.paragraphs || []).flatMap((p) => (p.runs || []).map((r) => r.text || "")).join("\n");
}

function renderElement(e, assets, diagnostics) {
  const f = frame(e), id = esc(e.id || "element");
  const common = ` data-officekit-id="${id}"`;
  if (e.type === "group") {
    const children = e.elements || e.children || [];
    return `<g${common}>${children.map((child) => renderElement(child, assets, diagnostics)).join("")}</g>`;
  }
  if (e.type === "placeholder") {
    diagnostics.push({ id: e.id, status: "partial", reason: "placeholder-preview" });
    return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3" stroke-dasharray="5 4"/><text x="${f.x + 8}" y="${f.y + 20}" font-size="13" fill="#667085">${esc(textValue(e.text) || e.name || "placeholder")}</text></g>`;
  }
  if (e.type === "component") {
    const items = e.repeat?.items || [e.arguments || {}], gap = num(e.repeat?.layout?.gap, 6), rh = Math.max(18, (f.height - gap * (items.length - 1)) / Math.max(1, items.length));
    diagnostics.push({ id: e.id, status: "partial", reason: "component-expanded-preview" });
    return `<g${common}>${items.map((item, i) => `<text x="${f.x}" y="${f.y + i * (rh + gap) + 16}" font-size="13" fill="#172033">${esc(item.arguments?.label || item.label || "")} ${esc(item.arguments?.value || item.value || "")}</text>`).join("")}</g>`;
  }
  if (e.type === "text" || e.text) {
    const lines = textValue(e.text).split("\n");
    return `<g${common}><text x="${f.x}" y="${f.y + 18}" font-family="Arial, sans-serif" font-size="${num(e.textStyle?.fontSize, 18)}" fill="${color(e.textStyle?.color, "#172033")}">${lines.map((line, i) => `<tspan x="${f.x}" dy="${i ? 22 : 0}">${esc(line)}</tspan>`).join("")}</text></g>`;
  }
  if (e.type === "image" || (e.type === "opaque" && e.asset)) {
    const asset = assets.get(e.asset); if (!asset) { diagnostics.push({ id: e.id, status: "unavailable", reason: "asset-missing" }); return `<rect${common} x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F3F4F6" stroke="#B8C0CC"/><text x="${f.x + 6}" y="${f.y + 20}"${common} fill="#667085">image unavailable</text>`; }
    return `<image${common} href="${esc(asset.href)}" x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" preserveAspectRatio="xMidYMid meet" opacity="${num(e.opacity, 1)}"/>`;
  }
  if (e.type === "connector") return `<line${common} x1="${f.x}" y1="${f.y + f.height / 2}" x2="${f.x + f.width}" y2="${f.y + f.height / 2}" stroke="${color(e.stroke?.color, "#667085")}" stroke-width="${num(e.stroke?.width, 2)}"/>`;
  if (e.type === "table") {
    const rows = e.rows || [], cols = Math.max(1, (e.columns || []).length || rows[0]?.cells?.length || 1), cw = f.width / cols, rh = f.height / Math.max(1, rows.length);
    return `<g${common}>${rows.map((r, ri) => (r.cells || []).map((c, ci) => `<rect x="${f.x + ci * cw}" y="${f.y + ri * rh}" width="${cw}" height="${rh}" fill="${ri ? "#FFFFFF" : "#E8EEF7"}" stroke="#B8C0CC"/><text x="${f.x + ci * cw + 5}" y="${f.y + ri * rh + 18}" font-size="12">${esc(textValue(c.text ?? c.value))}</text>`).join("")).join("")}</g>`;
  }
  if (e.type === "chart") {
    const series = e.data?.series || [], categories = e.data?.categories || [], values = series.flatMap((s) => (s.values || []).filter(Number.isFinite));
    const max = Math.max(1, ...values), base = f.y + f.height - 28, plotH = f.height - 48, step = f.width / Math.max(1, categories.length);
    const labels = categories.map((category, i) => `<text x="${f.x + i * step + step / 2}" y="${base + 16}" text-anchor="middle" font-size="10" fill="#667085">${esc(category)}</text>`).join("");
    if (e.chartType === "heatmap") {
      const rows = Math.max(1, series.length), cols = Math.max(1, categories.length), cw = f.width / cols, rh = Math.max(1, (f.height - 28) / rows);
      const all = values.length ? values : [0], lo = Math.min(...all), hi = Math.max(...all), cells = series.flatMap((s, ri) => (s.values || []).map((v, ci) => { const ratio = hi === lo ? .5 : (Number(v) - lo) / (hi - lo); const shade = Math.round(245 - Math.max(0, Math.min(1, ratio)) * 150); return Number.isFinite(v) ? `<rect x="${f.x + ci * cw}" y="${f.y + ri * rh}" width="${cw}" height="${rh}" fill="rgb(${shade},${Math.round(shade + 5)},${Math.round(shade + 10)})" stroke="#FFFFFF"/>` : `<rect x="${f.x + ci * cw}" y="${f.y + ri * rh}" width="${cw}" height="${rh}" fill="#F8FAFC" stroke="#FFFFFF" stroke-dasharray="2 2"/>`; })).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-heatmap-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${cells}${categories.map((c, i) => `<text x="${f.x + i * cw + cw / 2}" y="${f.y + f.height - 6}" text-anchor="middle" font-size="9" fill="#667085">${esc(c)}</text>`).join("")}</g>`;
    }
    if (e.chartType === "treemap") {
      const vals = (series[0]?.values || []).map(Number), total = vals.reduce((a, v) => a + (Number.isFinite(v) && v > 0 ? v : 0), 0) || 1; let x = f.x;
      const tiles = vals.map((v, i) => { const w = f.width * (Math.max(0, v) / total); const tile = `<rect x="${x}" y="${f.y}" width="${Math.max(0, w - 1)}" height="${f.height}" fill="hsl(${(i * 47) % 360} 45% 72%)" stroke="#FFFFFF"/><text x="${x + 5}" y="${f.y + 18}" font-size="11">${esc(categories[i] || String(i + 1))}</text>`; x += w; return tile; }).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-treemap-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${tiles}</g>`;
    }
    if (e.chartType === "sunburst") {
      const vals = (series[0]?.values || []).map(Number), total = vals.reduce((a, v) => a + (Number.isFinite(v) && v > 0 ? v : 0), 0) || 1, cx = f.x + f.width / 2, cy = f.y + f.height / 2, radius = Math.min(f.width, f.height) * .42; let start = -Math.PI / 2;
      const arcs = vals.map((v, i) => { const end = start + Math.max(0, v) / total * Math.PI * 2, large = end - start > Math.PI ? 1 : 0, x1 = cx + radius * Math.cos(start), y1 = cy + radius * Math.sin(start), x2 = cx + radius * Math.cos(end), y2 = cy + radius * Math.sin(end); const d = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z`; start = end; return `<path d="${d}" fill="hsl(${(i * 47) % 360} 45% 72%)" stroke="#FFFFFF"/>`; }).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-sunburst-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${arcs}</g>`;
    }
    if (e.chartType === "waterfall") {
      let running = 0; const vals = (series[0]?.values || []).map(Number), stepW = f.width / Math.max(1, vals.length), barsW = vals.map((v, i) => { const start = running; running += Number.isFinite(v) ? v : 0; const y = base - Math.max(start, running) / max * plotH, h = Math.abs(v || 0) / max * plotH; return Number.isFinite(v) ? `<rect x="${f.x + i * stepW + stepW * .15}" y="${y}" width="${stepW * .7}" height="${h}" fill="${v >= 0 ? "#6B8E9B" : "#C76B5C"}"/>` : ""; }).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-waterfall-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${barsW}${labels}</g>`;
    }
    if (e.chartType === "candlestick") {
      const vals = series[0]?.values || [], stepW = f.width / Math.max(1, vals.length), candles = vals.map((v, i) => { const o = Number(v?.open), h = Number(v?.high), l = Number(v?.low), c = Number(v?.close); if (![o, h, l, c].every(Number.isFinite)) return ""; const y = (n) => base - n / max * plotH; return `<line x1="${f.x + i * stepW + stepW / 2}" y1="${y(h)}" x2="${f.x + i * stepW + stepW / 2}" y2="${y(l)}" stroke="#344054"/><rect x="${f.x + i * stepW + stepW * .2}" y="${Math.min(y(o), y(c))}" width="${stepW * .6}" height="${Math.max(1, Math.abs(y(c) - y(o)))}" fill="${c >= o ? "#6B8E9B" : "#C76B5C"}"/>`; }).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-candlestick-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${candles}</g>`;
    }
    if (e.chartType === "sankey") {
      const nodes = e.data?.nodes || [], links = e.data?.links || [], col = Math.max(1, f.width / 3), nodeSvg = nodes.map((n, i) => `<rect x="${f.x + (n.x ?? (i % 2) * col)}" y="${f.y + (n.y ?? i * 24)}" width="${n.width ?? 56}" height="${n.height ?? 18}" fill="#B7C9D3"/><text x="${f.x + (n.x ?? (i % 2) * col) + 3}" y="${f.y + (n.y ?? i * 24) + 13}" font-size="10">${esc(n.label || n.name || i + 1)}</text>`).join("");
      diagnostics.push({ id: e.id, status: "partial", reason: "bounded-sankey-node-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${nodeSvg}<text x="${f.x + 8}" y="${f.y + f.height - 8}" font-size="10" fill="#667085">${links.length} flows</text></g>`;
    }
    if (e.chartType === "pictographic" || e.style?.symbol || series.some((s) => s.symbol)) {
      const vals = (series[0]?.values || []).map(Number), maxVal = Math.max(1, ...vals), unit = Math.max(1, Math.round(f.width / Math.max(1, vals.length * 8))), icons = vals.flatMap((v, i) => { const count = Math.min(32, Math.max(0, Math.round(v / maxVal * 12))); return Array.from({ length: count }, (_, j) => `<circle cx="${f.x + i * step + step * .2 + (j % 4) * unit}" cy="${f.y + f.height - 18 - Math.floor(j / 4) * unit}" r="${Math.max(3, unit * .28)}" fill="#6B8E9B"/>`); }).join("");
      diagnostics.push({ id: e.id, status: "supported", reason: "bounded-pictographic-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${icons}${labels}</g>`;
    }
    if (e.chartType === "streamgraph" || e.style?.stacking === "stream") {
      const bands = series.map((s, si) => { const points = (s.values || []).map((v, i) => `${f.x + i * step + step / 2},${base - num(v) / max * plotH - si * 8}`).join(" "); return `<polyline points="${points}" fill="none" stroke="hsl(${(si * 47) % 360} 45% 62%)" stroke-width="${Math.max(8, plotH / Math.max(2, series.length * 4))}" stroke-linecap="round"/>`; }).join("");
      diagnostics.push({ id: e.id, status: "partial", reason: "bounded-streamgraph-band-preview" });
      return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/>${bands}${labels}</g>`;
    }
    const bars = series.filter((s) => ["bar", "column"].includes(s.chartType)).flatMap((s) => (s.values || []).map((v, i) => `<rect x="${f.x + i * step + step * .18}" y="${base - num(v) / max * plotH}" width="${step * .64}" height="${num(v) / max * plotH}" fill="#98A2B3" opacity=".75"/>`)).join("");
    const lines = series.filter((s) => ["line", "area"].includes(s.chartType)).flatMap((s) => { const segments = []; let current = []; for (let i = 0; i < (s.values || []).length; i += 1) { const v = s.values[i]; if (Number.isFinite(v)) current.push(`${f.x + i * step + step / 2},${base - num(v) / max * plotH}`); else if (current.length) { segments.push(current); current = []; } } if (current.length) segments.push(current); return segments.filter((seg) => seg.length > 1).map((points) => `<polyline fill="none" stroke="#172033" stroke-width="${num(s.stroke?.width, 2.5)}" points="${points.join(" ")}"/>`); }).join("");
    const scatter = series.filter((s) => s.chartType === "scatter").flatMap((s) => (s.values || []).map((v, i) => { const y = typeof v === "object" ? v.y : v; const x = typeof v === "object" ? v.x : i; return Number.isFinite(y) ? `<circle cx="${f.x + (Number(x) / Math.max(1, categories.length - 1)) * f.width}" cy="${base - Number(y) / max * plotH}" r="3" fill="#0B5D5E"/>` : ""; })).join("");
    const pie = e.chartType === "pie" && values.length ? `<circle cx="${f.x + f.width / 2}" cy="${f.y + f.height / 2}" r="${Math.min(f.width, f.height) * .28}" fill="#D9E2F3" stroke="#667085"/><text x="${f.x + 8}" y="${f.y + 20}" font-size="13" fill="#475467">pie · ${esc(categories.join(" / "))}</text>` : "";
    const kind = pie ? "pie" : scatter ? "scatter" : "bounded-column-line";
    const status = SVG_PREVIEW_SUPPORTED_CHARTS.has(e.chartType) ? "supported" : "partial";
    diagnostics.push({ id: e.id, status, reason: SVG_PREVIEW_SUPPORTED_CHARTS.has(e.chartType) ? `${kind}-preview` : `chart-${e.chartType || "unknown"}-fallback` });
    return `<g${common}><rect x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F8FAFC" stroke="#98A2B3"/><line x1="${f.x + 12}" y1="${base}" x2="${f.x + f.width - 8}" y2="${base}" stroke="#667085"/>${bars}${lines}${scatter}${labels}${pie}</g>`;
  }
  if (e.type === "shape" || e.geometry) {
    const opacity = num(e.opacity ?? e.style?.opacity, 1);
    if (e.geometry?.customPaths || e.geometry?.path) diagnostics.push({ id: e.id, status: "partial", reason: "custom-geometry-bounded-rect-fallback" });
    return `<rect${common} x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" rx="${num(e.geometry?.radius, 0)}" fill="${paintColor(e.style?.fill, "#D9E2F3")}" fill-opacity="${opacity}" stroke="${paintColor(e.style?.stroke, "#667085")}" stroke-width="${num(e.style?.stroke?.width, 1)}"/>`;
  }
  if (e.type === "opaque" && e.previewAsset) {
    const asset = assets.get(e.previewAsset);
    if (asset?.href) { diagnostics.push({ id: e.id, status: "partial", reason: "opaque-source-preview" }); return `<image${common} href="${asset.href}" x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" preserveAspectRatio="xMidYMid meet"/>`; }
  }
  diagnostics.push({ id: e.id, status: "opaque", reason: `unsupported-${e.type || "unknown"}` });
  return `<rect${common} x="${f.x}" y="${f.y}" width="${f.width}" height="${f.height}" fill="#F9FAFB" stroke="#D0D5DD" stroke-dasharray="4 3"/>`;
}

export async function renderPpjToSvg(inputPath, { cwd = process.cwd(), outputDir } = {}) {
  const absolute = path.resolve(cwd, inputPath);
  const workspace = await loadPpjWorkspace(absolute, { cwd, retainRoot: true });
  const compiled = await compilePpjWorkspace(workspace, { includeNodeMap: false });
  const program = JSON.parse(Buffer.from(compiled.programJson).toString("utf8"));
  const canvas = program.design?.canvas || { width: 1280, height: 720 }, assets = new Map();
  for (const a of program.assets || []) {
    const assetPath = a.uri ? path.resolve(path.dirname(absolute), a.uri) : "";
    const bytes = assetPath ? await readFile(assetPath).catch(() => null) : null;
    assets.set(a.id, { href: bytes ? `data:${a.mimeType || "application/octet-stream"};base64,${Buffer.from(bytes).toString("base64")}` : "", mimeType: a.mimeType });
  }
  const diagnostics = [], pages = [];
  for (const page of program.pages || []) {
    const before = diagnostics.length;
    const body = (page.elements || []).map((e) => renderElement(e, assets, diagnostics)).join("\n");
    for (const element of page.elements || []) {
      const f = frame(element);
      if (f.x < 0 || f.y < 0 || f.x + f.width > canvas.width || f.y + f.height > canvas.height)
        diagnostics.push({ id: element.id, status: "partial", reason: "element-out-of-canvas" });
    }
    const pageFill = page.background?.fill?.color || page.background?.color || program.design?.theme?.background || "#FFFFFF";
    pages.push({ id: page.id, svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${canvas.width}" height="${canvas.height}" viewBox="0 0 ${canvas.width} ${canvas.height}"><rect width="100%" height="100%" fill="${esc(typeof pageFill === "string" ? pageFill : "#FFFFFF")}"/>${body}</svg>`, diagnostics: diagnostics.slice(before) });
  }
  const result = { renderer: "officekit-svg-preview", canvas, pages, diagnostics, status: diagnostics.length ? "partial" : "supported" };
  if (outputDir) {
    await mkdir(outputDir, { recursive: true });
    let sharp = null;
    try { const mod = await import("sharp"); sharp = mod.default || mod; } catch { diagnostics.push({ status: "unavailable", reason: "sharp-not-installed" }); }
    for (const page of pages) {
      await writeFile(path.join(outputDir, `${page.id}.svg`), page.svg);
      if (sharp) await sharp(Buffer.from(page.svg)).png().toFile(path.join(outputDir, `${page.id}.png`));
    }
    await writeFile(path.join(outputDir, "render.json"), `${JSON.stringify({ ...result, compile: { programSha256: compiled.programSha256, outputSha256: compiled.outputSha256 }, pages: pages.map(({ id, diagnostics: pageDiagnostics }) => ({ id, file: `${id}.svg`, png: `${id}.png`, diagnostics: pageDiagnostics })) }, null, 2)}\n`);
  }
  return result;
}
