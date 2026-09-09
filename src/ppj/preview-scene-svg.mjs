// Internal scene painter under construction. The production preview switches
// only after G-01's assessment/publication and end-to-end gates are complete.
// No PPJ interpretation, OOXML parsing, filesystem or raster backend here.
import { isFieldSet } from "@bufbuild/protobuf";
import { PresentationElementSchema, PresentationSlideSchema, PresentationTextBodySchema,
  PresentationTextParagraphSchema, PresentationTextRunSchema, PresentationTextStyleSchema,
  PresentationBackgroundSchema, PresentationTableRowSchema, PresentationTableCellSchema,
  PresentationTableCellFillSchema, PresentationTableCellBordersSchema,
  SpreadsheetChartLineStyleArtifactSchema, SpreadsheetColorSchema, SpreadsheetChartType,
  SpreadsheetChartSeriesArtifactSchema, SpreadsheetChartAxisArtifactSchema,
  SpreadsheetChartLineOptionsArtifactSchema, SpreadsheetChartMarkerArtifactSchema } from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { createPpjSceneView, scenePoints, sceneOpacity, sceneFontPoints } from "./preview-scene-view.mjs";
import { escapePreviewText as esc, previewDiagnostic, previewReliability, aggregatePreviewStatus } from "./preview-diagnostics.mjs";

const content = new Map(PresentationElementSchema.fields.filter(f => f.oneof?.localName === "content").map(f => [f.localName, f.message]));
const frameFields = ["leftEmu", "topEmu", "widthEmu", "heightEmu"];
const numeric = value => {
  if (typeof value === "bigint") {
    if (value > BigInt(Number.MAX_SAFE_INTEGER) || value < BigInt(Number.MIN_SAFE_INTEGER)) throw new RangeError("Unsafe path coordinate");
    value = Number(value);
  }
  if (typeof value !== "number" || !Number.isFinite(value)) throw new TypeError("Non-finite native geometry");
  return Object.is(value, -0) ? 0 : value;
};
const n = value => String(numeric(value));
const rgb = (value, fallback = "none") => /^[0-9a-f]{6}$/iu.test(value || "") ? `#${value}` : fallback;
const box = f => `x="${n(f.x)}" y="${n(f.y)}" width="${n(f.width)}" height="${n(f.height)}"`;

/** Literal native paths only. An unresolved command fails the whole path,
 * never drops one segment and rejoins unrelated endpoints. Coordinates are
 * mapped before painting so anisotropic viewport scaling does not scale pens. */
export function nativePathData(path, frame) {
  const width = numeric(path.width), height = numeric(path.height);
  if (width < 0 || height < 0) throw new RangeError("Negative path viewport");
  const sx = width === 0 ? 1 / 12700 : frame.width / width;
  const sy = height === 0 ? 1 / 12700 : frame.height / height;
  const point = value => {
    if (!value || value.xReference !== undefined || value.yReference !== undefined)
      throw new TypeError("Unresolved path point/reference");
    return `${n(frame.x + numeric(value.x) * sx)} ${n(frame.y + numeric(value.y) * sy)}`;
  };
  let started = false;
  return path.commands.map(({ command }) => {
    if (command.case === "moveTo") { started = true; return `M ${point(command.value)}`; }
    if (!started) throw new TypeError("Path requires moveTo");
    switch (command.case) {
      case "lineTo": return `L ${point(command.value)}`;
      case "cubicBezierTo": return `C ${point(command.value.control1)} ${point(command.value.control2)} ${point(command.value.end)}`;
      case "quadraticBezierTo": return `Q ${point(command.value.control)} ${point(command.value.end)}`;
      case "close": return "Z";
      default: throw new TypeError(`Unpainted native path command: ${command.case}`);
    }
  }).join(" ");
}

function frameTransform(f, transform) {
  if (!transform) return "";
  const cx = f.x + f.width / 2, cy = f.y + f.height / 2;
  return `translate(${n(cx)} ${n(cy)}) rotate(${n(transform.rotation ?? 0)}) scale(${transform.flipH ? -1 : 1} ${transform.flipV ? -1 : 1}) translate(${n(-cx)} ${n(-cy)})`;
}

/** Actual native-state-to-SVG drawing; deliberately not a public publication
 * receipt. No support promotion or G-11 rule retirement is implied. */
export function paintPpjSceneSvg(receipt) {
  const view = createPpjSceneView(receipt), diagnostics = [...view.diagnostics];
  const limit = (node, field, reason = "preview.scene.paint.unmapped", value, status = "partial") => {
    diagnostics.push(Object.freeze({ ...previewDiagnostic({
      pageId: node.pageId || undefined, id: node.pageId ? node.semanticId : undefined,
      path: node.path || "$", status, reason, value,
      action: "Review the native scene field; this internal painter is not full visual or edit-fidelity acceptance.",
    }), scenePath: `${node.scenePath}${field ? `.${field}` : ""}` }));
  };
  // Generated descriptors, not a second feature vocabulary, determine which
  // present fields remain unconsumed. Explicit false/zero retain presence.
  function unused(schema, value, handled, node, prefix = "") {
    const used = new Set(handled);
    for (const field of schema.fields) if (!used.has(field.localName) && isFieldSet(value, field))
      limit(node, `${prefix}${field.localName}`, "preview.scene.paint.unmapped", value[field.oneof?.localName || field.localName]);
  }
  function placeholder(node, label) {
    const f = node.frame;
    if (!f) return `<title>${esc(label)}</title>`;
    return `<rect ${box(f)} fill="#FFF7ED" stroke="#9A3412" stroke-dasharray="3 2"/><text x="${n(f.x + 2)}" y="${n(f.y + 12)}" font-size="10" fill="#9A3412">${esc(label)}</text>`;
  }
  function text(node, source = node.native, ownerField = "shape", fallback = {}) {
    const body = source.textBody, f = node.frame;
    if (!body && !source.text) return "";
    limit(node, `${ownerField}.textBody`, "preview.scene.paint.text-layout", "Font metrics, wrapping, AutoFit and inherited text state remain unverified.");
    // Rich paragraphs are authoritative. The compatibility string must never
    // duplicate or replace their run boundaries.
    const paragraphs = body?.paragraphs ?? [{ runs: [{ content: { case: "text", value: source.text } }] }];
    if (body) unused(PresentationTextBodySchema, body, ["paragraphs"], node, `${ownerField}.textBody.`);
    const properties = body?.bodyProperties;
    const inset = (key, selected, fallback) => properties?.[key]?.case === selected ? scenePoints(properties[key].value) : fallback;
    const left = f.x + inset("leftInset", "leftInsetEmu", 7.2), right = f.x + f.width - inset("rightInset", "rightInsetEmu", 7.2);
    let y = f.y + inset("topInset", "topInsetEmu", 3.6);
    return paragraphs.map((paragraph, pi) => {
      const defaults = paragraph.defaultRunStyle?.case === "defaultRunProperties" ? paragraph.defaultRunStyle.value : {};
      const prefix = `${ownerField}.textBody.paragraphs[${pi}].`;
      if (body) {
        unused(PresentationTextParagraphSchema, paragraph, ["runs", "alignment", "defaultRunProperties"], node, prefix);
        if (defaults.$typeName) unused(PresentationTextStyleSchema, defaults,
          ["fontSizePoints", "fontFamily", "bold", "italic", "colorRgb", "colorOpacityThousandthPercent"], node, `${prefix}defaultRunProperties.`);
      }
      const align = paragraph.alignment;
      if (align && !["left", "center", "right"].includes(align)) limit(node, `${prefix}alignment`, "preview.scene.paint.text-alignment", align);
      const x = align === "center" ? (left + right) / 2 : align === "right" ? right : left;
      const lines = [[]];
      for (const [ri, run] of paragraph.runs.entries()) {
        const rp = `${prefix}runs[${ri}].`;
        if (body) unused(PresentationTextRunSchema, run,
          ["text", "lineBreak", "fontSizePoints", "fontFamily", "bold", "italic", "colorRgb", "colorOpacityThousandthPercent"], node, rp);
        if (run.content.case === "lineBreak") { lines.push([]); continue; }
        if (run.content.case !== "text") { limit(node, `${rp}content`, "preview.scene.paint.text-content", run.content.case); continue; }
        const size = sceneFontPoints(run.fontSizePoints ?? defaults.fontSizePoints ?? fallback.fontSizePoints ?? 18);
        const defaultColor = defaults.color?.case === "colorRgb" ? defaults.color.value : fallback.color?.case === "colorRgb" ? fallback.color.value : undefined;
        const style = `font-family="${esc(run.fontFamily || defaults.fontFamily || fallback.fontFamily || "sans-serif")}" font-size="${n(size)}" font-weight="${(run.bold ?? defaults.bold ?? fallback.bold) ? "bold" : "normal"}" font-style="${(run.italic ?? defaults.italic ?? fallback.italic) ? "italic" : "normal"}" fill="${rgb(run.colorRgb || defaultColor, "#000000")}" fill-opacity="${n(sceneOpacity(run.colorOpacityThousandthPercent ?? defaults.colorOpacityThousandthPercent ?? fallback.colorOpacityThousandthPercent ?? 100000))}"`;
        // Preserve explicit newlines as line boundaries, not one line per run.
        run.content.value.split(/\r\n|\r|\n/u).forEach((segment, i) => {
          if (i) lines.push([]);
          lines.at(-1).push({ size, svg: `<tspan ${style}>${esc(segment)}</tspan>` });
        });
      }
      return lines.map(line => {
        const size = Math.max(sceneFontPoints(defaults.fontSizePoints ?? fallback.fontSizePoints ?? 18), ...line.map(run => run.size));
        y += size;
        const svg = `<text x="${n(x)}" y="${n(y)}" text-anchor="${align === "center" ? "middle" : align === "right" ? "end" : "start"}" xml:space="preserve">${line.map(run => run.svg).join("")}</text>`;
        y += size * .2;
        return svg;
      }).join("");
    }).join("");
  }
  function shape(node) {
    const s = node.native, f = node.frame;
    unused(content.get("shape"), s, [...frameFields, "geometry", "text", "textBody", "fillRgb", "lineRgb", "lineWidthEmu",
      "fillOpacityThousandthPercent", "lineOpacityThousandthPercent", "lineStyle", "transform", "customPaths"], node, "shape.");
    if (s.lineStyle && !["solid", "none"].includes(s.lineStyle)) limit(node, "shape.lineStyle", "preview.scene.paint.line-style", s.lineStyle);
    const paint = `fill="${rgb(s.fillRgb)}" fill-opacity="${n(sceneOpacity(s.fillOpacityThousandthPercent ?? 100000))}" stroke="${s.lineStyle === "none" ? "none" : rgb(s.lineRgb)}" stroke-opacity="${n(sceneOpacity(s.lineOpacityThousandthPercent ?? 100000))}" stroke-width="${n(scenePoints(s.lineWidthEmu))}"`;
    let geometry;
    if (s.customPaths.length) geometry = s.customPaths.map((path, i) => {
      const field = `shape.customPaths[${i}]`;
      try {
        if (![0, 1, 2].includes(path.fillMode)) throw new TypeError("Unknown path fill mode");
        const d = nativePathData(path, f);
        if (path.extrusionAllowed !== undefined) limit(node, `${field}.extrusionAllowed`, "preview.scene.paint.unmapped", path.extrusionAllowed);
        // Override the paint on a child, avoiding duplicate XML attributes.
        return `<g ${paint}><path data-officekit-path="${i}" d="${d}"${path.fillMode === 2 ? ' fill="none"' : ""}${path.stroke === false ? ' stroke="none"' : ""}/></g>`;
      } catch (error) {
        limit(node, field, "preview.scene.paint.path", error.message, "unavailable");
        return placeholder(node, "path unavailable");
      }
    }).join("");
    // The codec lowers its "textbox" marker to native rect geometry too.
    else if (["rect", "textbox", "flowChartProcess"].includes(s.geometry)) geometry = `<rect ${box(f)} ${paint}/>`;
    else if (s.geometry === "ellipse") geometry = `<ellipse cx="${n(f.x + f.width / 2)}" cy="${n(f.y + f.height / 2)}" rx="${n(f.width / 2)}" ry="${n(f.height / 2)}" ${paint}/>`;
    else if (["diamond", "flowChartDecision"].includes(s.geometry)) geometry = `<path d="M ${n(f.x + f.width / 2)} ${n(f.y)} L ${n(f.x + f.width)} ${n(f.y + f.height / 2)} L ${n(f.x + f.width / 2)} ${n(f.y + f.height)} L ${n(f.x)} ${n(f.y + f.height / 2)} Z" ${paint}/>`;
    else { limit(node, "shape.geometry", "preview.scene.paint.preset", s.geometry); geometry = placeholder(node, `geometry: ${s.geometry || "unresolved"}`); }
    return geometry + text(node);
  }
  function image(node) {
    const s = node.native, asset = view.asset(s.svgAssetId || s.assetId);
    unused(content.get("image"), s, [...frameFields, "assetId", "svgAssetId", "opacityThousandthPercent", "transform", "altText", "accessibilityTitle", "accessibilityDecorative"], node, "image.");
    if (!asset?.data?.byteLength) { limit(node, "image.assetId", "preview.scene.paint.asset", s.assetId, "unavailable"); return placeholder(node, "image unavailable"); }
    const href = `data:${asset.contentType};base64,${Buffer.from(asset.data).toString("base64")}`;
    return `<image ${box(node.frame)} href="${esc(href)}" preserveAspectRatio="none" opacity="${n(sceneOpacity(s.opacityThousandthPercent ?? 100000))}"><title>${esc(s.altText || s.accessibilityTitle || "")}</title></image>`;
  }
  function linePaint(node, field, color, width, opacity, dash, cap, join) {
    if (width < 0) throw new RangeError("Negative native line width");
    const patterns = { solid: [], dashed: [4, 3], dotted: [1, 3], "dash-dot": [4, 3, 1, 3], "dash-dot-dot": [8, 3, 1, 3, 1, 3] };
    if (!Object.hasOwn(patterns, dash) || cap && !["flat", "round", "square"].includes(cap) || join && !["miter", "round", "bevel"].includes(join))
      throw new TypeError("Unsupported native line paint token");
    if (patterns[dash].length) limit(node, `${field}.lineStyle`, "preview.scene.paint.dash-metrics", "Native dash kind retained; exact host dash lengths remain approximate.");
    return `stroke="${color}" stroke-width="${n(width)}" stroke-opacity="${n(opacity)}" stroke-linecap="${cap === "flat" || !cap ? "butt" : cap}" stroke-linejoin="${join || "miter"}"${patterns[dash].length ? ` stroke-dasharray="${patterns[dash].map(v => n(v * width)).join(" ")}"` : ""}`;
  }
  function connector(node) {
    const s = node.native, { start, end } = node.endpoints;
    unused(content.get("connector"), s, ["connectorType", "startXEmu", "startYEmu", "endXEmu", "endYEmu", "lineRgb", "lineWidthEmu", "lineStyle", "lineCap", "lineJoin",
      "lineOpacityThousandthPercent", "startArrow", "endArrow", "startArrowWidth", "startArrowLength", "endArrowWidth", "endArrowLength",
      "startTargetId", "endTargetId", "startConnectionSiteIndex", "endConnectionSiteIndex"], node, "connector.");
    // Consume the actual writer/candidate endpoints. Compiler anchor correctness
    // is a separate factual check; neither frame direction nor nearest objects
    // participate in this routing.
    const points = [start];
    if (s.connectorType === "elbow") {
      const mid = (start.x + end.x) / 2;
      points.push({ x: mid, y: start.y }, { x: mid, y: end.y });
      if (view.scene.origin === 2) limit(node, "connector.connectorType", "preview.scene.paint.imported-route", "Canonical midpoint elbow; imported preset rotation/adjustment provenance is not carried separately.");
    } else if (s.connectorType !== "straight") {
      limit(node, "connector.connectorType", "preview.scene.paint.connector-route", s.connectorType, "unavailable");
      return placeholder(node, `${s.connectorType}: route unavailable`);
    }
    points.push(end);
    const width = scenePoints(s.lineWidthEmu), opacity = sceneOpacity(s.lineOpacityThousandthPercent ?? 100000);
    const color = s.lineStyle === "none" ? "none" : rgb(s.lineRgb);
    const paint = linePaint(node, "connector", color, width, opacity, s.lineStyle === "none" ? "solid" : s.lineStyle || "solid", s.lineCap, s.lineJoin);
    const route = points.map((p, i) => `${i ? "L" : "M"} ${n(p.x)} ${n(p.y)}`).join(" ");
    const endpoint = (name, p, neighbours) => {
      const kind = s[`${name}Arrow`];
      if (!kind || kind === "none" || color === "none" || width === 0) return "";
      const previous = neighbours.find(q => q.x !== p.x || q.y !== p.y);
      if (!previous) { limit(node, `connector.${name}Arrow`, "preview.scene.paint.arrow-direction", "Coincident endpoints have no direction", "unavailable"); return ""; }
      const size = token => { if (!token) return 3; if (!Object.hasOwn({ sm: 2, med: 3, lg: 5 }, token)) throw new TypeError("Unknown native arrow size"); return { sm: 2, med: 3, lg: 5 }[token]; };
      const length = width * size(s[`${name}ArrowLength`]), half = width * size(s[`${name}ArrowWidth`]) / 2;
      const angle = Math.atan2(p.y - previous.y, p.x - previous.x) * 180 / Math.PI;
      const tails = `L ${n(-length)} ${n(-half)}`;
      let head;
      if (kind === "triangle") head = `<path d="M 0 0 ${tails} L ${n(-length)} ${n(half)} Z"/>`;
      else if (kind === "stealth") head = `<path d="M 0 0 ${tails} L ${n(-length * .7)} 0 L ${n(-length)} ${n(half)} Z"/>`;
      else if (kind === "diamond") head = `<path d="M 0 0 L ${n(-length / 2)} ${n(-half)} L ${n(-length)} 0 L ${n(-length / 2)} ${n(half)} Z"/>`;
      else if (kind === "oval") head = `<ellipse cx="${n(-length / 2)}" cy="0" rx="${n(length / 2)}" ry="${n(half)}"/>`;
      else if (kind === "arrow") head = `<path d="M ${n(-length)} ${n(-half)} L 0 0 L ${n(-length)} ${n(half)}" fill="none" stroke="${color}" stroke-width="${n(width)}"/>`;
      else { limit(node, `connector.${name}Arrow`, "preview.scene.paint.arrow-kind", kind, "unavailable"); return ""; }
      limit(node, `connector.${name}Arrow`, "preview.scene.paint.arrow-metrics", "Direction, kind and relative size mapped; host contour metrics remain approximate.");
      return `<g data-officekit-arrow="${name}" data-officekit-arrow-kind="${esc(kind)}" transform="translate(${n(p.x)} ${n(p.y)}) rotate(${n(angle)})" fill="${color}" opacity="${n(opacity)}">${head}</g>`;
    };
    return `<g data-officekit-start-target="${esc(s.startTargetId)}" data-officekit-start-site="${s.startConnectionSiteIndex}" data-officekit-end-target="${esc(s.endTargetId)}" data-officekit-end-site="${s.endConnectionSiteIndex}"><path data-officekit-connector="${esc(s.connectorType)}" d="${route}" fill="none" ${paint}/>${endpoint("start", start, points.slice(1))}${endpoint("end", end, points.slice(0, -1).reverse())}</g>`;
  }
  function table(node) {
    const s = node.native, f = node.frame;
    unused(content.get("table"), s, [...frameFields, "columnWidthsEmu", "rows", "mergeRanges", "frameTransform", "firstRow", "defaultCellFillRgb", "noDefaultCellFill", "defaultTextStyle"], node, "table.");
    const widths = s.columnWidthsEmu.map(scenePoints), heights = s.rows.map(r => scenePoints(r.heightEmu));
    if (!widths.length || !heights.length || [...widths, ...heights].some(v => v <= 0) || s.rows.some(r => r.cells.length !== widths.length))
      throw new RangeError("Invalid native table grid");
    const offsets = values => values.reduce((all, value) => { all.push(all.at(-1) + value); return all; }, [0]);
    const xs = offsets(widths), ys = offsets(heights), sx = f.width / xs.at(-1), sy = f.height / ys.at(-1);
    const merged = new Map();
    for (const range of s.mergeRanges) {
      const { startRow: r0, endRow: r1, startColumn: c0, endColumn: c1 } = range;
      if (![r0, r1, c0, c1].every(Number.isSafeInteger) || r0 < 0 || c0 < 0 || r1 < r0 || c1 < c0 || r1 >= heights.length || c1 >= widths.length || r0 === r1 && c0 === c1)
        throw new RangeError("Invalid native table merge range");
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
        const key = `${r}:${c}`, cell = s.rows[r].cells[c], origin = r === r0 && c === c0;
        if (merged.has(key)) throw new RangeError("Overlapping native table merge ranges");
        if (!origin && (cell.text || cell.textBody?.paragraphs.some(p => p.runs.some(run => run.content.case === "text" ? run.content.value.length > 0 : ["field", "formula"].includes(run.content.case)))))
          throw new TypeError("Covered native table cell contains visible content");
        merged.set(key, { range, origin });
      }
    }
    const styleFields = ["fontSizePoints", "fontFamily", "bold", "italic", "colorRgb", "colorOpacityThousandthPercent"];
    if (s.defaultTextStyle) unused(PresentationTextStyleSchema, s.defaultTextStyle, styleFields, node, "table.defaultTextStyle.");
    return s.rows.map((row, r) => {
      unused(PresentationTableRowSchema, row, ["heightEmu", "cells"], node, `table.rows[${r}].`);
      return row.cells.map((cell, c) => {
        const prefix = `table.rows[${r}].cells[${c}]`, merge = merged.get(`${r}:${c}`);
        if (merge && !merge.origin) {
          if (cell.fill || cell.borders) limit(node, prefix, "preview.scene.paint.merged-cell-style", "Covered physical cell paint/borders require perimeter/style resolution.");
          return "";
        }
        const c1 = merge?.range.endColumn ?? c, r1 = merge?.range.endRow ?? r;
        const cellFrame = { x: f.x + xs[c] * sx, y: f.y + ys[r] * sy, width: (xs[c1 + 1] - xs[c]) * sx, height: (ys[r1 + 1] - ys[r]) * sy };
        unused(PresentationTableCellSchema, cell, ["text", "textBody", "fill", "borders", "textStyle"], node, `${prefix}.`);
        const header = s.firstRow === true && r === 0, authored = view.scene.origin === 1;
        let fill = "none", opacity = 1;
        if (cell.fill) {
          unused(PresentationTableCellFillSchema, cell.fill, ["noFill", "solidRgb", "opacityThousandthPercent"], node, `${prefix}.fill.`);
          if (cell.fill.kind.case === "solidRgb") { fill = rgb(cell.fill.kind.value); opacity = sceneOpacity(cell.fill.opacityThousandthPercent ?? 100000); }
          else if (cell.fill.kind.case !== "noFill") limit(node, `${prefix}.fill`, "preview.scene.paint.table-fill", cell.fill.kind.case);
        } else if (authored) {
          if (s.defaultCellFill.case !== "noDefaultCellFill") fill = s.defaultCellFill.case === "defaultCellFillRgb" ? rgb(s.defaultCellFill.value) : header ? "#EDEDED" : "#FFFFFF";
        } else limit(node, `${prefix}.fill`, "preview.scene.paint.table-inherited-fill", "No direct candidate fill; do not substitute authored defaults.");
        let borderSvg = "";
        if (cell.borders) {
          unused(PresentationTableCellBordersSchema, cell.borders, ["left", "right", "top", "bottom"], node, `${prefix}.borders.`);
          const { x, y, width: w, height: h } = cellFrame;
          const edges = { left: [x, y, x, y + h], right: [x + w, y, x + w, y + h], top: [x, y, x + w, y], bottom: [x, y + h, x + w, y + h] };
          for (const [side, [x1, y1, x2, y2]] of Object.entries(edges)) {
            const b = cell.borders[side]; if (!b) continue;
            const field = `${prefix}.borders.${side}`;
            unused(SpreadsheetChartLineStyleArtifactSchema, b, ["color", "widthPoints", "opacityThousandthPercent", "dashStyle", "cap", "join"], node, `${field}.`);
            if (b.color) unused(SpreadsheetColorSchema, b.color, ["rgb"], node, `${field}.color.`);
            if (b.color?.source.case !== "rgb" || b.widthPoints === undefined) { limit(node, field, "preview.scene.paint.table-border-inheritance", "Border color or width remains inherited."); continue; }
            const dash = ["solid", "solid", "dashed", "dotted", "dash-dot", "dash-dot-dot"][b.dashStyle];
            borderSvg += `<line data-officekit-cell-border="${side}" x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" ${linePaint(node, field, rgb(b.color.source.value), b.widthPoints, sceneOpacity(b.opacityThousandthPercent ?? 100000), dash, b.cap, b.join)}/>`;
          }
        }
        if (cell.textStyle) unused(PresentationTextStyleSchema, cell.textStyle, styleFields, node, `${prefix}.textStyle.`);
        // Authored rich bodies already contain compiler-resolved text state.
        // Candidate uniform textStyle is direct imported state, not a new PPJ
        // model. Legacy authored defaults match PptxTableCodec.BuildCell.
        const directStyle = cell.textStyle ?? s.defaultTextStyle;
        const fallback = cell.textBody ? {} : { fontSizePoints: 13.5, ...directStyle,
          color: directStyle?.color?.case ? directStyle.color : { case: "colorRgb", value: header ? "000000" : "0F172A" },
          bold: cell.textStyle ? cell.textStyle.bold : (directStyle?.bold ?? false) || header };
        return `<g data-officekit-table-cell="${r}:${c}" data-officekit-row-span="${r1 - r + 1}" data-officekit-column-span="${c1 - c + 1}"><rect ${box(cellFrame)} fill="${fill}" fill-opacity="${n(opacity)}"/>${borderSvg}${text({ ...node, frame: cellFrame }, cell, prefix, fallback)}</g>`;
      }).join("");
    }).join("");
  }
  function chart(node) {
    const s = node.native, f = node.frame;
    if (s.type !== SpreadsheetChartType.LINE) {
      limit(node, "chart.type", "preview.scene.paint.chart-type", s.type, "opaque");
      return placeholder(node, `chart type ${s.type}: not painted`);
    }
    const fail = (field, message) => {
      limit(node, field, "preview.scene.paint.chart-semantics", message, "unavailable");
      throw new TypeError(message);
    };
    unused(content.get("chart"), s, [...frameFields, "frameTransform", "type", "categories", "series", "xAxis", "yAxis", "lineOptions", "displayBlanksAs"], node, "chart.");
    if (s.comboSeries.length || s.secondaryXAxis || s.secondaryYAxis)
      fail("chart.comboSeries", "A simple line plot cannot replace mixed/secondary-axis topology.");
    if (s.lineOptions) {
      unused(SpreadsheetChartLineOptionsArtifactSchema, s.lineOptions, ["grouping", "smooth"], node, "chart.lineOptions.");
      if (s.lineOptions.grouping !== undefined && ![0, 1].includes(s.lineOptions.grouping))
        fail("chart.lineOptions.grouping", "Stacked lines require cumulative coordinates, not ordinary lines.");
      if (s.lineOptions.smooth === true) fail("chart.lineOptions.smooth", "Smooth line interpolation is not yet mapped.");
    }
    if (s.grouping && s.grouping !== "standard") fail("chart.grouping", "Nonstandard line grouping is not yet mapped.");
    const blank = s.displayBlanksAs ?? "gap";
    if (!["gap", "zero", "span"].includes(blank)) fail("chart.displayBlanksAs", "Unknown native blank-display strategy.");
    const series = s.series.map((entry, si) => {
      const prefix = `chart.series[${si}]`;
      unused(SpreadsheetChartSeriesArtifactSchema, entry, ["name", "values", "missingValueIndexes", "line", "marker"], node, `${prefix}.`);
      if (entry.xValues.length || entry.bubbleSizes.length) fail(prefix, "Category-line data cannot reinterpret numeric X or size channels.");
      if (entry.values.length !== s.categories.length) fail(`${prefix}.values`, "Category and logical value counts differ.");
      let previous = -1;
      for (const i of entry.missingValueIndexes) {
        if (!Number.isSafeInteger(i) || i <= previous || i >= entry.values.length) fail(`${prefix}.missingValueIndexes`, "Missing indexes must be sorted, unique and in range.");
        previous = i;
      }
      const missing = new Set(entry.missingValueIndexes);
      entry.values.forEach((value, i) => {
        if (!Number.isFinite(value)) fail(`${prefix}.values[${i}]`, "Nonfinite chart value.");
        if (missing.has(i) && value !== 0) fail(`${prefix}.missingValueIndexes`, "A missing native observation must carry its canonical zero placeholder.");
      });
      if (missing.size && blank !== "gap") fail("chart.displayBlanksAs", `Explicit ${blank} would transform missing observations; this painter has not implemented separately labeled display-policy evidence.`);
      return { entry, si, prefix, missing };
    });
    const xa = s.xAxis, ya = s.yAxis;
    if (xa) {
      unused(SpreadsheetChartAxisArtifactSchema, xa, ["reverse"], node, "chart.xAxis.");
      if ([xa.minimum, xa.maximum, xa.logBase].some(v => v !== undefined)) fail("chart.xAxis", "Numeric category-axis bounds cannot be treated as evenly spaced labels.");
    }
    if (ya) unused(SpreadsheetChartAxisArtifactSchema, ya, ["minimum", "maximum", "reverse", "logBase"], node, "chart.yAxis.");
    const logBase = ya?.logBase;
    if (logBase !== undefined && (!Number.isFinite(logBase) || logBase < 2 || logBase > 1000)) fail("chart.yAxis.logBase", "Invalid logarithm base.");
    const scaleValue = value => {
      if (logBase !== undefined && value <= 0) fail("chart.yAxis.logBase", "Nonpositive observations/bounds cannot be shown on a logarithmic value axis.");
      return logBase === undefined ? value : Math.log(value) / Math.log(logBase);
    };
    // Missing native zeros never participate in the observed domain.
    const observed = series.flatMap(({ entry, missing }) => entry.values.filter((_, i) => !missing.has(i)));
    let low = Infinity, high = -Infinity;
    for (const value of observed) { const scaled = scaleValue(value); low = Math.min(low, scaled); high = Math.max(high, scaled); }
    if (!observed.length) { low = 0; high = 1; }
    const explicitLow = ya?.minimum !== undefined, explicitHigh = ya?.maximum !== undefined;
    if (explicitLow) low = scaleValue(numeric(ya.minimum));
    if (explicitHigh) high = scaleValue(numeric(ya.maximum));
    if (low >= high) {
      if (explicitLow && explicitHigh) fail("chart.yAxis", "Explicit minimum must be less than maximum.");
      const pad = Math.max(1, Math.abs(explicitLow ? low : high) * .1);
      if (!explicitLow) low = high - pad;
      if (!explicitHigh) high = low + pad * (explicitLow ? 1 : 2);
    }
    if (![low, high, high - low].every(Number.isFinite)) fail("chart.yAxis", "Unrepresentable chart range.");
    const plot = { x: f.x + f.width * .1, y: f.y + f.height * .15, width: f.width * .8, height: f.height * .7 };
    if (plot.width <= 0 || plot.height <= 0) fail("chart", "Nonpositive line-plot extent.");
    limit(node, "chart", "preview.scene.paint.chart-layout", "Native values and explicit scales are consumed; plot margins, automatic bounds, title/legend/ticks and host layout are not fully mapped.");
    const point = (value, i) => {
      let xr = s.categories.length <= 1 ? .5 : i / (s.categories.length - 1), yr = (scaleValue(value) - low) / (high - low);
      if (xa?.reverse === true) xr = 1 - xr;
      if (ya?.reverse !== true) yr = 1 - yr;
      return { x: plot.x + xr * plot.width, y: plot.y + yr * plot.height };
    };
    const output = series.map(({ entry, si, prefix, missing }) => {
      const line = entry.line, colors = ["#2563EB", "#B45309", "#047857", "#9333EA"];
      let color = colors[si % colors.length], width = 1.5, alpha = 1;
      if (line) {
        unused(SpreadsheetChartLineStyleArtifactSchema, line, ["color", "widthPoints", "opacityThousandthPercent", "dashStyle", "cap", "join"], node, `${prefix}.line.`);
        if (line.color) unused(SpreadsheetColorSchema, line.color, ["rgb"], node, `${prefix}.line.color.`);
        if (line.color?.source.case === "rgb") color = rgb(line.color.source.value);
        if (line.widthPoints !== undefined) width = numeric(line.widthPoints);
        alpha = sceneOpacity(line.opacityThousandthPercent ?? 100000);
      }
      if (!line || line.color?.source.case !== "rgb" || line.widthPoints === undefined)
        limit(node, `${prefix}.line`, "preview.scene.paint.chart-inherited-paint", "Review palette/width used where native line paint remains inherited.");
      const paint = linePaint(node, `${prefix}.line`, color, width, alpha,
        ["solid", "solid", "dashed", "dotted", "dash-dot", "dash-dot-dot"][line?.dashStyle ?? 0], line?.cap, line?.join);
      const segments = []; let segment = [];
      const missingEvidence = [];
      entry.values.forEach((value, i) => {
        if (missing.has(i)) {
          if (segment.length) segments.push(segment);
          segment = [];
          missingEvidence.push(`<g data-officekit-missing-point="${i}"><title>${esc(s.categories[i])}: missing observation</title></g>`);
        } else segment.push({ ...point(value, i), value, index: i });
      });
      if (segment.length) segments.push(segment);
      const marker = entry.marker;
      if (marker) unused(SpreadsheetChartMarkerArtifactSchema, marker, ["symbol", "size", "fill", "fillOpacityThousandthPercent"], node, `${prefix}.marker.`);
      const symbol = marker?.symbol ?? 1, radius = (marker?.size ?? 5) / 2;
      if (!Number.isFinite(radius) || radius <= 0) fail(`${prefix}.marker.size`, "Invalid marker size.");
      if (marker?.fill) unused(SpreadsheetColorSchema, marker.fill, ["rgb"], node, `${prefix}.marker.fill.`);
      const markerFill = marker?.fill?.source.case === "rgb" ? rgb(marker.fill.source.value) : color;
      if (symbol !== 1 && ![2, 3, 4, 5, 6].includes(symbol))
        limit(node, `${prefix}.marker.symbol`, "preview.scene.paint.chart-marker", symbol);
      const geometry = segments.map(points => {
        const path = points.length > 1 ? `<path data-officekit-line-segment="${points[0].index}:${points.at(-1).index}" d="${points.map((p, i) => `${i ? "L" : "M"} ${n(p.x)} ${n(p.y)}`).join(" ")}" fill="none" ${paint}/>` : "";
        const dots = points.map(p => {
          let mark = "";
          if ([2, 3].includes(symbol)) mark = `<circle cx="${n(p.x)}" cy="${n(p.y)}" r="${n(radius)}"/>`;
          if (symbol === 4) mark = `<rect ${box({ x: p.x - radius, y: p.y - radius, width: radius * 2, height: radius * 2 })}/>`;
          if (symbol === 5) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y)} L ${n(p.x)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y)} Z"/>`;
          if (symbol === 6) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y + radius)} Z"/>`;
          if (!mark && points.length === 1) {
            limit(node, `${prefix}.values[${p.index}]`, "preview.scene.paint.chart-isolated-review-point", "Isolated observation shown as a hollow review dot, not an authored marker.");
            mark = `<circle data-officekit-review-point="isolated" cx="${n(p.x)}" cy="${n(p.y)}" r="3" fill="none" stroke="${color}" stroke-dasharray="1 1"/>`;
          }
          return `<g data-officekit-point="${p.index}" data-officekit-value="${n(p.value)}" fill="${markerFill}" fill-opacity="${n(sceneOpacity(marker?.fillOpacityThousandthPercent ?? 100000))}"><title>${esc(entry.name)} / ${esc(s.categories[p.index])}: ${n(p.value)}</title>${mark}</g>`;
        }).join("");
        return path + dots;
      }).join("");
      return `<g data-officekit-series="${si}" data-officekit-series-name="${esc(entry.name)}">${missingEvidence.join("")}${geometry}</g>`;
    }).join("");
    // Nested SVG supplies a local clipping viewport without globally colliding
    // clipPath IDs when different pages/nested groups reuse native IDs.
    return `<g data-officekit-chart="line" data-officekit-blank-policy="${blank}" data-officekit-scale-min="${n(low)}" data-officekit-scale-max="${n(high)}" data-officekit-log-base="${logBase ?? "linear"}"><svg ${box(plot)} viewBox="${n(plot.x)} ${n(plot.y)} ${n(plot.width)} ${n(plot.height)}" overflow="hidden">${output}</svg>${observed.length ? "" : `<text x="${n(plot.x)}" y="${n(plot.y + 12)}" font-size="10">No observed data</text>`}</g>`;
  }
  function draw(node) {
    const identity = `data-officekit-native-id="${esc(node.nativeId)}" data-officekit-scene-path="${esc(node.scenePath)}"${node.semanticId ? ` data-officekit-id="${esc(node.semanticId)}"` : ""}`;
    let svg;
    try {
      if (node.hidden === true) return `<g ${identity} display="none"/>`;
      if (!node.frame) throw new TypeError("Missing native frame");
      for (const value of Object.values(node.frame)) numeric(value);
      if (node.frame.width < 0 || node.frame.height < 0) throw new RangeError("Negative native frame");
      if (node.kind === "group") {
        const f = node.frame, c = node.childFrame;
        unused(content.get("group"), node.native, [...frameFields, "childLeftEmu", "childTopEmu", "childWidthEmu", "childHeightEmu", "frameTransform", "children"], node, "group.");
        if (!c || c.width <= 0 || c.height <= 0) throw new RangeError("Nonpositive native group child extent");
        const matrix = `translate(${n(f.x)} ${n(f.y)}) scale(${n(f.width / c.width)} ${n(f.height / c.height)}) translate(${n(-c.x)} ${n(-c.y)})`;
        svg = `<g transform="${matrix}">${node.children.map(draw).join("")}</g>`;
      } else if (node.kind === "shape") svg = shape(node);
      else if (node.kind === "image") svg = image(node);
      else if (node.kind === "connector") svg = connector(node);
      else if (node.kind === "table") svg = table(node);
      else if (node.kind === "chart") svg = chart(node);
      else { limit(node, node.kind, "preview.scene.paint.content", node.kind, "opaque"); svg = placeholder(node, `${node.kind}: not painted`); }
      return `<g ${identity} transform="${frameTransform(node.frame, node.transform)}">${svg}</g>`;
    } catch (error) {
      // Do not erase diagnostics already emitted for this node's descendants.
      limit(node, "", "preview.scene.paint.failed", error.message, "unavailable");
      let fallback = `<title>${esc(error.message)}</title>`;
      try { fallback += placeholder(node, "native drawing unavailable"); } catch { /* Invalid frame: page warning remains visible. */ }
      return `<g ${identity}>${fallback}</g>`;
    } finally {
      // Native element metadata (including animation/source-related fields)
      // remains explicitly unassessed, even when a primitive was drawn.
      unused(PresentationElementSchema, node.element, ["id", "hidden", node.kind], node);
    }
  }
  const pages = view.pages.map(page => {
    const pageNode = { ...page, path: "$" };
    limit(pageNode, "", "preview.scene.paint.integration-pending", "Internal scene painter: production G-11/G-12 integration and complete field coverage remain pending.");
    unused(PresentationSlideSchema, page.native, ["id", "elements", "background", "hidden"], pageNode);
    const background = page.native.background;
    if (background) unused(PresentationBackgroundSchema, background, ["colorRgb", "opacityThousandthPercent", "solid"], pageNode, "background.");
    const fill = background?.color.case === "colorRgb" ? rgb(background.color.value, "#FFFFFF") : "#FFFFFF";
    const body = page.nodes.map(draw).join("");
    const pageDiagnostics = diagnostics.filter(d => d.scenePath === page.scenePath || d.scenePath.startsWith(`${page.scenePath}.`));
    const status = aggregatePreviewStatus(pageDiagnostics.map(d => d.status));
    const reliability = previewReliability(pageDiagnostics, status);
    const bannerHeight = Math.min(24, view.canvas.height * .12);
    const banner = `<g data-officekit-review="${reliability.status}"><rect width="${n(view.canvas.width)}" height="${n(bannerHeight)}" fill="${reliability.status === "failed" ? "#991B1B" : "#92400E"}"/><text x="2" y="${n(bannerHeight * .7)}" font-size="${n(bannerHeight * .5)}" fill="#FFFFFF">INTERNAL SCENE PREVIEW · ${esc(reliability.status)}</text></g>`;
    return Object.freeze({ pageId: page.pageId, nativeId: page.nativeId, hidden: page.hidden,
      status, reliability, diagnostics: Object.freeze(pageDiagnostics),
      svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${n(view.canvas.width)}" height="${n(view.canvas.height)}" viewBox="0 0 ${n(view.canvas.width)} ${n(view.canvas.height)}"><rect width="100%" height="100%" fill="${fill}" fill-opacity="${n(sceneOpacity(background?.opacityThousandthPercent ?? 100000))}"/>${body}${banner}</svg>` });
  });
  const status = aggregatePreviewStatus(diagnostics.map(d => d.status));
  return Object.freeze({ renderer: "officekit-native-scene-svg-internal", scene: view.scene, canvas: view.canvas,
    pages: Object.freeze(pages), diagnostics: Object.freeze(diagnostics), status, reliability: previewReliability(diagnostics, status) });
}
