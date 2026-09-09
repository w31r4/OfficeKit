// Internal scene painter under construction. The production preview switches
// only after G-01's assessment/publication and end-to-end gates are complete.
// No PPJ interpretation, OOXML parsing, filesystem or raster backend here.
import { isFieldSet } from "@bufbuild/protobuf";
import presetProfiles from "./preset-geometry-profiles.json" with { type: "json" };
import { PresentationElementSchema, PresentationSlideSchema, PresentationTextBodySchema,
  PresentationTextParagraphSchema, PresentationTextRunSchema, PresentationTextStyleSchema,
  PresentationBackgroundSchema, PresentationTableRowSchema, PresentationTableCellSchema,
  PresentationTableCellFillSchema, PresentationTableCellBordersSchema,
  SpreadsheetChartLineStyleArtifactSchema, SpreadsheetColorSchema, SpreadsheetChartType,
  SpreadsheetChartSeriesArtifactSchema, SpreadsheetChartAxisArtifactSchema,
  SpreadsheetChartLineOptionsArtifactSchema, SpreadsheetChartMarkerArtifactSchema,
  SpreadsheetChartPointStyleArtifactSchema, SpreadsheetChartSurfaceFillSchema } from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { createPpjSceneView, scenePoints, sceneOpacity, sceneFontPoints } from "./preview-scene-view.mjs";
import { escapePreviewText as esc, previewDiagnostic, previewAssessment } from "./preview-diagnostics.mjs";
import { ppjPreviewSceneIdentity } from "./preview-scene.mjs";
import { assessPpjPreviewInput } from "./preview-input-assessment.mjs";

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
const arcN = value => {
  const number = numeric(value), rounded = Math.round(number);
  return Math.abs(number - rounded) < 1e-12 ? String(rounded) : String(number);
};
const rgb = (value, fallback = "none") => /^[0-9a-f]{6}$/iu.test(value || "") ? `#${value}` : fallback;
const box = f => `x="${n(f.x)}" y="${n(f.y)}" width="${n(f.width)}" height="${n(f.height)}"`;
// roundRect: a = pin(0, adj, 50000), radius = min(w,h) * a / 100000.
// Formula source is the pinned preset definition referenced by presetProfiles.
function roundedRectangle(f, adjustments = [], paint = "") {
  if (adjustments.length > 1 || adjustments.some(v => !Number.isInteger(v) ||
      v < presetProfiles.minimumValue || v > presetProfiles.maximumValue)) throw new TypeError("Invalid roundRect adjustment");
  const adjustment = adjustments[0] ?? presetProfiles.profiles.roundRect.defaults[0];
  const radius = Math.min(f.width, f.height) * Math.max(0, Math.min(50000, adjustment)) / 100000;
  return `<rect ${box(f)} rx="${n(radius)}" ry="${n(radius)}" ${paint}/>`;
}

/** Literal native paths only. An unresolved command fails the whole path,
 * never drops one segment and rejoins unrelated endpoints. Coordinates are
 * mapped before painting so anisotropic viewport scaling does not scale pens.
 *
 * DrawingML arcTo uses a view angle (0 degrees is the positive x axis and
 * positive sweep is clockwise in the document's y-down coordinate system).
 * The current pen position fixes the ellipse centre; the angle must therefore
 * be converted to the ellipse parameter angle before emitting SVG's A
 * command. This is the same distinction made by the established OOXML
 * implementations, and matters whenever the two radii differ. */
export function nativePathData(path, frame) {
  const width = numeric(path.width), height = numeric(path.height);
  if (width < 0 || height < 0) throw new RangeError("Negative path viewport");
  const frameX = numeric(frame.x), frameY = numeric(frame.y);
  const frameWidth = numeric(frame.width), frameHeight = numeric(frame.height);
  const sx = width === 0 ? 1 / 12700 : frameWidth / width;
  const sy = height === 0 ? 1 / 12700 : frameHeight / height;
  const localPoint = value => {
    if (!value || value.xReference !== undefined || value.yReference !== undefined)
      throw new TypeError("Unresolved path point/reference");
    return { x: numeric(value.x), y: numeric(value.y) };
  };
  const mappedPoint = value => ({ x: frameX + value.x * sx, y: frameY + value.y * sy });
  const ellipseParameterAngle = (viewAngle, radiusX, radiusY) =>
    Math.atan2(radiusX * Math.sin(viewAngle), radiusY * Math.cos(viewAngle));
  const arc = (value, current) => {
    if (!value || value.widthRadiusReference !== undefined || value.heightRadiusReference !== undefined ||
        value.startAngleReference !== undefined || value.sweepAngleReference !== undefined)
      throw new TypeError("Unresolved arc value/reference");
    const radiusX = numeric(value.widthRadius), radiusY = numeric(value.heightRadius);
    if (!(radiusX > 0) || !(radiusY > 0)) throw new RangeError("Nonpositive arc radius");
    const startUnits = numeric(value.startAngle), sweepUnits = numeric(value.sweepAngle);
    const fullTurn = 360 * 60000;
    if (!Number.isInteger(startUnits) || !Number.isInteger(sweepUnits) ||
        sweepUnits === 0 || Math.abs(sweepUnits) > fullTurn)
      throw new RangeError("Invalid arc angle");
    if (!(sx > 0) || !(sy > 0)) throw new RangeError("Nonpositive arc extent");
    const start = startUnits / 60000 * Math.PI / 180;
    const sweep = sweepUnits / 60000 * Math.PI / 180;
    const startParameter = ellipseParameterAngle(start, radiusX, radiusY);
    const startOffset = { x: radiusX * Math.cos(startParameter), y: radiusY * Math.sin(startParameter) };
    const center = { x: current.x - startOffset.x, y: current.y - startOffset.y };
    const endParameter = ellipseParameterAngle(start + sweep, radiusX, radiusY);
    let parameterSweep = endParameter - startParameter;
    const direction = sweep >= 0 ? 1 : -1;
    // atan2 wraps at +/-pi. Unwrap into the direction selected by swAng;
    // the absolute sweep is bounded by one full turn by the codec contract.
    while (direction > 0 && parameterSweep <= 0) parameterSweep += Math.PI * 2;
    while (direction < 0 && parameterSweep >= 0) parameterSweep -= Math.PI * 2;
    if (Math.abs(Math.abs(sweep) - Math.PI * 2) < 1e-12) parameterSweep = direction * Math.PI * 2;
    if (Math.abs(parameterSweep) > Math.PI * 2 + 1e-9)
      throw new RangeError("Arc sweep exceeds one turn");
    const end = Math.abs(Math.abs(sweep) - Math.PI * 2) < 1e-12
      ? current
      : { x: center.x + radiusX * Math.cos(startParameter + parameterSweep),
          y: center.y + radiusY * Math.sin(startParameter + parameterSweep) };
    const rx = radiusX * sx, ry = radiusY * sy;
    const mapped = p => mappedPoint(p);
    const fmt = p => {
      const m = mapped(p);
      return `${arcN(m.x)} ${arcN(m.y)}`;
    };
    const sweepFlag = direction > 0 ? 1 : 0;
    const largeArc = Math.abs(parameterSweep) > Math.PI + 1e-12 ? 1 : 0;
    const emit = (target, large) => `A ${arcN(rx)} ${arcN(ry)} 0 ${large} ${sweepFlag} ${fmt(target)}`;
    // SVG cannot represent a complete turn with one A command because its
    // start and end points would coincide. Split it into two half turns.
    let d;
    if (Math.abs(parameterSweep) > Math.PI * 2 - 1e-12) {
      const middle = { x: center.x + radiusX * Math.cos(startParameter + direction * Math.PI),
        y: center.y + radiusY * Math.sin(startParameter + direction * Math.PI) };
      d = `${emit(middle, 0)} ${emit(end, 0)}`;
    } else {
      d = emit(end, largeArc);
    }
    return { d, end };
  };
  let started = false;
  let current, subpathStart;
  return path.commands.map(({ command }) => {
    if (command.case === "moveTo") {
      const value = localPoint(command.value);
      started = true; current = subpathStart = value;
      return `M ${fmtPoint(mappedPoint(value))}`;
    }
    if (!started) throw new TypeError("Path requires moveTo");
    switch (command.case) {
      case "lineTo": {
        const value = localPoint(command.value); current = value;
        return `L ${fmtPoint(mappedPoint(value))}`;
      }
      case "cubicBezierTo": {
        const control1 = localPoint(command.value.control1), control2 = localPoint(command.value.control2);
        const end = localPoint(command.value.end); current = end;
        return `C ${fmtPoint(mappedPoint(control1))} ${fmtPoint(mappedPoint(control2))} ${fmtPoint(mappedPoint(end))}`;
      }
      case "quadraticBezierTo": {
        const control = localPoint(command.value.control), end = localPoint(command.value.end); current = end;
        return `Q ${fmtPoint(mappedPoint(control))} ${fmtPoint(mappedPoint(end))}`;
      }
      case "arcTo": {
        const result = arc(command.value, current);
        current = result.end;
        return result.d;
      }
      case "close":
        if (command.value !== true) throw new TypeError("Invalid close command");
        current = subpathStart;
        return "Z";
      default: throw new TypeError(`Unpainted native path command: ${command.case}`);
    }
  }).join(" ");
}

const fmtPoint = value => `${n(value.x)} ${n(value.y)}`;

function frameTransform(f, transform) {
  if (!transform) return "";
  const cx = f.x + f.width / 2, cy = f.y + f.height / 2;
  return `translate(${n(cx)} ${n(cy)}) rotate(${n(transform.rotation ?? 0)}) scale(${transform.flipH ? -1 : 1} ${transform.flipV ? -1 : 1}) translate(${n(-cx)} ${n(-cy)})`;
}

/** Actual native-state-to-SVG drawing; deliberately not a public publication
 * receipt. No support promotion or G-11 rule retirement is implied. */
export function paintPpjSceneSvg(receipt, { assessInput = false } = {}) {
  const view = createPpjSceneView(receipt), diagnostics = [...view.diagnostics];
  const visitedNodes = new Set();
  const hiddenScenePaths = new Set();
  const transformedScenePaths = new Set();
  let imageMaskSequence = 0;
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
    const properties = body?.bodyProperties;
    const anchor = properties?.anchor?.case === "verticalAnchor" ? properties.anchor.value : undefined;
    const anchored = ["top", "center", "bottom"].includes(anchor);
    if (body) {
      unused(PresentationTextBodySchema, body, ["paragraphs", "bodyProperties"], node, `${ownerField}.textBody.`);
      if (body.bodyProperties) unused(PresentationTextBodySchema.fields.find(field => field.localName === "bodyProperties").message,
        body.bodyProperties, ["leftInsetEmu", "rightInsetEmu", "topInsetEmu", ...(anchored ? ["verticalAnchor", "bottomInsetEmu"] : [])], node, `${ownerField}.textBody.bodyProperties.`);
    }
    if (anchor !== undefined && !anchored) {
      limit(node, `${ownerField}.textBody.bodyProperties.verticalAnchor`, "preview.scene.paint.text-anchor", anchor, "unavailable");
      return placeholder(node, "Text anchor unavailable");
    }
    const inset = (key, selected, fallback) => properties?.[key]?.case === selected ? scenePoints(properties[key].value) : fallback;
    const left = f.x + inset("leftInset", "leftInsetEmu", 7.2), right = f.x + f.width - inset("rightInset", "rightInsetEmu", 7.2);
    let y = f.y + inset("topInset", "topInsetEmu", 3.6);
    const top = y;
    const paintedText = paragraphs.map((paragraph, pi) => {
      const defaults = paragraph.defaultRunStyle?.case === "defaultRunProperties" ? paragraph.defaultRunStyle.value : {};
      const prefix = `${ownerField}.textBody.paragraphs[${pi}].`;
      const leftAligned = !paragraph.alignment || paragraph.alignment === "left";
      if (body) {
        unused(PresentationTextParagraphSchema, paragraph, ["runs", "alignment", "defaultRunProperties", "lineSpacingPoints", "spaceBeforePoints", "spaceAfterPoints", "marginLeftEmu", ...(leftAligned ? ["indentEmu"] : [])], node, prefix);
        if (defaults.$typeName) unused(PresentationTextStyleSchema, defaults,
          ["fontSizePoints", "fontFamily", "bold", "italic", "colorRgb", "colorOpacityThousandthPercent", "underline", "strike", "fontBaselinePercent", "fontSpacingPoints"], node, `${prefix}defaultRunProperties.`);
      }
      const align = paragraph.alignment;
      if (align && !["left", "center", "right"].includes(align)) limit(node, `${prefix}alignment`, "preview.scene.paint.text-alignment", align);
      const paragraphLeft = left + (paragraph.leftMargin?.case === "marginLeftEmu" ? scenePoints(paragraph.leftMargin.value) : 0);
      const firstLineIndent = leftAligned && paragraph.indentation?.case === "indentEmu" ? scenePoints(paragraph.indentation.value) : 0;
      const x = align === "center" ? (paragraphLeft + right) / 2 : align === "right" ? right : paragraphLeft;
      const pointSpacing = (field, selected) => {
        const choice = paragraph[field];
        if (choice?.case !== selected) return undefined;
        const value = choice.value;
        if (!Number.isFinite(value) || value < 0 || field === "lineSpacing" && value === 0) {
          limit(node, `${prefix}${selected}`, "preview.scene.paint.paragraph-spacing", value, "unavailable");
          throw new RangeError("Invalid native paragraph spacing");
        }
        return value;
      };
      const lineSpacing = pointSpacing("lineSpacing", "lineSpacingPoints");
      const spaceBefore = pointSpacing("spaceBefore", "spaceBeforePoints") ?? 0;
      const spaceAfter = pointSpacing("spaceAfter", "spaceAfterPoints") ?? 0;
      y += spaceBefore;
      const lines = [[]];
      for (const [ri, run] of paragraph.runs.entries()) {
        const rp = `${prefix}runs[${ri}].`;
        if (body) unused(PresentationTextRunSchema, run,
          ["text", "lineBreak", "fontSizePoints", "fontFamily", "bold", "italic", "colorRgb", "colorOpacityThousandthPercent", "underline", "strike", "fontBaselinePercent", "fontSpacingPoints"], node, rp);
        if (run.content.case === "lineBreak") { lines.push([]); continue; }
        if (run.content.case !== "text") { limit(node, `${rp}content`, "preview.scene.paint.text-content", run.content.case); continue; }
        const size = sceneFontPoints(run.fontSizePoints ?? defaults.fontSizePoints ?? fallback.fontSizePoints ?? 18);
        const decorations = [];
        for (const [field, off, on, svgValue] of [["underline", "none", "sng", "underline"], ["strike", "noStrike", "sngStrike", "line-through"]]) {
          const value = run[field] ?? defaults[field] ?? fallback[field];
          if (value === on) decorations.push(svgValue);
          else if (value !== undefined && value !== off)
            limit(node, `${rp}${field}`, "preview.scene.paint.text-decoration", value, "unavailable");
        }
        const baseline = run.fontBaselinePercent ?? defaults.fontBaselinePercent ?? fallback.fontBaselinePercent ?? 0;
        if (!Number.isFinite(baseline) || baseline < -400 || baseline > 400) {
          limit(node, `${rp}fontBaselinePercent`, "preview.scene.paint.text-baseline", baseline, "unavailable");
          throw new RangeError("Invalid native text baseline percentage");
        }
        const shift = -size * baseline / 100;
        const spacing = run.fontSpacingPoints ?? defaults.fontSpacingPoints ?? fallback.fontSpacingPoints;
        if (spacing !== undefined && (!Number.isFinite(spacing) || spacing < -768 || spacing > 768)) {
          limit(node, `${rp}fontSpacingPoints`, "preview.scene.paint.text-spacing", spacing, "unavailable");
          throw new RangeError("Invalid native text spacing in points");
        }
        const defaultColor = defaults.color?.case === "colorRgb" ? defaults.color.value : fallback.color?.case === "colorRgb" ? fallback.color.value : undefined;
        const style = `font-family="${esc(run.fontFamily || defaults.fontFamily || fallback.fontFamily || "sans-serif")}" font-size="${n(size)}" font-weight="${(run.bold ?? defaults.bold ?? fallback.bold) ? "bold" : "normal"}" font-style="${(run.italic ?? defaults.italic ?? fallback.italic) ? "italic" : "normal"}" fill="${rgb(run.colorRgb || defaultColor, "#000000")}" fill-opacity="${n(sceneOpacity(run.colorOpacityThousandthPercent ?? defaults.colorOpacityThousandthPercent ?? fallback.colorOpacityThousandthPercent ?? 100000))}"`;
        // Preserve explicit newlines as line boundaries, not one line per run.
        run.content.value.split(/\r\n|\r|\n/u).forEach((segment, i) => {
          if (i) lines.push([]);
          lines.at(-1).push({ size, shift, spacing, style, decoration: decorations.join(" ") || "none", segment });
        });
      }
      let previousSize = 0;
      const paragraphSvg = lines.map((line, index) => {
        // Each run already resolved its own inheritance above. A default that
        // every run overrides must not enlarge this line's baseline advance.
        const size = line.length ? Math.max(...line.map(run => run.size))
          : sceneFontPoints(defaults.fontSizePoints ?? fallback.fontSizePoints ?? 18);
        y += index === 0 ? size : lineSpacing ?? previousSize * .2 + size;
        previousSize = size;
        let previousShift = 0;
        const spans = line.map(run => {
          // SVG dy changes the current text position. Undo the previous run's
          // shift so following runs (including explicit zero) do not inherit it.
          const dy = run.shift - previousShift;
          previousShift = run.shift;
          return `<tspan text-decoration="${run.decoration}" dy="${n(dy)}"${run.spacing === undefined ? "" : ` letter-spacing="${n(run.spacing)}"`} ${run.style}>${esc(run.segment)}</tspan>`;
        }).join("");
        const svg = `<text x="${n(x + (index === 0 ? firstLineIndent : 0))}" y="${n(y)}" text-anchor="${align === "center" ? "middle" : align === "right" ? "end" : "start"}" xml:space="preserve">${spans}</text>`;
        return svg;
      }).join("");
      y += previousSize * .2 + spaceAfter;
      return paragraphSvg;
    }).join("");
    if (!anchored) return paintedText;
    // Align the complete explicit-line block, including paragraph spacing and
    // our existing logical descent. This is not font-metric/AutoFit evidence:
    // the text-layout limitation remains even when direct anchoring is used.
    const height = y - top, available = f.y + f.height - inset("bottomInset", "bottomInsetEmu", 3.6) - top;
    if (available < 0 || anchor !== "top" && height > available) {
      limit(node, `${ownerField}.textBody.bodyProperties.verticalAnchor`, "preview.scene.paint.text-anchor-overflow",
        "Text block exceeds inset bounds; overflow placement is unresolved", "unavailable");
      return placeholder(node, "Text anchor overflow unavailable");
    }
    const shift = anchor === "center" ? (available - height) / 2 : anchor === "bottom" ? available - height : 0;
    return `<g data-officekit-text-anchor="${anchor}" transform="translate(0 ${n(shift)})">${paintedText}</g>`;
  }
  function shape(node) {
    const s = node.native, f = node.frame;
    unused(content.get("shape"), s, [...frameFields, "geometry", "text", "textBody", "fillRgb", "lineRgb", "lineWidthEmu",
      "fillOpacityThousandthPercent", "lineOpacityThousandthPercent", "lineStyle", "lineCap", "lineJoin", "transform", "customPaths",
      ...(s.geometry === "roundRect" && !s.customPaths.length ? ["presetAdjustments"] : [])], node, "shape.");
    const outline = linePaint(node, "shape", s.lineStyle === "none" ? "none" : rgb(s.lineRgb), scenePoints(s.lineWidthEmu),
      sceneOpacity(s.lineOpacityThousandthPercent ?? 100000), s.lineStyle === "none" ? "solid" : s.lineStyle || "solid", s.lineCap, s.lineJoin, "lineStyle");
    const paint = `fill="${rgb(s.fillRgb)}" fill-opacity="${n(sceneOpacity(s.fillOpacityThousandthPercent ?? 100000))}" ${outline}`;
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
    else if (s.geometry === "roundRect") geometry = roundedRectangle(f, s.presetAdjustments, paint);
    else if (s.geometry === "ellipse") geometry = `<ellipse cx="${n(f.x + f.width / 2)}" cy="${n(f.y + f.height / 2)}" rx="${n(f.width / 2)}" ry="${n(f.height / 2)}" ${paint}/>`;
    else if (["diamond", "flowChartDecision"].includes(s.geometry)) geometry = `<path d="M ${n(f.x + f.width / 2)} ${n(f.y)} L ${n(f.x + f.width)} ${n(f.y + f.height / 2)} L ${n(f.x + f.width / 2)} ${n(f.y + f.height)} L ${n(f.x)} ${n(f.y + f.height / 2)} Z" ${paint}/>`;
    else { limit(node, "shape.geometry", "preview.scene.paint.preset", s.geometry); geometry = placeholder(node, `geometry: ${s.geometry || "unresolved"}`); }
    return geometry + text(node);
  }
  function image(node) {
    const s = node.native, asset = view.asset(s.svgAssetId || s.assetId);
    unused(content.get("image"), s, [...frameFields, "assetId", "svgAssetId", "opacityThousandthPercent", "transform", "altText", "accessibilityTitle", "accessibilityDecorative", "crop", "maskPreset", "maskPresetAdjustments", "customMaskPaths", "border"], node, "image.");
    if (!asset?.data?.byteLength) { limit(node, "image.assetId", "preview.scene.paint.asset", s.assetId, "unavailable"); return placeholder(node, "image unavailable"); }
    // A parameter-free native tile is not a stretch. Until intrinsic sizing
    // and DPI semantics are available, drawing one stretched image invents
    // visible coverage and can conceal missing/repeated content.
    if (s.tiled) {
      limit(node, "image.tiled", "preview.scene.paint.image-tile", "Native tile sizing/DPI is not resolved", "unavailable");
      return placeholder(node, "Tiled image unavailable");
    }
    const href = `data:${asset.contentType};base64,${Buffer.from(asset.data).toString("base64")}`;
    let mask = "";
    const f = node.frame;
    try {
      if (s.maskPresetAdjustments.length && s.maskPreset !== "roundRect") throw new TypeError("Adjusted image preset is not mapped");
      if (s.customMaskPaths.length) {
        if (s.maskPreset) throw new TypeError("Conflicting custom and preset masks");
        mask = s.customMaskPaths.map(p => {
          if (![0, 1].includes(p.fillMode)) throw new TypeError("Custom mask fill mode is not mapped");
          return `<path d="${nativePathData(p, f)}"/>`;
        }).join("");
      } else if (s.maskPreset === "roundRect") mask = roundedRectangle(f, s.maskPresetAdjustments);
      else if (s.maskPreset === "ellipse") mask = `<ellipse cx="${n(f.x + f.width / 2)}" cy="${n(f.y + f.height / 2)}" rx="${n(f.width / 2)}" ry="${n(f.height / 2)}"/>`;
      else if (s.maskPreset === "diamond") mask = `<path d="M ${n(f.x + f.width / 2)} ${n(f.y)} L ${n(f.x + f.width)} ${n(f.y + f.height / 2)} L ${n(f.x + f.width / 2)} ${n(f.y + f.height)} L ${n(f.x)} ${n(f.y + f.height / 2)} Z"/>`;
      else if (s.maskPreset && s.maskPreset !== "rect") throw new TypeError(`Unmapped image mask: ${s.maskPreset}`);
    } catch (error) {
      limit(node, s.customMaskPaths.length ? "image.customMaskPaths" : s.maskPresetAdjustments.length ? "image.maskPresetAdjustments" : "image.maskPreset", "preview.scene.paint.image-mask", error.message, "unavailable");
      return placeholder(node, "Image mask unavailable");
    }
    let border = "";
    if (s.border) {
      const b = s.border;
      try {
        if (b.colorScheme || !/^[0-9a-f]{6}$/iu.test(b.colorRgb)) throw new TypeError("Image border color needs resolved RGB");
        const paint = linePaint(node, "image.border", rgb(b.colorRgb), scenePoints(b.widthEmu),
          sceneOpacity(b.opacityThousandthPercent ?? 100000), b.style || "solid", b.cap, b.join, "style");
        const outline = s.customMaskPaths.length ? s.customMaskPaths.filter(p => p.stroke !== false)
          .map(p => `<path d="${nativePathData(p, f)}"/>`).join("") : mask || `<rect ${box(f)}/>`;
        border = `<g data-officekit-image-border="true" fill="none" ${paint}>${outline}</g>`;
      } catch (error) {
        limit(node, "image.border", "preview.scene.paint.image-border", error.message, "unavailable");
      }
    }
    const masked = markup => {
      if (!mask) return markup + border;
      const id = `officekit-image-mask-${imageMaskSequence++}`;
      return `<defs><clipPath id="${id}" clipPathUnits="userSpaceOnUse">${mask}</clipPath></defs><g clip-path="url(#${id})">${markup}</g>${border}`;
    };
    if (s.crop) {
      const edges = ["left", "top", "right", "bottom"].map(side => s.crop[`${side}ThousandthPercent`]);
      const [left, top, right, bottom] = edges;
      if (edges.some(v => !Number.isInteger(v) || v < -100000 || v > 100000) || left + right >= 100000 || top + bottom >= 100000) {
        limit(node, "image.crop", "preview.scene.paint.image-crop", "Invalid native source rectangle", "unavailable");
        return placeholder(node, "Image crop unavailable");
      }
      const f = node.frame, width = f.width / (1 - (left + right) / 100000), height = f.height / (1 - (top + bottom) / 100000);
      // A nested viewport clips to the picture frame without shared clip IDs.
      // Negative edges leave transparent letterbox space; they do not add pixels.
      return masked(`<svg ${box(f)} viewBox="0 0 ${n(f.width)} ${n(f.height)}" overflow="hidden"><image x="${n(-left / 100000 * width)}" y="${n(-top / 100000 * height)}" width="${n(width)}" height="${n(height)}" href="${esc(href)}" preserveAspectRatio="none" opacity="${n(sceneOpacity(s.opacityThousandthPercent ?? 100000))}"><title>${esc(s.altText || s.accessibilityTitle || "")}</title></image></svg>`);
    }
    return masked(`<image ${box(node.frame)} href="${esc(href)}" preserveAspectRatio="none" opacity="${n(sceneOpacity(s.opacityThousandthPercent ?? 100000))}"><title>${esc(s.altText || s.accessibilityTitle || "")}</title></image>`);
  }
  function linePaint(node, field, color, width, opacity, dash, cap, join, dashField = "dashStyle") {
    if (width < 0) throw new RangeError("Negative native line width");
    const patterns = { solid: [], dashed: [4, 3], dotted: [1, 3], "dash-dot": [4, 3, 1, 3], "dash-dot-dot": [8, 3, 1, 3, 1, 3] };
    if (!Object.hasOwn(patterns, dash) || cap && !["flat", "round", "square"].includes(cap) || join && !["miter", "round", "bevel"].includes(join))
      throw new TypeError("Unsupported native line paint token");
    if (patterns[dash].length) limit(node, `${field}.${dashField}`, "preview.scene.paint.dash-metrics", "Native dash kind retained; exact host dash lengths remain approximate.");
    return `stroke="${color}" stroke-width="${n(width)}" stroke-opacity="${n(opacity)}" stroke-linecap="${cap === "flat" || !cap ? "butt" : cap}" stroke-linejoin="${join || "miter"}"${patterns[dash].length ? ` stroke-dasharray="${patterns[dash].map(v => n(v * width)).join(" ")}"` : ""}`;
  }
  function connector(node) {
    const s = node.native, { start, end } = node.endpoints;
    unused(content.get("connector"), s, ["connectorType", "startXEmu", "startYEmu", "endXEmu", "endYEmu", "lineRgb", "lineWidthEmu", "lineStyle", "lineCap", "lineJoin",
      "lineOpacityThousandthPercent", "startArrow", "endArrow", "startArrowWidth", "startArrowLength", "endArrowWidth", "endArrowLength",
      "startTargetId", "endTargetId", "startConnectionSiteIndex", "endConnectionSiteIndex", "bendAdjustment"], node, "connector.");
    // Consume the actual writer/candidate endpoints. Compiler anchor correctness
    // is a separate factual check; neither frame direction nor nearest objects
    // participate in this routing.
    const points = [start];
    if (s.bendAdjustment !== undefined && (!Number.isInteger(s.bendAdjustment) ||
        s.bendAdjustment < -2147483648 || s.bendAdjustment > 2147483647 || s.connectorType === "straight")) {
      limit(node, "connector.bendAdjustment", "preview.scene.paint.connector-bend", "Bend requires a non-straight connector and a native signed 32-bit integer.", "unavailable");
      return placeholder(node, "Invalid connector bend");
    }
    if (s.connectorType === "elbow") {
      // bentConnector3 adj1 is a fraction of the directed horizontal extent.
      // Signed and >100% adjustments intentionally place the bend outside it.
      const mid = start.x + (end.x - start.x) * ((s.bendAdjustment ?? 50000) / 100000);
      points.push({ x: mid, y: start.y }, { x: mid, y: end.y });
      if (view.scene.origin === 2) limit(node, "connector.connectorType", "preview.scene.paint.imported-route", "Native directed elbow and literal adjustment mapped; imported preset transform provenance is not carried separately.");
    } else if (s.connectorType !== "straight") {
      limit(node, "connector.connectorType", "preview.scene.paint.connector-route", s.connectorType, "unavailable");
      return placeholder(node, `${s.connectorType}: route unavailable`);
    }
    points.push(end);
    const width = scenePoints(s.lineWidthEmu), opacity = sceneOpacity(s.lineOpacityThousandthPercent ?? 100000);
    const color = s.lineStyle === "none" ? "none" : rgb(s.lineRgb);
    const paint = linePaint(node, "connector", color, width, opacity, s.lineStyle === "none" ? "solid" : s.lineStyle || "solid", s.lineCap, s.lineJoin, "lineStyle");
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
  function chartFail(node, field, message) {
    limit(node, field, "preview.scene.paint.chart-semantics", message, "unavailable");
    throw new TypeError(message);
  }
  // Native missing indexes own the distinction between absence and real zero
  // for all category-chart painters; do not reconstruct it from numeric values.
  function categorySeries(node, s, handled) {
    return s.series.map((entry, si) => {
      const prefix = `chart.series[${si}]`, fail = (field, message) => chartFail(node, field, message);
      unused(SpreadsheetChartSeriesArtifactSchema, entry, handled, node, `${prefix}.`);
      if (entry.xValues.length || entry.bubbleSizes.length) fail(prefix, "Category data cannot reinterpret numeric X or size channels.");
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
      return { entry, si, prefix, missing };
    });
  }
  function chartFill(node, surface, legacy, fallback, field) {
    if (surface && legacy) chartFail(node, field, "Conflicting native solid/series fill owners.");
    if (surface) {
      unused(SpreadsheetChartSurfaceFillSchema, surface, ["noFill", "solidRgb", "opacityThousandthPercent"], node, `${field}.`);
      if (surface.fill.case === "noFill" && surface.fill.value === true) return 'fill="none"';
      if (surface.fill.case === "solidRgb") return `fill="${rgb(surface.fill.value)}" fill-opacity="${n(sceneOpacity(surface.opacityThousandthPercent ?? 100000))}"`;
      limit(node, field, "preview.scene.paint.chart-inherited-paint", "Native fill is not mapped; using a labeled review color.");
    } else if (legacy) {
      unused(SpreadsheetColorSchema, legacy, ["rgb"], node, `${field}.`);
      if (legacy.source.case === "rgb") return `fill="${rgb(legacy.source.value)}"`;
      limit(node, field, "preview.scene.paint.chart-inherited-paint", "Native color inheritance remains unresolved.");
    } else limit(node, field, "preview.scene.paint.chart-inherited-paint", "Review palette is not a resolved theme.");
    return `fill="${fallback}"`;
  }
  function chartOutline(node, value, field) {
    if (!value) return 'stroke="none"';
    unused(SpreadsheetChartLineStyleArtifactSchema, value, ["color", "widthPoints", "opacityThousandthPercent", "dashStyle", "cap", "join"], node, `${field}.`);
    if (value.color) unused(SpreadsheetColorSchema, value.color, ["rgb"], node, `${field}.color.`);
    if (value.color?.source.case !== "rgb" || value.widthPoints === undefined)
      limit(node, field, "preview.scene.paint.chart-inherited-paint", "Review outline defaults do not resolve native theme paint.");
    return linePaint(node, field, value.color?.source.case === "rgb" ? rgb(value.color.source.value) : "#64748B",
      value.widthPoints ?? 1, sceneOpacity(value.opacityThousandthPercent ?? 100000),
      ["solid", "solid", "dashed", "dotted", "dash-dot", "dash-dot-dot"][value.dashStyle], value.cap, value.join);
  }
  function circularChart(node) {
    const s = node.native, f = node.frame, ring = s.type === SpreadsheetChartType.DOUGHNUT;
    const fail = (field, message) => chartFail(node, field, message);
    unused(content.get("chart"), s, [...frameFields, "frameTransform", "type", "categories", "series", "firstSliceAngle", "doughnutHoleSize", "title", "titleBody", "displayBlanksAs"], node, "chart.");
    if (s.series.length !== 1 || s.comboSeries.length || s.xAxis || s.yAxis || s.secondaryXAxis || s.secondaryYAxis)
      fail("chart.series", "This circular painter requires one native series without Cartesian/mixed axes; multi-ring topology is not yet mapped.");
    const angle = s.firstSliceAngle ?? 0, hole = ring ? (s.doughnutHoleSize ?? 50) : 0;
    if (!Number.isSafeInteger(angle) || angle < 0 || angle > 360) fail("chart.firstSliceAngle", "Circular start angle must be in 0..360 degrees.");
    if ((!ring && s.doughnutHoleSize !== undefined) || (ring && (!Number.isSafeInteger(hole) || hole < 10 || hole > 90)))
      fail("chart.doughnutHoleSize", "Doughnut hole size must be in 10..90 percent and cannot belong to a pie.");
    const [{ entry, prefix, missing }] = categorySeries(node, s, ["name", "values", "missingValueIndexes", "fill", "seriesFill", "line", "pointStyles", "explosion"]);
    const blank = s.displayBlanksAs ?? "gap";
    if (!["gap", "zero", "span"].includes(blank) || (missing.size && blank !== "gap"))
      fail("chart.displayBlanksAs", "Missing circular observations require gap until explicit display-policy evidence is mapped.");
    const explosionValue = (value, field) => {
      if (value !== undefined && (!Number.isSafeInteger(value) || value < 0 || value > 400))
        fail(field, "Circular explosion must be an integer in 0..400.");
      return value ?? 0;
    };
    const seriesExplosion = explosionValue(entry.explosion, `${prefix}.explosion`);
    let maximum = 0;
    entry.values.forEach((value, i) => {
      if (!missing.has(i) && value < 0) fail(`${prefix}.values[${i}]`, "Negative circular values need an explicit signed display policy, not an implicit absolute value.");
      if (!missing.has(i)) maximum = Math.max(maximum, value);
    });
    const denominator = maximum ? entry.values.reduce((total, value, i) => total + (missing.has(i) ? 0 : value / maximum), 0) : 0;
    const pointStyles = new Map(); let previousStyle = -1;
    for (const [pi, style] of entry.pointStyles.entries()) {
      const field = `${prefix}.pointStyles[${pi}]`;
      unused(SpreadsheetChartPointStyleArtifactSchema, style, ["index", "fill", "line", "explosion"], node, `${field}.`);
      if (!Number.isSafeInteger(style.index) || style.index <= previousStyle || style.index >= entry.values.length || missing.has(style.index))
        fail(`${field}.index`, "Point-style indexes must increase, address a real observation and be in range.");
      previousStyle = style.index;
      explosionValue(style.explosion, `${field}.explosion`);
      pointStyles.set(style.index, { style, field });
    }
    // A present point-level zero overrides series explosion, not inherits it.
    const explosions = entry.values.map((value, i) => missing.has(i) || value === 0 ? 0
      : pointStyles.get(i)?.style.explosion ?? seriesExplosion);
    const maxExplosion = explosions.reduce((maximum, value) => Math.max(maximum, value), 0);
    const palette = ["#2563EB", "#B45309", "#047857", "#9333EA", "#BE123C", "#0E7490"];
    // Fixed review plot, not a second PPJ layout pass. Clockwise from up is
    // the native firstSliceAngle convention. Two arcs also handle a full turn.
    // Review layout follows the radial ring-thickness/bisector construction
    // used by LibreOffice's non-concentric PieChart::createDataPoint. Do not
    // clamp native 100..400 values or claim Office's exact plot fitting.
    const fit = 1 + (1 - hole / 100) * maxExplosion / 100;
    const cx = f.x + f.width / 2, cy = f.y + f.height * .52,
      radius = Math.min(f.width * .4, f.height * .34) / fit, inner = radius * hole / 100;
    if (radius <= 0) fail("chart", "Nonpositive circular plot extent.");
    const point = (r, a) => `${n(cx + r * Math.sin(a * Math.PI / 180))} ${n(cy - r * Math.cos(a * Math.PI / 180))}`;
    limit(node, "chart", "preview.scene.paint.chart-layout", "Circular proportions and direct point paint mapped; fixed plot margins, labels/legend, inheritance and host layout remain incomplete.");
    if (maxExplosion) limit(node, prefix, "preview.scene.paint.chart-explosion-layout", "Radial review layout offsets slices by ring thickness and fits them within the plot; exact Office explosion spacing/concentric layout remains unverified.");
    if (missing.size) limit(node, `${prefix}.missingValueIndexes`, "preview.scene.paint.chart-missing-share", "Fractions describe observed positive values only; missing observations are unknown, not zero.");
    let cumulative = 0;
    const slices = entry.values.map((value, i) => {
      if (missing.has(i)) return `<g data-officekit-missing-point="${i}"><title>${esc(s.categories[i])}: missing observation</title></g>`;
      const fraction = denominator ? (value / maximum) / denominator : 0;
      if (value > 0 && (!(fraction > 0) || cumulative + fraction === cumulative)) fail(`${prefix}.values[${i}]`, "Slice ratio is below representable precision.");
      const start = angle + cumulative * 360;
      cumulative += fraction;
      const end = angle + cumulative * 360, middle = (start + end) / 2;
      const selected = pointStyles.get(i);
      const explosion = selected?.style.explosion ?? seriesExplosion;
      const distance = (radius - inner) * explosions[i] / 100;
      const dx = distance * Math.sin(middle * Math.PI / 180), dy = -distance * Math.cos(middle * Math.PI / 180);
      const paint = selected?.style.fill
        ? chartFill(node, selected.style.fill, null, palette[i % palette.length], `${selected.field}.fill`)
        : chartFill(node, entry.seriesFill, entry.fill, palette[i % palette.length], `${prefix}.${entry.fill ? "fill" : "seriesFill"}`);
      const stroke = chartOutline(node, selected?.style.line ?? entry.line, selected?.style.line ? `${selected.field}.line` : `${prefix}.line`);
      let geometry = "";
      if (fraction > 0) {
        const outerArc = `M ${point(radius, start)} A ${n(radius)} ${n(radius)} 0 0 1 ${point(radius, middle)} A ${n(radius)} ${n(radius)} 0 0 1 ${point(radius, end)}`;
        const d = ring ? `${outerArc} L ${point(inner, end)} A ${n(inner)} ${n(inner)} 0 0 0 ${point(inner, middle)} A ${n(inner)} ${n(inner)} 0 0 0 ${point(inner, start)} Z` : `${outerArc} L ${n(cx)} ${n(cy)} Z`;
        geometry = `<path data-officekit-slice="${i}" d="${d}"${distance ? ` transform="translate(${n(dx)} ${n(dy)})"` : ""} ${paint} ${stroke}/>`;
      }
      return `<g data-officekit-point="${i}" data-officekit-value="${n(value)}" data-officekit-fraction="${n(fraction)}" data-officekit-start-angle="${n(start)}" data-officekit-sweep-angle="${n(fraction * 360)}" data-officekit-explosion="${n(explosion)}" data-officekit-explosion-owner="${selected?.style.explosion !== undefined ? "point" : entry.explosion !== undefined ? "series" : "default"}" data-officekit-offset-x="${n(dx)}" data-officekit-offset-y="${n(dy)}"><title>${esc(entry.name)} / ${esc(s.categories[i])}: ${n(value)} (${n(fraction * 100)}% of observed values)</title>${geometry}</g>`;
    }).join("");
    const heading = text({ ...node, frame: { x: f.x, y: f.y, width: f.width, height: f.height * .16 } }, { text: s.title, textBody: s.titleBody }, "chart");
    const zeroCount = entry.values.filter((v, i) => !missing.has(i) && v === 0).length;
    const counts = `${missing.size} missing; ${zeroCount} zero`;
    const note = !denominator ? `No positive observed data; ${counts}` : missing.size || zeroCount ? `Observed values only; ${counts}` : "";
    return `<g data-officekit-chart="${ring ? "doughnut" : "pie"}" data-officekit-first-slice-angle="${angle}" data-officekit-hole-size="${hole}" data-officekit-outer-radius="${n(radius)}" data-officekit-inner-radius="${n(inner)}" data-officekit-explosion-layout="radial-review">${heading}<g data-officekit-series="0" data-officekit-series-name="${esc(entry.name)}">${slices}</g>${note ? `<text data-officekit-circular-data-note="true" x="${n(f.x + 4)}" y="${n(f.y + f.height - 6)}" font-size="10">${esc(note)}</text>` : ""}</g>`;
  }
  function barChart(node) {
    const s = node.native, f = node.frame, fail = (field, message) => chartFail(node, field, message);
    unused(content.get("chart"), s, [...frameFields, "frameTransform", "type", "categories", "series", "barDirection", "grouping", "gapWidth", "overlap", "varyColors",
      "displayBlanksAs", "xAxis", "yAxis", "showCategoryAxis", "showValueAxis", "title", "titleBody"], node, "chart.");
    // BAR is the shared native enum; column is its default direction. xAxis
    // always owns categories and yAxis owns values, even for horizontal bars.
    const direction = s.barDirection || "column", horizontal = direction === "bar";
    if (!["column", "bar"].includes(direction)) fail("chart.barDirection", "Unknown native bar direction.");
    if (s.comboSeries.length || s.secondaryXAxis || s.secondaryYAxis || s.lineOptions)
      fail("chart", "A category bar plot cannot replace mixed/secondary-axis or line topology.");
    const grouping = s.grouping || "none", percent = grouping === "percent-stacked", stacked = grouping === "stacked" || percent;
    if (!["none", "stacked", "percent-stacked"].includes(grouping)) fail("chart.grouping", "Unknown native bar grouping.");
    const gap = s.gapWidth ?? 150, overlap = s.overlap ?? 0;
    if (!Number.isSafeInteger(gap) || gap < 0 || gap > 500) fail("chart.gapWidth", "Bar gap must be an integer in 0..500 percent.");
    if (!Number.isSafeInteger(overlap) || overlap < -100 || overlap > 100) fail("chart.overlap", "Bar overlap must be an integer in -100..100 percent.");
    const series = categorySeries(node, s, ["name", "values", "missingValueIndexes", "fill", "seriesFill", "line", "pointStyles"]);
    const blank = s.displayBlanksAs ?? "gap", missingCount = series.reduce((total, v) => total + v.missing.size, 0);
    if (!["gap", "zero", "span"].includes(blank) || (missingCount && blank !== "gap"))
      fail("chart.displayBlanksAs", "Missing bars require gap until explicit display-policy evidence is mapped.");
    const xa = s.xAxis, ya = s.yAxis;
    if (xa) {
      unused(SpreadsheetChartAxisArtifactSchema, xa, ["reverse", "visible", "axisLineVisible", "tickLabelsVisible", "tickLabelInterval"], node, "chart.xAxis.");
      if ([xa.minimum, xa.maximum, xa.logBase].some(v => v !== undefined)) fail("chart.xAxis", "Numeric category bounds cannot be treated as equally spaced bands.");
    }
    if (ya) unused(SpreadsheetChartAxisArtifactSchema, ya, ["minimum", "maximum", "reverse", "visible", "axisLineVisible", "tickLabelsVisible"], node, "chart.yAxis.");
    if (ya?.logBase !== undefined) fail("chart.yAxis.logBase", "Logarithmic bars require a separately verified nonzero baseline.");
    const observed = series.flatMap(({ entry, missing }) => entry.values.filter((_, i) => !missing.has(i)));
    // Unknown terms cannot silently become zero in a cumulative sum. Keep the
    // whole incomplete category unpositioned, with original observations below.
    const incomplete = new Set(stacked ? s.categories.flatMap((_, i) => series.some(v => v.missing.has(i)) ? [i] : []) : []);
    const zeroTotal = new Set(), denominators = new Map();
    if (percent) {
      for (const { entry, missing, prefix } of series) for (const [i, value] of entry.values.entries())
        if (!missing.has(i) && value < 0) fail(`${prefix}.values[${i}]`, "Negative percentage stacks need a verified signed-denominator policy; absolute values cannot silently replace signed observations.");
      for (let i = 0; i < s.categories.length; i++) {
        if (incomplete.has(i)) continue;
        let scale = 0;
        for (const { entry } of series) scale = Math.max(scale, entry.values[i]);
        if (scale === 0) { zeroTotal.add(i); continue; }
        // Scaling before addition avoids overflowing a finite set's total.
        let total = 0;
        for (const { entry } of series) total += entry.values[i] / scale;
        denominators.set(i, { scale, total });
      }
    }
    const positive = s.categories.map(() => 0), negative = s.categories.map(() => 0);
    const segments = series.map(({ entry, missing, prefix }) => entry.values.map((value, i) => {
      if (missing.has(i) || incomplete.has(i) || zeroTotal.has(i)) return null;
      const denominator = denominators.get(i), display = percent ? (value / denominator.scale) / denominator.total : value;
      const totals = value < 0 ? negative : positive, start = stacked ? totals[i] : 0, end = start + display;
      if (!Number.isFinite(end) || (value !== 0 && end === start))
        fail(`${prefix}.values[${i}]`, "Cumulative bar endpoint is not representable without losing an observation.");
      if (stacked) totals[i] = end;
      return { start, end, display };
    }));
    for (const i of incomplete) limit(node, "chart.series", "preview.scene.paint.chart-stack-incomplete",
      `Category ${i} (${s.categories[i]}): missing terms make the full stack unknown; values are retained but no cumulative positions are invented.`);
    for (const i of zeroTotal) limit(node, "chart.series", "preview.scene.paint.chart-percent-zero-total",
      `Category ${i} (${s.categories[i]}): the total is zero; observations are retained but percentages are undefined.`);
    // Both clustered and stacked domains include zero, unless explicitly bounded.
    let low = 0, high = 0;
    for (const row of segments) for (const segment of row) if (segment) {
      low = Math.min(low, segment.start, segment.end); high = Math.max(high, segment.start, segment.end);
    }
    if (percent) { low = 0; high = 1; }
    const explicitLow = ya?.minimum !== undefined, explicitHigh = ya?.maximum !== undefined;
    if (explicitLow) low = numeric(ya.minimum);
    if (explicitHigh) high = numeric(ya.maximum);
    if (low >= high) {
      if (explicitLow && explicitHigh) fail("chart.yAxis", "Explicit minimum must be less than maximum.");
      const pad = Math.max(1, Math.abs(explicitLow ? low : high) * .1);
      if (!explicitLow && explicitHigh) low = high - pad;
      else if (!explicitHigh) high = low + pad;
    }
    if (![low, high, high - low].every(Number.isFinite)) fail("chart.yAxis", "Unrepresentable bar range.");
    const plot = { x: f.x + f.width * .1, y: f.y + f.height * .15, width: f.width * .8, height: f.height * .7 };
    if (plot.width <= 0 || plot.height <= 0) fail("chart", "Nonpositive bar plot extent.");
    limit(node, "chart", "preview.scene.paint.chart-layout", "Native category bars mapped; fixed review margins/ticks, full labels/legend, inheritance and host layout remain incomplete.");
    const valueAt = value => {
      let ratio = (value - low) / (high - low);
      if (ya?.reverse === true) ratio = 1 - ratio;
      return horizontal ? plot.x + ratio * plot.width : plot.y + (1 - ratio) * plot.height;
    };
    const categoryExtent = horizontal ? plot.height : plot.width, count = Math.max(1, series.length);
    const band = categoryExtent / Math.max(1, s.categories.length);
    // gapWidth is relative to one bar, not the whole cluster. Negative overlap
    // leaves within-cluster space. Missing entries keep their series slot.
    const thickness = band / (count - (count - 1) * overlap / 100 + gap / 100);
    const stride = thickness * (1 - overlap / 100), cluster = thickness + stride * (count - 1);
    const categoryAt = offset => {
      const position = xa?.reverse === true ? categoryExtent - offset : offset;
      return horizontal ? plot.y + plot.height - position : plot.x + position;
    };
    const palette = ["#2563EB", "#B45309", "#047857", "#9333EA", "#BE123C", "#0E7490"];
    const output = series.map(({ entry, si, prefix, missing }) => {
      const styles = new Map(); let previous = -1;
      for (const [pi, style] of entry.pointStyles.entries()) {
        const field = `${prefix}.pointStyles[${pi}]`;
        unused(SpreadsheetChartPointStyleArtifactSchema, style, ["index", "fill", "line"], node, `${field}.`);
        if (!Number.isSafeInteger(style.index) || style.index <= previous || style.index >= entry.values.length || missing.has(style.index))
          fail(`${field}.index`, "Point-style indexes must increase, address a real observation and be in range.");
        previous = style.index; styles.set(style.index, { style, field });
      }
      const marks = entry.values.map((value, i) => {
        if (missing.has(i)) return `<g data-officekit-missing-point="${i}"><title>${esc(s.categories[i])}: missing observation</title></g>`;
        if (incomplete.has(i) || zeroTotal.has(i)) return `<g data-officekit-point="${i}" data-officekit-value="${n(value)}" data-officekit-stack-position="${zeroTotal.has(i) ? "zero-total" : "unknown"}"><title>${esc(entry.name)} / ${esc(s.categories[i])}: ${n(value)}; ${zeroTotal.has(i) ? "zero total, undefined percentage" : "incomplete stack, position not inferred"}</title></g>`;
        const selected = styles.get(i), fallback = palette[(s.varyColors ? i : si) % palette.length];
        const fill = selected?.style.fill ? chartFill(node, selected.style.fill, null, fallback, `${selected.field}.fill`)
          : chartFill(node, entry.seriesFill, entry.fill, fallback, `${prefix}.${entry.fill ? "fill" : "seriesFill"}`);
        const stroke = chartOutline(node, selected?.style.line ?? entry.line, selected?.style.line ? `${selected.field}.line` : `${prefix}.line`);
        const offset = i * band + (band - cluster) / 2 + si * stride;
        const catStart = Math.min(categoryAt(offset), categoryAt(offset + thickness));
        const segment = segments[si][i];
        const baseline = valueAt(segment.start), endpoint = valueAt(segment.end), extent = Math.abs(endpoint - baseline);
        if (value !== 0 && extent === 0) fail(`${prefix}.values[${i}]`, "Bar length is below representable precision.");
        const rect = horizontal ? { x: Math.min(baseline, endpoint), y: catStart, width: extent, height: thickness }
          : { x: catStart, y: Math.min(baseline, endpoint), width: thickness, height: extent };
        let geometry = `<rect data-officekit-bar="${i}" ${box(rect)} ${fill} ${stroke}/>`;
        if (value === 0) {
          limit(node, `${prefix}.values[${i}]`, "preview.scene.paint.chart-zero-review-mark", "Zero shown by a dashed baseline tick, not a nonzero data bar.");
          const x1 = horizontal ? baseline : catStart, y1 = horizontal ? catStart : baseline;
          geometry = `<path data-officekit-review-point="zero" d="M ${n(x1)} ${n(y1)} L ${n(horizontal ? x1 : x1 + thickness)} ${n(horizontal ? y1 + thickness : y1)}" fill="none" stroke="#64748B" stroke-dasharray="2 2"/>`;
        }
        return `<g data-officekit-point="${i}" data-officekit-value="${n(value)}" data-officekit-baseline="${n(segment.start)}"${Math.min(segment.start, segment.end) < low || Math.max(segment.start, segment.end) > high ? ' data-officekit-point-outside-plot="true"' : ""}${stacked ? ` data-officekit-stack-end="${n(segment.end)}"` : ""}${percent ? ` data-officekit-fraction="${n(segment.display)}"` : ""}><title>${esc(entry.name)} / ${esc(s.categories[i])}: ${n(value)}${percent ? ` (${n(segment.display * 100)}% of complete category)` : ""}</title>${geometry}</g>`;
      }).join("");
      return `<g data-officekit-series="${si}" data-officekit-series-name="${esc(entry.name)}">${marks}</g>`;
    }).join("");
    let axes = "";
    for (const i of new Set([...incomplete, ...zeroTotal])) {
      const values = series.map(({ entry, missing }) => `${entry.name}: ${missing.has(i) ? "?" : n(entry.values[i])}`).join("; ");
      axes += `<text data-officekit-${zeroTotal.has(i) ? "zero-total" : "incomplete"}-stack="${i}" x="${n(horizontal ? plot.x + 4 : categoryAt((i + .5) * band))}" y="${n(horizontal ? categoryAt((i + .5) * band) : plot.y + 12)}" text-anchor="${horizontal ? "start" : "middle"}" font-size="10" fill="#92400E"><title>${esc(values)}</title>${zeroTotal.has(i) ? "Zero total" : "Incomplete"}</text>`;
    }
    const bottom = plot.y + plot.height, right = plot.x + plot.width;
    if ((xa?.visible ?? s.showCategoryAxis ?? true) === true) {
      if (xa?.axisLineVisible !== false) axes += horizontal
        ? `<line data-officekit-axis="category" x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(bottom)}" stroke="#64748B"/>`
        : `<line data-officekit-axis="category" x1="${n(plot.x)}" y1="${n(bottom)}" x2="${n(right)}" y2="${n(bottom)}" stroke="#64748B"/>`;
      const interval = xa?.tickLabelInterval ?? 1;
      if (!Number.isSafeInteger(interval) || interval < 1) fail("chart.xAxis.tickLabelInterval", "Invalid category tick interval.");
      if (xa?.tickLabelsVisible !== false && xa?.tickLabelPosition !== "none") axes += s.categories.map((label, i) => i % interval ? "" :
        `<text data-officekit-category="${i}" x="${n(horizontal ? plot.x - 4 : categoryAt((i + .5) * band))}" y="${n(horizontal ? categoryAt((i + .5) * band) + 3 : bottom + 12)}" text-anchor="${horizontal ? "end" : "middle"}" font-size="10" fill="#334155">${esc(label)}</text>`).join("");
    }
    if ((ya?.visible ?? s.showValueAxis ?? true) === true) {
      if (ya?.axisLineVisible !== false) axes += horizontal
        ? `<line data-officekit-axis="value" x1="${n(plot.x)}" y1="${n(bottom)}" x2="${n(right)}" y2="${n(bottom)}" stroke="#64748B"/>`
        : `<line data-officekit-axis="value" x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(bottom)}" stroke="#64748B"/>`;
      // A signed bar chart needs an explicit zero reference; the middle of an
      // asymmetric range is not the data baseline. This is a review guide,
      // not a claim that native crosses/gridline styling has been resolved.
      if (low < 0 && high > 0) axes += horizontal
        ? `<line data-officekit-bar-zero-baseline="review" x1="${n(valueAt(0))}" y1="${n(plot.y)}" x2="${n(valueAt(0))}" y2="${n(bottom)}" stroke="#CBD5E1" stroke-dasharray="2 2"/>`
        : `<line data-officekit-bar-zero-baseline="review" x1="${n(plot.x)}" y1="${n(valueAt(0))}" x2="${n(right)}" y2="${n(valueAt(0))}" stroke="#CBD5E1" stroke-dasharray="2 2"/>`;
      if (ya?.tickLabelsVisible !== false && ya?.tickLabelPosition !== "none") axes += [low, low < 0 && high > 0 ? 0 : low + (high - low) / 2, high].map(value => {
        return `<text data-officekit-value-tick="${n(value)}" x="${n(horizontal ? valueAt(value) : plot.x - 4)}" y="${n(horizontal ? bottom + 12 : valueAt(value) + 3)}" text-anchor="${horizontal ? "middle" : "end"}" font-size="10" fill="#334155">${percent ? `${n(value * 100)}%` : n(value)}</text>`;
      }).join("");
    }
    const heading = text({ ...node, frame: { x: f.x, y: f.y, width: f.width, height: f.height * .15 } }, { text: s.title, textBody: s.titleBody }, "chart");
    const zeroCount = observed.filter(v => v === 0).length;
    const note = percent
      ? `Complete-category shares; ${missingCount} missing; ${incomplete.size} incomplete; ${zeroTotal.size} zero totals (percentages undefined)`
      : `${observed.length ? "Observed values" : "No observed data"}; ${missingCount} missing; ${zeroCount} zero${incomplete.size ? `; ${incomplete.size} incomplete stacks (not positioned)` : " (dashed review ticks)"}`;
    return `<g data-officekit-chart="${direction}" data-officekit-blank-policy="${blank}" data-officekit-scale-min="${n(low)}" data-officekit-scale-max="${n(high)}" data-officekit-grouping="${grouping}" data-officekit-gap-width="${gap}" data-officekit-overlap="${overlap}" data-officekit-bar-thickness="${n(thickness)}">${heading}${axes}<svg data-officekit-bar-clip="plot" ${box(plot)} viewBox="${n(plot.x)} ${n(plot.y)} ${n(plot.width)} ${n(plot.height)}" overflow="hidden">${output}</svg><text x="${n(f.x + 4)}" y="${n(f.y + f.height - 6)}" font-size="10">${esc(note)}</text></g>`;
  }
  function scatterChart(node) {
    const s = node.native, f = node.frame, fail = (field, message) => chartFail(node, field, message);
    unused(content.get("chart"), s, [...frameFields, "frameTransform", "type", "series", "xAxis", "yAxis", "scatterStyle",
      "displayBlanksAs", "showCategoryAxis", "showValueAxis", "title", "titleBody"], node, "chart.");
    // Consume native IR canonical tokens, not raw ChartML lineMarker.
    const style = s.scatterStyle || "marker", connected = style === "line" || style === "lineWithMarkers";
    if (!["marker", "line", "lineWithMarkers"].includes(style)) fail("chart.scatterStyle", "Smooth scatter interpolation is not mapped.");
    if (s.comboSeries.length || s.secondaryXAxis || s.secondaryYAxis || s.lineOptions || s.grouping && s.grouping !== "none")
      fail("chart", "Numeric scatter cannot reinterpret mixed, stacked or line-chart topology.");
    const blank = s.displayBlanksAs ?? "gap";
    if (!["gap", "zero", "span"].includes(blank)) fail("chart.displayBlanksAs", "Unknown native blank-display strategy.");
    const series = s.series.map((entry, si) => {
      const prefix = `chart.series[${si}]`;
      unused(SpreadsheetChartSeriesArtifactSchema, entry, ["name", "values", "xValues", "missingValueIndexes", "marker", "line"], node, `${prefix}.`);
      if (entry.bubbleSizes.length || entry.xValues.length !== entry.values.length)
        fail(prefix, "Scatter requires equal native X/Y counts and no size channel.");
      const missing = new Set(); let previous = -1;
      for (const i of entry.missingValueIndexes) {
        if (!Number.isSafeInteger(i) || i <= previous || i >= entry.values.length)
          fail(`${prefix}.missingValueIndexes`, "Missing indexes must be sorted, unique and in range.");
        missing.add(i); previous = i;
      }
      entry.values.forEach((value, i) => {
        if (!Number.isFinite(value) || !Number.isFinite(entry.xValues[i])) fail(prefix, "Nonfinite numeric observation.");
        if (missing.has(i) && value !== 0) fail(`${prefix}.missingValueIndexes`, "Missing Y must retain its canonical zero placeholder.");
      });
      if (missing.size && blank !== "gap") fail("chart.displayBlanksAs", "Missing numeric pairs require gap; transformed display policies remain unimplemented.");
      return { entry, si, prefix, missing };
    });
    const observed = series.flatMap(({ entry, missing }) => entry.values.flatMap((y, i) => missing.has(i) ? [] : [{ x: entry.xValues[i], y }]));
    function scale(axis, name) {
      const field = `chart.${name}Axis`;
      if (axis) unused(SpreadsheetChartAxisArtifactSchema, axis, ["minimum", "maximum", "logBase", "reverse", "visible", "axisLineVisible", "tickLabelsVisible"], node, `${field}.`);
      const base = axis?.logBase;
      if (base !== undefined && (!Number.isFinite(base) || base < 2 || base > 1000)) fail(`${field}.logBase`, "Invalid logarithm base.");
      const transform = value => {
        if (!Number.isFinite(value) || base !== undefined && value <= 0) fail(field, "Axis observations and bounds must be finite and positive for logarithmic scales.");
        return base === undefined ? value : Math.log(value) / Math.log(base);
      };
      let low = Infinity, high = -Infinity;
      for (const p of observed) { const value = transform(p[name]); low = Math.min(low, value); high = Math.max(high, value); }
      if (!observed.length) { low = 0; high = 1; }
      const hasLow = axis?.minimum !== undefined, hasHigh = axis?.maximum !== undefined;
      if (hasLow) low = transform(axis.minimum);
      if (hasHigh) high = transform(axis.maximum);
      if (low >= high) {
        if (hasLow && hasHigh) fail(field, "Explicit minimum must be less than maximum.");
        const pad = Math.max(1, Math.abs(hasLow ? low : high) * .1);
        if (!hasLow) low = high - pad;
        if (!hasHigh) high = low + pad * (hasLow ? 1 : 2);
      }
      if (![low, high, high - low].every(Number.isFinite) || high <= low) fail(field, "Unrepresentable numeric range.");
      return { low, high, base, ratio: value => (transform(value) - low) / (high - low),
        tick: ratio => base === undefined ? low + ratio * (high - low) : Math.pow(base, low + ratio * (high - low)) };
    }
    const xs = scale(s.xAxis, "x"), ys = scale(s.yAxis, "y");
    const plot = { x: f.x + f.width * .1, y: f.y + f.height * .15, width: f.width * .8, height: f.height * .7 };
    if (plot.width <= 0 || plot.height <= 0) fail("chart", "Nonpositive numeric plot extent.");
    const point = (x, y) => ({ x: plot.x + (s.xAxis?.reverse ? 1 - xs.ratio(x) : xs.ratio(x)) * plot.width,
      y: plot.y + (s.yAxis?.reverse ? ys.ratio(y) : 1 - ys.ratio(y)) * plot.height });
    limit(node, "chart", "preview.scene.paint.chart-layout", "Native per-series X/Y pairs are mapped; automatic bounds, tick formatting, legend, theme and exact host layout remain review approximations.");
    const output = series.map(({ entry, si, prefix, missing }) => {
      const marker = entry.marker, color = ["#2563EB", "#B45309", "#047857", "#9333EA"][si % 4];
      if (marker) unused(SpreadsheetChartMarkerArtifactSchema, marker, ["symbol", "size", "fill", "fillOpacityThousandthPercent", "line"], node, `${prefix}.marker.`);
      const symbol = marker?.symbol ?? 2, radius = (marker?.size ?? 5) / 2;
      if (!Number.isFinite(radius) || radius <= 0) fail(`${prefix}.marker.size`, "Invalid scatter marker size.");
      if (![1, 2, 3, 4, 5, 6].includes(symbol)) fail(`${prefix}.marker.symbol`, "Scatter marker geometry is not mapped.");
      if (marker?.fill) unused(SpreadsheetColorSchema, marker.fill, ["rgb"], node, `${prefix}.marker.fill.`);
      if (!marker || marker.fill?.source.case !== "rgb") limit(node, `${prefix}.marker`, "preview.scene.paint.chart-inherited-paint", "Scatter marker uses review defaults where native paint remains inherited.");
      const fill = marker?.fill?.source.case === "rgb" ? rgb(marker.fill.source.value) : color;
      const outline = chartOutline(node, marker?.line, `${prefix}.marker.line`);
      const line = entry.line ? chartOutline(node, entry.line, `${prefix}.line`) : `stroke="${color}" stroke-width="1.5"`;
      // Current ChartSpace writer always writes scatter spPr/ln/noFill; IR
      // does not preserve that no-fill as a line object. Do not fabricate a
      // visible default line from scatterStyle alone.
      if (connected && !entry.line) limit(node, `${prefix}.line`, "preview.scene.paint.scatter-line-unresolved", "Scatter style requests connections but native line paint is absent; writer no-fill and inherited lines cannot be distinguished by this scene.", "unavailable");
      let segments = [], segment = [], marks = [];
      entry.values.forEach((y, i) => {
        const x = entry.xValues[i];
        if (missing.has(i)) {
          if (segment.length) segments.push(segment); segment = [];
          marks.push(`<g data-officekit-missing-point="${i}" data-officekit-x-value="${n(x)}"><title>Missing Y observation</title></g>`); return;
        }
        const p = point(x, y), outside = p.x < plot.x || p.x > plot.x + plot.width || p.y < plot.y || p.y > plot.y + plot.height;
        segment.push({ ...p, index: i }); let mark = "";
        if (style !== "line" && symbol !== 1) {
          if ([2, 3].includes(symbol)) mark = `<circle cx="${n(p.x)}" cy="${n(p.y)}" r="${n(radius)}"/>`;
          if (symbol === 4) mark = `<rect ${box({ x: p.x - radius, y: p.y - radius, width: radius * 2, height: radius * 2 })}/>`;
          if (symbol === 5) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y)} L ${n(p.x)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y)} Z"/>`;
          if (symbol === 6) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y + radius)} Z"/>`;
        }
        marks.push(`<g data-officekit-point="${i}" data-officekit-x-value="${n(x)}" data-officekit-value="${n(y)}"${outside ? ' data-officekit-point-outside-plot="true"' : ""} fill="${fill}" fill-opacity="${n(sceneOpacity(marker?.fillOpacityThousandthPercent ?? 100000))}" ${outline}><title>${esc(entry.name)}: (${n(x)}, ${n(y)})</title>${outside ? "" : mark}</g>`);
      });
      if (segment.length) segments.push(segment);
      const paths = connected && entry.line ? segments.filter(points => points.length > 1).map(points => `<path data-officekit-line-segment="${points[0].index}:${points.at(-1).index}" d="${points.map((p, i) => `${i ? "L" : "M"} ${n(p.x)} ${n(p.y)}`).join(" ")}" fill="none" ${line}/>`).join("") : "";
      if ((style === "line" || symbol === 1) && segments.some(points => !connected || points.length === 1))
        limit(node, prefix, "preview.scene.paint.scatter-invisible-observation", "Some observations have neither a marker nor a connecting segment; retained in evidence, not silently reported visible.", "unavailable");
      return `<g data-officekit-series="${si}" data-officekit-series-name="${esc(entry.name)}"><svg ${box(plot)} viewBox="${n(plot.x)} ${n(plot.y)} ${n(plot.width)} ${n(plot.height)}" overflow="hidden">${paths}</svg>${marks.join("")}</g>`;
    }).join("");
    let axes = "";
    for (const [name, axis, domain, visible] of [["x", s.xAxis, xs, s.showCategoryAxis], ["y", s.yAxis, ys, s.showValueAxis]]) {
      if ((axis?.visible ?? visible ?? true) !== true) continue;
      if (axis?.axisLineVisible !== false) axes += `<line data-officekit-axis="${name}" x1="${n(plot.x)}" y1="${n(plot.y + plot.height)}" x2="${n(name === "x" ? plot.x + plot.width : plot.x)}" y2="${n(name === "x" ? plot.y + plot.height : plot.y)}" stroke="#64748B"/>`;
      if (axis?.tickLabelsVisible !== false && axis?.tickLabelPosition !== "none") axes += [0, .5, 1].map(r => {
        const value = domain.tick(r), ratio = axis?.reverse ? 1 - r : r;
        if (!Number.isFinite(value)) fail(`chart.${name}Axis`, "Unrepresentable numeric tick.");
        return `<text data-officekit-${name}-tick="${n(value)}" x="${n(name === "x" ? plot.x + ratio * plot.width : plot.x - 4)}" y="${n(name === "x" ? plot.y + plot.height + 12 : plot.y + (1 - ratio) * plot.height + 3)}" text-anchor="${name === "x" ? "middle" : "end"}" font-size="10">${n(value)}</text>`;
      }).join("");
    }
    const heading = text({ ...node, frame: { x: f.x, y: f.y, width: f.width, height: f.height * .15 } }, { text: s.title, textBody: s.titleBody }, "chart");
    return `<g data-officekit-chart="scatter" data-officekit-scatter-style="${style}" data-officekit-x-min="${n(xs.low)}" data-officekit-x-max="${n(xs.high)}">${heading}${axes}<svg ${box(f)} viewBox="${n(f.x)} ${n(f.y)} ${n(f.width)} ${n(f.height)}" overflow="hidden">${output}</svg>${observed.length ? "" : '<text font-size="10">No observed data</text>'}</g>`;
  }
  function chart(node) {
    const s = node.native, f = node.frame;
    if ([SpreadsheetChartType.PIE, SpreadsheetChartType.DOUGHNUT].includes(s.type)) return circularChart(node);
    if (s.type === SpreadsheetChartType.BAR) return barChart(node);
    if (s.type === SpreadsheetChartType.SCATTER) return scatterChart(node);
    if (s.type !== SpreadsheetChartType.LINE) {
      limit(node, "chart.type", "preview.scene.paint.chart-type", s.type, "opaque");
      return placeholder(node, `chart type ${s.type}: not painted`);
    }
    const fail = (field, message) => {
      limit(node, field, "preview.scene.paint.chart-semantics", message, "unavailable");
      throw new TypeError(message);
    };
    unused(content.get("chart"), s, [...frameFields, "frameTransform", "type", "categories", "series", "xAxis", "yAxis", "lineOptions", "grouping", "displayBlanksAs", "showCategoryAxis", "showValueAxis", "title", "titleBody"], node, "chart.");
    if (s.comboSeries.length || s.secondaryXAxis || s.secondaryYAxis)
      fail("chart.comboSeries", "A simple line plot cannot replace mixed/secondary-axis topology.");
    if (s.lineOptions) {
      unused(SpreadsheetChartLineOptionsArtifactSchema, s.lineOptions, ["grouping", "smooth"], node, "chart.lineOptions.");
      if (s.lineOptions.grouping !== undefined && ![0, 1].includes(s.lineOptions.grouping))
        fail("chart.lineOptions.grouping", "Stacked lines require cumulative coordinates, not ordinary lines.");
      if (s.lineOptions.smooth === true) fail("chart.lineOptions.smooth", "Smooth line interpolation is not yet mapped.");
    }
    // OpenXmlChartSpaceCodec.NativeGrouping maps native "standard" to "none".
    // This is the writer IR vocabulary, not the raw ChartML token.
    if (s.grouping && s.grouping !== "none") fail("chart.grouping", "Nonstandard line grouping is not yet mapped.");
    const blank = s.displayBlanksAs ?? "gap";
    if (!["gap", "zero", "span"].includes(blank)) fail("chart.displayBlanksAs", "Unknown native blank-display strategy.");
    const series = categorySeries(node, s, ["name", "values", "missingValueIndexes", "line", "marker"]);
    if (series.some(({ missing }) => missing.size) && blank !== "gap") fail("chart.displayBlanksAs", `Explicit ${blank} would transform missing observations; this painter has not implemented separately labeled display-policy evidence.`);
    const xa = s.xAxis, ya = s.yAxis;
    if (xa) {
      unused(SpreadsheetChartAxisArtifactSchema, xa, ["reverse", "visible", "axisLineVisible", "tickLabelsVisible", "tickLabelInterval"], node, "chart.xAxis.");
      if ([xa.minimum, xa.maximum, xa.logBase].some(v => v !== undefined)) fail("chart.xAxis", "Numeric category-axis bounds cannot be treated as evenly spaced labels.");
    }
    if (ya) unused(SpreadsheetChartAxisArtifactSchema, ya, ["minimum", "maximum", "reverse", "logBase", "visible", "axisLineVisible", "tickLabelsVisible"], node, "chart.yAxis.");
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
      if (marker) unused(SpreadsheetChartMarkerArtifactSchema, marker, ["symbol", "size", "fill", "fillOpacityThousandthPercent", "line"], node, `${prefix}.marker.`);
      const symbol = marker?.symbol ?? 1, radius = (marker?.size ?? 5) / 2;
      if (!Number.isFinite(radius) || radius <= 0) fail(`${prefix}.marker.size`, "Invalid marker size.");
      if (marker?.fill) unused(SpreadsheetColorSchema, marker.fill, ["rgb"], node, `${prefix}.marker.fill.`);
      const markerFill = marker?.fill?.source.case === "rgb" ? rgb(marker.fill.source.value) : color;
      let markerOutline = 'stroke="none"';
      if (marker?.line) {
        const ml = marker.line, field = `${prefix}.marker.line`;
        unused(SpreadsheetChartLineStyleArtifactSchema, ml, ["color", "widthPoints", "opacityThousandthPercent", "dashStyle", "cap", "join"], node, `${field}.`);
        if (ml.color) unused(SpreadsheetColorSchema, ml.color, ["rgb"], node, `${field}.color.`);
        if (ml.color?.source.case !== "rgb" || ml.widthPoints === undefined)
          limit(node, field, "preview.scene.paint.chart-inherited-paint", "Marker outline uses review defaults where direct native paint is absent.");
        markerOutline = linePaint(node, field, ml.color?.source.case === "rgb" ? rgb(ml.color.source.value) : color,
          numeric(ml.widthPoints ?? 1), sceneOpacity(ml.opacityThousandthPercent ?? 100000),
          ["solid", "solid", "dashed", "dotted", "dash-dot", "dash-dot-dot"][ml.dashStyle], ml.cap, ml.join);
      }
      if (symbol !== 1 && ![2, 3, 4, 5, 6].includes(symbol))
        limit(node, `${prefix}.marker.symbol`, "preview.scene.paint.chart-marker", symbol);
      const paths = [], marks = [];
      for (const points of segments) {
        const path = points.length > 1 ? `<path data-officekit-line-segment="${points[0].index}:${points.at(-1).index}" d="${points.map((p, i) => `${i ? "L" : "M"} ${n(p.x)} ${n(p.y)}`).join(" ")}" fill="none" ${paint}/>` : "";
        paths.push(path);
        const dots = points.map(p => {
          const outside = p.x < plot.x || p.x > plot.x + plot.width || p.y < plot.y || p.y > plot.y + plot.height;
          let mark = "";
          if ([2, 3].includes(symbol)) mark = `<circle cx="${n(p.x)}" cy="${n(p.y)}" r="${n(radius)}"/>`;
          if (symbol === 4) mark = `<rect ${box({ x: p.x - radius, y: p.y - radius, width: radius * 2, height: radius * 2 })}/>`;
          if (symbol === 5) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y)} L ${n(p.x)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y)} Z"/>`;
          if (symbol === 6) mark = `<path d="M ${n(p.x)} ${n(p.y - radius)} L ${n(p.x + radius)} ${n(p.y + radius)} L ${n(p.x - radius)} ${n(p.y + radius)} Z"/>`;
          if (!mark && points.length === 1 && !outside) {
            limit(node, `${prefix}.values[${p.index}]`, "preview.scene.paint.chart-isolated-review-point", "Isolated observation shown as a hollow review dot, not an authored marker.");
            mark = `<circle data-officekit-review-point="isolated" cx="${n(p.x)}" cy="${n(p.y)}" r="3" fill="none" stroke="${color}" stroke-dasharray="1 1"/>`;
          }
          // A point whose centre is outside explicit axis bounds stays in the
          // evidence, but must not be projected onto the edge as a false value.
          // In-range markers can extend half their size beyond the plot edge.
          return `<g data-officekit-point="${p.index}" data-officekit-value="${n(p.value)}"${outside ? ' data-officekit-point-outside-plot="true"' : ""} fill="${markerFill}" fill-opacity="${n(sceneOpacity(marker?.fillOpacityThousandthPercent ?? 100000))}" ${markerOutline}><title>${esc(entry.name)} / ${esc(s.categories[p.index])}: ${n(p.value)}</title>${outside ? "" : mark}</g>`;
        }).join("");
        marks.push(dots);
      }
      return `<g data-officekit-series="${si}" data-officekit-series-name="${esc(entry.name)}">${missingEvidence.join("")}<svg data-officekit-line-clip="plot" ${box(plot)} viewBox="${n(plot.x)} ${n(plot.y)} ${n(plot.width)} ${n(plot.height)}" overflow="hidden">${paths.join("")}</svg><g data-officekit-markers="unclipped-at-plot-edge">${marks.join("")}</g></g>`;
    }).join("");
    // Per-series nested SVG clips lines at the plot; the outer chart viewport
    // bounds markers at the chart frame, without cutting edge observations in
    // half. Local viewports avoid colliding IDs across pages or nested groups.
    let axes = "";
    const bottom = plot.y + plot.height, right = plot.x + plot.width;
    if ((xa?.visible ?? s.showCategoryAxis ?? true) === true) {
      if (xa?.axisLineVisible !== false) axes += `<line data-officekit-axis="category" x1="${n(plot.x)}" y1="${n(bottom)}" x2="${n(right)}" y2="${n(bottom)}" stroke="#64748B"/>`;
      const interval = xa?.tickLabelInterval ?? 1;
      if (!Number.isSafeInteger(interval) || interval < 1) fail("chart.xAxis.tickLabelInterval", "Invalid category tick interval.");
      if (xa?.tickLabelsVisible !== false && xa?.tickLabelPosition !== "none") axes += s.categories.map((label, i) => {
        if (i % interval) return "";
        let ratio = s.categories.length <= 1 ? .5 : i / (s.categories.length - 1);
        if (xa?.reverse === true) ratio = 1 - ratio;
        return `<text data-officekit-category="${i}" x="${n(plot.x + ratio * plot.width)}" y="${n(bottom + 12)}" text-anchor="middle" font-size="10" fill="#334155">${esc(label)}</text>`;
      }).join("");
    }
    if ((ya?.visible ?? s.showValueAxis ?? true) === true) {
      if (ya?.axisLineVisible !== false) axes += `<line data-officekit-axis="value" x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(bottom)}" stroke="#64748B"/>`;
      if (ya?.tickLabelsVisible !== false && ya?.tickLabelPosition !== "none") axes += [0, .5, 1].map(ratio => {
        const scaled = low + ratio * (high - low), value = logBase === undefined ? scaled : Math.pow(logBase, scaled);
        const y = plot.y + (ya?.reverse === true ? ratio : 1 - ratio) * plot.height;
        return `<text data-officekit-value-tick="${n(value)}" x="${n(plot.x - 4)}" y="${n(y + 3)}" text-anchor="end" font-size="10" fill="#334155">${n(value)}</text>`;
      }).join("");
    }
    const heading = text({ ...node, frame: { x: plot.x, y: f.y, width: plot.width, height: f.height * .15 } }, { text: s.title, textBody: s.titleBody }, "chart");
    return `<g data-officekit-chart="line" data-officekit-blank-policy="${blank}" data-officekit-scale-min="${n(low)}" data-officekit-scale-max="${n(high)}" data-officekit-log-base="${logBase ?? "linear"}">${heading}${axes}<svg ${box(f)} viewBox="${n(f.x)} ${n(f.y)} ${n(f.width)} ${n(f.height)}" overflow="hidden">${output}</svg>${observed.length ? "" : `<text x="${n(plot.x)}" y="${n(plot.y + 12)}" font-size="10">No observed data</text>`}</g>`;
  }
  function paintGroup(node, group, prefix) {
    const f = group.frame, c = group.childFrame;
    unused(content.get("group"), group.native, [...frameFields, "childLeftEmu", "childTopEmu", "childWidthEmu", "childHeightEmu", "frameTransform", "children"], node, `${prefix}.`);
    if (!f || !c || c.width <= 0 || c.height <= 0) throw new RangeError("Nonpositive native group child extent");
    for (const value of Object.values(f)) numeric(value);
    if (f.width < 0 || f.height < 0) throw new RangeError("Negative native group frame");
    const matrix = `translate(${n(f.x)} ${n(f.y)}) scale(${n(f.width / c.width)} ${n(f.height / c.height)}) translate(${n(-c.x)} ${n(-c.y)})`;
    return `<g transform="${matrix}">${group.children.map(draw).join("")}</g>`;
  }
  function draw(node) {
    visitedNodes.add(node.scenePath);
    const identity = `data-officekit-native-id="${esc(node.nativeId)}" data-officekit-scene-path="${esc(node.scenePath)}"${node.semanticId ? ` data-officekit-id="${esc(node.semanticId)}"` : ""}`;
    let svg;
    try {
      if (node.hidden === true) {
        hiddenScenePaths.add(node.scenePath);
        return `<g ${identity} display="none"/>`;
      }
      if (!node.frame) throw new TypeError("Missing native frame");
      for (const value of Object.values(node.frame)) numeric(value);
      if (node.frame.width < 0 || node.frame.height < 0) throw new RangeError("Negative native frame");
      if (node.kind === "group") {
        svg = paintGroup(node, node, "group");
      } else if (node.kind === "diagram") {
        unused(content.get("diagram"), node.native, [...frameFields, "drawing", "drawingCacheVerified"], node, "diagram.");
        if (!node.native.drawingCacheVerified || !node.drawing?.children.length) {
          limit(node, "diagram.drawing", "preview.scene.paint.diagram-cache", "No verified cached drawing is available", "opaque");
          svg = placeholder(node, "diagram: verified drawing unavailable");
        } else if (view.scene.origin === 2) {
          // ReadCachedDrawing currently proves semantic-node geometry only;
          // it does not import the complete cached pens/fills/connections.
          limit(node, "diagram.drawing", "preview.scene.paint.diagram-import-incomplete", "Candidate-import cache lacks complete native paint and connection state", "unavailable");
          svg = placeholder(node, "diagram: imported drawing incomplete");
        } else {
          limit(node, "diagram.drawing", "preview.scene.paint.diagram-cache-scope", "Displaying verified native cache; semantic layout/edit fidelity is not independently established by this image.");
          svg = `<g data-officekit-diagram="verified-cache" transform="${frameTransform(node.drawing.frame, node.drawing.transform)}">${paintGroup(node, node.drawing, "diagram.drawing")}</g>`;
        }
      } else if (node.kind === "shape") svg = shape(node);
      else if (node.kind === "image") svg = image(node);
      else if (node.kind === "connector") svg = connector(node);
      else if (node.kind === "table") svg = table(node);
      else if (node.kind === "chart") svg = chart(node);
      else { limit(node, node.kind, "preview.scene.paint.content", node.kind, "opaque"); svg = placeholder(node, `${node.kind}: not painted`); }
      const transform = frameTransform(node.frame, node.transform);
      transformedScenePaths.add(node.scenePath);
      return `<g ${identity} transform="${transform}">${svg}</g>`;
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
  // Complete actual drawing before the input profile can inspect successful
  // hidden/transform branches, including owners lowered across multiple pages.
  const bodies = view.pages.map(page => page.nodes.map(draw).join(""));
  const paintIdentity = { renderer: "officekit-native-scene-svg-internal", scene: view.scene,
    sceneEvidence: ppjPreviewSceneIdentity(view.scene),
    hiddenScenePaths: Object.freeze([...hiddenScenePaths]),
    transformedScenePaths: Object.freeze([...transformedScenePaths]) };
  let inputAssessment;
  if (assessInput) {
    const program = JSON.parse(new TextDecoder().decode(receipt.programJson));
    // Public asset IDs need not match native IDs. Resolve the canonical
    // declaration through already-verified MIME/hash scene payloads, never IO.
    const assetKey = (mime, sha) => `${mime}:${String(sha).toLowerCase()}`;
    const payloads = new Map(view.assets.map(asset => [assetKey(asset.contentType, asset.sha256), asset.data]));
    const inputAssets = new Map((program.assets ?? []).map(asset => [asset.id, {
      ...asset, data: payloads.get(assetKey(asset.mimeType, asset.sha256)),
    }]));
    inputAssessment = assessPpjPreviewInput(program, {
      rendererProfile: "native-scene-svg", sceneReceipt: receipt, scenePaint: paintIdentity,
      assets: inputAssets,
    });
    // A character trie finds the deepest semantic owner in O(path length),
    // without scanning every scene binding for every input diagnostic.
    const root = { next: new Map() };
    for (const binding of view.scene.bindings) {
      let cursor = root;
      for (const char of binding.programPath) {
        if (!cursor.next.has(char)) cursor.next.set(char, { next: new Map() });
        cursor = cursor.next.get(char);
      }
      (cursor.bindings ??= []).push(binding);
    }
    const pagePaths = new Map(view.pages.filter(page => page.pageId).map(page => [page.pageId, page.scenePath]));
    for (const diagnostic of inputAssessment.diagnostics) {
      let cursor = root, owners;
      for (let i = 0; i < diagnostic.path.length; i++) {
        cursor = cursor.next.get(diagnostic.path[i]);
        if (!cursor) break;
        if (cursor.bindings && (i + 1 === diagnostic.path.length || ".[".includes(diagnostic.path[i + 1]))) {
          const matching = cursor.bindings.filter(binding => binding.pageId === diagnostic.pageId);
          if (matching.length) owners = matching;
        }
      }
      if (owners) for (const owner of owners) diagnostics.push(Object.freeze({ ...diagnostic, scenePath: owner.scenePath }));
      else diagnostics.push(Object.freeze({ ...diagnostic,
        scenePath: pagePaths.get(diagnostic.pageId) ?? "$.presentation" }));
    }
  }
  // Native fields outside any page affect every page's review evidence.
  // Keep program ownership separate from the native scene address.
  const globalDiagnostics = diagnostics.filter(d => !view.pages.some(page =>
    d.scenePath === page.scenePath || d.scenePath.startsWith(`${page.scenePath}.`)));
  const pages = view.pages.map((page, index) => {
    const pageNode = { ...page, path: "$" };
    limit(pageNode, "", "preview.scene.paint.integration-pending", "Internal scene painter: production G-11/G-12 integration and complete field coverage remain pending.");
    unused(PresentationSlideSchema, page.native, ["id", "elements", "background", "hidden"], pageNode);
    const background = page.native.background;
    if (background) unused(PresentationBackgroundSchema, background, ["colorRgb", "opacityThousandthPercent", "solid"], pageNode, "background.");
    const fill = background?.color.case === "colorRgb" ? rgb(background.color.value, "#FFFFFF") : "#FFFFFF";
    const body = bodies[index];
    const pageDiagnostics = [...globalDiagnostics, ...diagnostics.filter(d => d.scenePath === page.scenePath || d.scenePath.startsWith(`${page.scenePath}.`))];
    const ownDiagnostics = new Map();
    const childrenOf = node => node.children.length ? node.children : node.drawing?.children ?? [];
    function indexNode(node) {
      ownDiagnostics.set(node.scenePath, []);
      childrenOf(node).forEach(indexNode);
    }
    page.nodes.forEach(indexNode);
    // Resolve each field to the closest actual node by structural address, not
    // semantic ID (generated siblings may share one owner). No prefix scan of
    // every node for every diagnostic is needed.
    for (const diagnostic of pageDiagnostics) {
      let address = diagnostic.scenePath;
      while (address && !ownDiagnostics.has(address)) {
        const dot = address.lastIndexOf(".");
        address = dot < 0 ? undefined : address.slice(0, dot);
      }
      if (address) ownDiagnostics.get(address).push(diagnostic);
    }
    function assessNode(node) {
      return previewAssessment({ path: node.path, scenePath: node.scenePath,
        pageId: node.pageId, id: node.semanticId, assessed: visitedNodes.has(node.scenePath),
        diagnostics: ownDiagnostics.get(node.scenePath), children: childrenOf(node).map(assessNode) });
    }
    const assessment = previewAssessment({ path: "$", scenePath: page.scenePath, pageId: page.pageId,
      assessed: true, diagnostics: pageDiagnostics, children: page.nodes.map(assessNode) });
    const { status, reliability } = assessment;
    const bannerHeight = Math.min(24, view.canvas.height * .12);
    const banner = `<g data-officekit-review="${reliability.status}"><rect width="${n(view.canvas.width)}" height="${n(bannerHeight)}" fill="${reliability.status === "failed" ? "#991B1B" : "#92400E"}"/><text x="2" y="${n(bannerHeight * .7)}" font-size="${n(bannerHeight * .5)}" fill="#FFFFFF">INTERNAL SCENE PREVIEW · ${esc(reliability.status)}</text></g>`;
    return Object.freeze({ id: page.pageId ?? page.nativeId, pageId: page.pageId, nativeId: page.nativeId, hidden: page.hidden,
      assessment, status, reliability, diagnostics: assessment.diagnostics,
      svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${n(view.canvas.width)}" height="${n(view.canvas.height)}" viewBox="0 0 ${n(view.canvas.width)} ${n(view.canvas.height)}"><rect width="100%" height="100%" fill="${fill}" fill-opacity="${n(sceneOpacity(background?.opacityThousandthPercent ?? 100000))}"/>${body}${banner}</svg>` });
  });
  const assessment = previewAssessment({ path: "$", scenePath: "$.presentation", assessed: true, diagnostics: globalDiagnostics,
    children: pages.map(page => page.assessment) });
  return Object.freeze({ ...paintIdentity, canvas: view.canvas,
    ...(inputAssessment ? { inputAssessment } : {}),
    pages: Object.freeze(pages), assessment, diagnostics: assessment.diagnostics, status: assessment.status, reliability: assessment.reliability });
}
