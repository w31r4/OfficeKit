import { previewDiagnostic } from "./preview-diagnostics.mjs";

// These conditions describe the current SVG drawing branches. Removing a
// limitation requires a drawing regression, not a weaker diagnostic severity.
export function previewFactualErrors(element, { path, pageId, id }, support,
  { hiddenOwnerPaths = new Set(), resolvedTransformPaths = new Set(), resolvedGroupCoordinates = false, resolvedConnectorPaths = new Set(), resolvedIsolatedLinePaths = new Set(), resolvedLineSeriesPaths = new Set(), resolvedDataset = false, resolvedShapeGeometry = false } = {}) {
  const diagnostics = [];
  const add = (name, suffix, value) => {
    const rule = support.factual[name];
    if (!rule) throw new Error(`Missing registry preview factual rule: ${name}`);
    diagnostics.push(previewDiagnostic({ path: `${path}${suffix}`, pageId, id, value,
      status: rule.status, reason: rule.reason, severity: rule.severity,
      action: "Do not accept this preview as correctness evidence; repair the drawing for this input field and rerun its regression.",
    }));
  };
  if (!element || typeof element !== "object") return diagnostics;
  if (element.type === "shape") {
    const geometry = element.geometry;
    if (!resolvedShapeGeometry && (element.text || geometry && (geometry.kind !== "preset" || geometry.preset !== "rect"))) add("shapeGeometry", ".geometry", geometry);
  }
  if (element.type === "connector") {
    for (const key of ["from", "to"]) if (element[key] !== undefined && !resolvedConnectorPaths.has(`${path}.${key}`)) add("connector", `.${key}`, element[key]);
  }
  if (element.hidden === true && !hiddenOwnerPaths.has(path)) add("visibility", ".hidden", true);
  if (Number.isFinite(element.frame?.rotation) && element.frame.rotation % 360 !== 0
      && !resolvedTransformPaths.has(`${path}.frame.rotation`)) add("transform", ".frame.rotation", element.frame.rotation);
  for (const key of ["flipH", "flipV"]) if (element.frame?.[key] === true
      && !resolvedTransformPaths.has(`${path}.frame.${key}`)) add("transform", `.frame.${key}`, true);
  if (!resolvedGroupCoordinates && element.type === "group" && element.childFrame && ["x", "y", "width", "height"].some((key) =>
    Number.isFinite(element.frame?.[key]) && Number.isFinite(element.childFrame[key]) && element.frame[key] !== element.childFrame[key])) add("groupCoordinates", ".childFrame", element.childFrame);
  if (element.type !== "chart") return diagnostics;

  const type = element.chartType, series = element.data?.series || [];
  const drawnMaximum = series.reduce((max, item) => (item.values || []).reduce((current, value) => Number.isFinite(value) ? Math.max(current, value) : current, max), 1);
  const cartesian = ["bar", "column", "line", "area", "combo"].includes(type);
  if (element.data?.dataset !== undefined && !resolvedDataset) add("chartChannels", ".data.dataset", element.data.dataset);
  if (element.style?.stacking !== undefined && element.style.stacking !== "none") add("chartStacking", ".style.stacking", element.style.stacking);
  if (element.style?.symbol !== undefined) add("chartSymbol", ".style.symbol", element.style.symbol);
  if (type === "bar" && series.some((s) => s.values?.some(Number.isFinite))) add("chartChannels", ".chartType", type);
  if (["doughnut", "radar", "bubble", "candlestick", "sankey"].includes(type) && series.some((s) => s.values?.some(Number.isFinite))) add("chartChannels", ".data", element.data);
  if (type === "pie" && series.flatMap((s) => s.values || []).filter((v) => Number.isFinite(v) && v > 0).length > 1) add("chartProportion", ".data", element.data);
  if (series.filter((s) => ["bar", "column"].includes(s.chartType)).length > 1) add("chartStacking", ".data.series", series);
  for (const key of ["xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis", "spokeAxis"]) {
    const value = element[key];
    if (value === undefined) continue;
    const axes = Array.isArray(value) ? value : [value];
    axes.forEach((axis, index) => {
      const owner = `.${key}${Array.isArray(value) ? `[${index}]` : ""}`;
      for (const field of ["min", "max", "logBase", "majorUnit", "minorUnit"]) {
        // Unresolved tokens are limitations, not proof of contradictory values.
        // The old renderer uses 0..maximum for Y and category-index X.
        const expectedMax = key.includes("XAxis") || key === "xAxis" ? Math.max(1, (element.data?.categories?.length || 0) - 1) : drawnMaximum;
        if (Number.isFinite(axis?.[field]) && !(field === "min" && axis[field] === 0)
          && !(field === "max" && axis[field] === expectedMax)) add("chartScale", `${owner}.${field}`, axis[field]);
      }
      if (axis?.reverse === true) add("chartScale", `${owner}.reverse`, axis.reverse);
      if (axis?.visible === false) add("chartScale", `${owner}.visible`, false);
    });
  }
  series.forEach((s, index) => {
    const at = `.data.series[${index}]`, values = s.values || [];
    if (cartesian && s.chartType === undefined && values.some(Number.isFinite) && !resolvedLineSeriesPaths.has(`${path}${at}`)) add("chartSeriesType", `${at}.chartType`, undefined);
    for (const key of ["xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "sources", "targets"]) {
      if (s[key] !== undefined) add("chartChannels", `${at}.${key}`, s[key]);
    }
    if (s.parents?.some((parent) => parent !== null) || s.levels > 1) add("chartHierarchy", `${at}.${s.parents ? "parents" : "levels"}`, s.parents ?? s.levels);
    if (s.symbol !== undefined) add("chartSymbol", `${at}.symbol`, s.symbol);
    for (const key of ["axis", "xAxisIndex", "yAxisIndex"]) {
      if (s[key] === "secondary" || s[key] === 1) add("chartScale", `${at}.${key}`, s[key]);
    }
    if (s.chartType === "area" && values.some(Number.isFinite)) add("chartStacking", `${at}.chartType`, s.chartType);
    values.forEach((value, point) => {
      const coerced = ["treemap", "sunburst", "waterfall"].includes(type) || element.style?.stacking === "stream" || s.symbol
        || ["bar", "column", "scatter"].includes(s.chartType);
      if (!Number.isFinite(value) && coerced) add("chartMissing", `${at}.values[${point}]`, value);
      if (Number.isFinite(value) && ["line", "area"].includes(s.chartType)
        && !Number.isFinite(values[point - 1]) && !Number.isFinite(values[point + 1])
        && !resolvedIsolatedLinePaths.has(`${path}${at}.values[${point}]`)) add("chartMissing", `${at}.values[${point}]`, value);
      if (cartesian && Number.isFinite(value) && value < 0) add("chartScale", `${at}.values[${point}]`, value);
    });
    if (type === "waterfall") {
      s.pointRoles?.forEach((role, point) => { if (role === "total") add("chartTotals", `${at}.pointRoles[${point}]`, role); });
      let running = 0;
      if (values.some((value) => { running += Number.isFinite(value) ? value : 0; return running < 0 || running > drawnMaximum; })) add("chartScale", `${at}.values`, values);
    }
  });
  return diagnostics;
}
