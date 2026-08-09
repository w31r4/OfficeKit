import {
  PresentationChartAxisGroup,
  SpreadsheetChartDataLabelPosition,
  SpreadsheetChartErrorBarDirection,
  SpreadsheetChartErrorBarType,
  SpreadsheetChartErrorBarValueType,
  SpreadsheetChartLineDashStyle,
  SpreadsheetChartMarkerSymbol,
  SpreadsheetChartTrendlineType,
  SpreadsheetChartType,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import { normalizePresentationChartErrorBars, normalizePresentationChartTrendlines } from "../presentation/ooxml-charts.mjs";

const MAX_PRESENTATION_CHART_POINTS = 1_048_576;
const PRESENTATION_CHART_TYPES_TO_WIRE = new Map([
  ["bar", SpreadsheetChartType.BAR],
  ["line", SpreadsheetChartType.LINE],
  ["pie", SpreadsheetChartType.PIE],
  ["area", SpreadsheetChartType.AREA],
  ["doughnut", SpreadsheetChartType.DOUGHNUT],
  ["scatter", SpreadsheetChartType.SCATTER],
  ["bubble", SpreadsheetChartType.BUBBLE],
  ["combo", SpreadsheetChartType.COMBO],
]);
const PRESENTATION_CHART_TYPES_FROM_WIRE = new Map([...PRESENTATION_CHART_TYPES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_NUMERIC_X_CHART_TYPES = new Set([SpreadsheetChartType.SCATTER, SpreadsheetChartType.BUBBLE]);
const PRESENTATION_CIRCULAR_CHART_TYPES = new Set([SpreadsheetChartType.PIE, SpreadsheetChartType.DOUGHNUT]);
const PRESENTATION_CHART_AXIS_GROUPS_TO_WIRE = new Map([
  ["primary", PresentationChartAxisGroup.PRIMARY],
  ["secondary", PresentationChartAxisGroup.SECONDARY],
]);
const PRESENTATION_CHART_AXIS_GROUPS_FROM_WIRE = new Map([...PRESENTATION_CHART_AXIS_GROUPS_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_LINE_STYLES_TO_WIRE = new Map([
  ["solid", SpreadsheetChartLineDashStyle.SOLID],
  ["dashed", SpreadsheetChartLineDashStyle.DASHED],
  ["dotted", SpreadsheetChartLineDashStyle.DOTTED],
  ["dash-dot", SpreadsheetChartLineDashStyle.DASH_DOT],
  ["dash-dot-dot", SpreadsheetChartLineDashStyle.DASH_DOT_DOT],
]);
const PRESENTATION_CHART_LINE_STYLES_FROM_WIRE = new Map([...PRESENTATION_CHART_LINE_STYLES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE = new Map([
  ["bestFit", SpreadsheetChartDataLabelPosition.BEST_FIT],
  ["bottom", SpreadsheetChartDataLabelPosition.BOTTOM],
  ["center", SpreadsheetChartDataLabelPosition.CENTER],
  ["insideBase", SpreadsheetChartDataLabelPosition.INSIDE_BASE],
  ["insideEnd", SpreadsheetChartDataLabelPosition.INSIDE_END],
  ["left", SpreadsheetChartDataLabelPosition.LEFT],
  ["outsideEnd", SpreadsheetChartDataLabelPosition.OUTSIDE_END],
  ["right", SpreadsheetChartDataLabelPosition.RIGHT],
  ["top", SpreadsheetChartDataLabelPosition.TOP],
]);
const PRESENTATION_CHART_LABEL_POSITIONS_FROM_WIRE = new Map([...PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_MARKERS_TO_WIRE = new Map([
  ["none", SpreadsheetChartMarkerSymbol.NONE],
  ["dot", SpreadsheetChartMarkerSymbol.DOT],
  ["circle", SpreadsheetChartMarkerSymbol.CIRCLE],
  ["square", SpreadsheetChartMarkerSymbol.SQUARE],
  ["diamond", SpreadsheetChartMarkerSymbol.DIAMOND],
  ["triangle", SpreadsheetChartMarkerSymbol.TRIANGLE],
  ["x", SpreadsheetChartMarkerSymbol.X],
  ["star", SpreadsheetChartMarkerSymbol.STAR],
  ["plus", SpreadsheetChartMarkerSymbol.PLUS],
  ["dash", SpreadsheetChartMarkerSymbol.DASH],
]);
const PRESENTATION_CHART_MARKERS_FROM_WIRE = new Map([...PRESENTATION_CHART_MARKERS_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_TRENDLINES_TO_WIRE = new Map([
  ["exp", SpreadsheetChartTrendlineType.EXPONENTIAL],
  ["linear", SpreadsheetChartTrendlineType.LINEAR],
  ["log", SpreadsheetChartTrendlineType.LOGARITHMIC],
  ["movingAvg", SpreadsheetChartTrendlineType.MOVING_AVERAGE],
  ["poly", SpreadsheetChartTrendlineType.POLYNOMIAL],
  ["power", SpreadsheetChartTrendlineType.POWER],
]);
const PRESENTATION_CHART_TRENDLINES_FROM_WIRE = new Map([...PRESENTATION_CHART_TRENDLINES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_ERROR_BAR_DIRECTIONS_TO_WIRE = new Map([
  ["x", SpreadsheetChartErrorBarDirection.X],
  ["y", SpreadsheetChartErrorBarDirection.Y],
]);
const PRESENTATION_CHART_ERROR_BAR_DIRECTIONS_FROM_WIRE = new Map([...PRESENTATION_CHART_ERROR_BAR_DIRECTIONS_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_ERROR_BAR_TYPES_TO_WIRE = new Map([
  ["both", SpreadsheetChartErrorBarType.BOTH],
  ["minus", SpreadsheetChartErrorBarType.MINUS],
  ["plus", SpreadsheetChartErrorBarType.PLUS],
]);
const PRESENTATION_CHART_ERROR_BAR_TYPES_FROM_WIRE = new Map([...PRESENTATION_CHART_ERROR_BAR_TYPES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_ERROR_BAR_VALUE_TYPES_TO_WIRE = new Map([
  ["cust", SpreadsheetChartErrorBarValueType.CUSTOM],
  ["fixedVal", SpreadsheetChartErrorBarValueType.FIXED_VALUE],
  ["percentage", SpreadsheetChartErrorBarValueType.PERCENTAGE],
  ["stdDev", SpreadsheetChartErrorBarValueType.STANDARD_DEVIATION],
  ["stdErr", SpreadsheetChartErrorBarValueType.STANDARD_ERROR],
]);
const PRESENTATION_CHART_ERROR_BAR_VALUE_TYPES_FROM_WIRE = new Map([...PRESENTATION_CHART_ERROR_BAR_VALUE_TYPES_TO_WIRE].map(([name, value]) => [value, name]));

function chartColor(value, chart, field, rgb) {
  if (value == null || value === "") return undefined;
  const normalized = rgb(value, `${chart.id}.${field}`);
  if (!normalized) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} cannot be transparent.`, [], { code: "unsupported_presentation_features" });
  return { source: { case: "rgb", value: normalized } };
}

function chartLine(line, chart, field, rgb) {
  if (!line) return undefined;
  const styleName = String(line.style || "solid");
  const aliases = { dot: "dotted", dash: "dashed", dashDot: "dash-dot", longDashDotDot: "dash-dot-dot" };
  const dashStyle = PRESENTATION_CHART_LINE_STYLES_TO_WIRE.get(aliases[styleName] || styleName);
  if (dashStyle == null) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} uses unsupported line style ${styleName}.`, [], { code: "unsupported_presentation_features" });
  const width = Number(line.width ?? line.weight ?? 1);
  if (!Number.isFinite(width) || width < 0 || width > 1584) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} has an invalid line width.`, [], { code: "invalid_presentation_chart" });
  return {
    ...(line.fill || line.color ? { color: chartColor(line.fill || line.color, chart, `${field}.color`, rgb) } : {}),
    dashStyle,
    widthPoints: width,
  };
}

function chartMarker(marker, chart, field, rgb) {
  if (!marker) return undefined;
  const symbol = PRESENTATION_CHART_MARKERS_TO_WIRE.get(String(marker.symbol || "circle"));
  if (symbol == null) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} uses unsupported marker ${marker.symbol}.`, [], { code: "unsupported_presentation_features" });
  const size = marker.size == null ? undefined : Number(marker.size);
  if (size !== undefined && (!Number.isInteger(size) || size < 2 || size > 72)) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} has an invalid marker size.`, [], { code: "invalid_presentation_chart" });
  return {
    symbol,
    ...(size === undefined ? {} : { size }),
    ...(marker.fill ? { fill: chartColor(marker.fill, chart, `${field}.fill`, rgb) } : {}),
    ...(marker.line || marker.stroke ? { line: chartLine(marker.line || marker.stroke, chart, `${field}.line`, rgb) } : {}),
  };
}

function chartTrendlines(value, valueCount, effectiveType, chart, field, rgb) {
  let trendlines;
  try { trendlines = normalizePresentationChartTrendlines(value, valueCount, effectiveType); }
  catch (error) {
    throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} ${String(error?.message || error)}`, [], { code: "invalid_presentation_chart" });
  }
  return trendlines.map((trendline) => ({
    type: PRESENTATION_CHART_TRENDLINES_TO_WIRE.get(trendline.type),
    name: trendline.name || "",
    polynomialOrder: trendline.order,
    period: trendline.period,
    forward: trendline.forward,
    backward: trendline.backward,
    intercept: trendline.intercept,
    displayEquation: trendline.displayEquation,
    displayRSquared: trendline.displayRSquared,
    ...(trendline.line ? { line: chartLine(trendline.line, chart, `${field}.line`, rgb) } : {}),
  }));
}

function chartErrorBarData(errorBars, side) {
  const values = errorBars[`${side}Values`];
  const formatCode = errorBars[`${side}FormatCode`];
  if (values == null && formatCode == null) return undefined;
  return { values: values || [], formula: "", formatCode: formatCode || "" };
}

function chartErrorBars(value, valueCount, effectiveType, chart, field, rgb) {
  let errorBars;
  try { errorBars = normalizePresentationChartErrorBars(value, effectiveType, valueCount); }
  catch (error) {
    throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field} ${String(error?.message || error)}`, [], { code: "invalid_presentation_chart" });
  }
  if (!errorBars) return undefined;
  return {
    direction: PRESENTATION_CHART_ERROR_BAR_DIRECTIONS_TO_WIRE.get(errorBars.direction),
    type: PRESENTATION_CHART_ERROR_BAR_TYPES_TO_WIRE.get(errorBars.type),
    valueType: PRESENTATION_CHART_ERROR_BAR_VALUE_TYPES_TO_WIRE.get(errorBars.valueType),
    value: errorBars.value,
    noEndCap: errorBars.noEndCap,
    plus: chartErrorBarData(errorBars, "plus"),
    minus: chartErrorBarData(errorBars, "minus"),
    ...(errorBars.line ? { line: chartLine(errorBars.line, chart, `${field}.line`, rgb) } : {}),
  };
}

function chartAxis(axis, chart, field, original) {
  const title = typeof axis?.title === "object" ? axis.title.text : axis?.title;
  const result = {
    title: String(title || ""),
    numberFormatCode: String(axis?.numberFormatCode || axis?.numberFormat || ""),
    ...(axis?.tickLabelInterval == null ? {} : { tickLabelInterval: Number(axis.tickLabelInterval) }),
    ...(axis?.min == null ? {} : { minimum: Number(axis.min) }),
    ...(axis?.max == null ? {} : { maximum: Number(axis.max) }),
    ...(axis?.majorUnit == null ? {} : { majorUnit: Number(axis.majorUnit) }),
  };
  const hasSemantics = result.title || result.numberFormatCode || result.tickLabelInterval !== undefined || result.minimum !== undefined || result.maximum !== undefined || result.majorUnit !== undefined;
  if (!hasSemantics && !original) return undefined;
  for (const [name, value] of Object.entries(result)) {
    if (typeof value === "number" && !Number.isFinite(value)) throw new OfficeKitCodecError(`Presentation chart ${chart.id} ${field}.${name} must be finite.`, [], { code: "invalid_presentation_chart" });
  }
  return result;
}

function chartDataLabels(labels, chart, original) {
  const source = labels || {};
  const positionName = ({ b: "bottom", ctr: "center", inBase: "insideBase", inEnd: "insideEnd", l: "left", outEnd: "outsideEnd", r: "right", t: "top" })[source.position] || source.position || "bestFit";
  const position = PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE.get(positionName);
  if (position == null) throw new OfficeKitCodecError(`Presentation chart ${chart.id} uses unsupported data-label position ${source.position}.`, [], { code: "unsupported_presentation_features" });
  const hasSemantics = Boolean(source.showValue || source.showCategoryName || source.showSeriesName || source.showPercent) || (source.position && source.position !== "bestFit");
  if (!hasSemantics && !original) return undefined;
  return {
    showValue: Boolean(source.showValue),
    showCategoryName: Boolean(source.showCategoryName),
    ...(source.showSeriesName == null ? {} : { showSeriesName: Boolean(source.showSeriesName) }),
    ...(source.showPercent == null ? {} : { showPercent: Boolean(source.showPercent) }),
    ...(source.position == null && !original?.position ? {} : { position }),
  };
}

export function presentationChartToWire(chart, original, { emuFromPixels, rgb, sourceBoundFrameEmuFromPixels }) {
  const originalChart = original?.content?.case === "chart" ? original.content.value : undefined;
  const type = PRESENTATION_CHART_TYPES_TO_WIRE.get(String(chart.chartType));
  if (type == null) throw new OfficeKitCodecError(`Presentation chart ${chart.id} supports only bar, line, pie, area, doughnut, marker scatter, 2D bubble, or bounded combo.`, [], { code: "unsupported_presentation_features" });
  const combo = type === SpreadsheetChartType.COMBO;
  const numericX = PRESENTATION_NUMERIC_X_CHART_TYPES.has(type);
  const circular = PRESENTATION_CIRCULAR_CHART_TYPES.has(type);
  const formulaErrorBarSeries = chart.series.findIndex((series) => series.errorBars?.plusFormula || series.errorBars?.minusFormula);
  if (formulaErrorBarSeries >= 0) {
    const side = chart.series[formulaErrorBarSeries].errorBars.plusFormula ? "plus" : "minus";
    throw new OfficeKitCodecError(
      `Presentation chart ${chart.id} series[${formulaErrorBarSeries}].errorBars.${side}Formula requires an embedded workbook path outside the bounded native chart profile.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  if (chart.externalData || chart.series.some((series) => series.categoryFormula || series.valueFormula || series.categoriesFormula || series.valuesFormula || series.xValueFormula || series.bubbleSizeFormula)) {
    throw new OfficeKitCodecError(`Presentation chart ${chart.id} must use literal categories and values.`, [], { code: "unsupported_presentation_features" });
  }
  if (!Array.isArray(chart.categories) || chart.categories.length > MAX_PRESENTATION_CHART_POINTS || chart.series.length < 1 || chart.series.length > 256) {
    throw new OfficeKitCodecError(`Presentation chart ${chart.id} exceeds the bounded category or series budget.`, [], { code: "invalid_presentation_chart" });
  }
  if (numericX && chart.categories.length) throw new OfficeKitCodecError(`Presentation ${chart.chartType} chart ${chart.id} must use per-series numeric xValues instead of shared categories.`, [], { code: "invalid_presentation_chart" });
  const originalSeries = originalChart?.type === SpreadsheetChartType.COMBO ? originalChart.comboSeries || [] : originalChart?.series || [];
  if (originalChart && (originalChart.type !== type || originalSeries.length !== chart.series.length || originalChart.categories.length !== chart.categories.length)) {
    throw new OfficeKitCodecError(`Presentation chart ${chart.id} cannot change its imported type, series count, or point topology.`, [], { code: "presentation_chart_topology_changed" });
  }
  const position = chart.position || {};
  const series = chart.series.map((item, index) => {
    const values = Array.isArray(item.values) ? item.values.map(Number) : [];
    const xValues = Array.isArray(item.xValues) ? item.xValues.map(Number) : [];
    const bubbleSizes = Array.isArray(item.bubbleSizes) ? item.bubbleSizes.map(Number) : [];
    if (values.some((value) => !Number.isFinite(value)) || values.length > MAX_PRESENTATION_CHART_POINTS) throw new OfficeKitCodecError(`Presentation chart ${chart.id} series ${index + 1} contains invalid or excessive values.`, [], { code: "invalid_presentation_chart" });
    if (numericX) {
      if (values.length < 1 || xValues.length !== values.length || xValues.some((value) => !Number.isFinite(value))) throw new OfficeKitCodecError(`Presentation ${chart.chartType} chart ${chart.id} series ${index + 1} must contain aligned finite xValues and values.`, [], { code: "invalid_presentation_chart" });
    } else if (values.length !== chart.categories.length || xValues.length) {
      throw new OfficeKitCodecError(`Presentation chart ${chart.id} series ${index + 1} must contain one finite value per category and no xValues.`, [], { code: "invalid_presentation_chart" });
    }
    if (type === SpreadsheetChartType.BUBBLE) {
      if (bubbleSizes.length !== values.length || bubbleSizes.some((value) => !Number.isFinite(value) || value <= 0)) throw new OfficeKitCodecError(`Presentation bubble chart ${chart.id} series ${index + 1} must contain one positive bubbleSize per value.`, [], { code: "invalid_presentation_chart" });
    } else if (bubbleSizes.length) throw new OfficeKitCodecError(`Presentation ${chart.chartType} chart ${chart.id} series ${index + 1} cannot carry bubbleSizes.`, [], { code: "invalid_presentation_chart" });
    const seriesType = combo ? PRESENTATION_CHART_TYPES_TO_WIRE.get(String(item.chartType || "")) : undefined;
    if (combo && seriesType !== SpreadsheetChartType.BAR && seriesType !== SpreadsheetChartType.LINE) throw new OfficeKitCodecError(`Presentation combo chart ${chart.id} series ${index + 1} must be bar or line.`, [], { code: "unsupported_presentation_features" });
    const axisGroup = item.axisGroup === "secondary" ? "secondary" : "primary";
    if ((!combo && (axisGroup === "secondary" || item.chartType)) || item.points?.length || item.dataLabels || item.smooth != null) {
      throw new OfficeKitCodecError(`Presentation chart ${chart.id} series ${index + 1} uses semantics outside the bounded native chart slice.`, [], { code: "unsupported_presentation_features" });
    }
    const effectiveType = combo ? String(item.chartType || "") : String(chart.chartType || "");
    const trendlines = chartTrendlines(item.trendlines, values.length, effectiveType, chart, `series[${index}].trendlines`, rgb);
    const errorBars = chartErrorBars(item.errorBars, values.length, effectiveType, chart, `series[${index}].errorBars`, rgb);
    if (item.marker && !["line", "scatter"].includes(effectiveType)) throw new OfficeKitCodecError(`Presentation ${effectiveType} chart ${chart.id} series ${index + 1} cannot carry a marker.`, [], { code: "unsupported_presentation_features" });
    if (type === SpreadsheetChartType.SCATTER && (item.line || item.stroke)) throw new OfficeKitCodecError(`Presentation marker-scatter chart ${chart.id} series ${index + 1} cannot carry a series line; use marker.line for marker borders.`, [], { code: "unsupported_presentation_features" });
    return {
      name: String(item.name || `Series ${index + 1}`),
      values,
      ...(numericX ? { xValues } : {}),
      ...(type === SpreadsheetChartType.BUBBLE ? { bubbleSizes } : {}),
      ...(item.fill || item.color ? { fill: chartColor(item.fill || item.color, chart, `series[${index}].fill`, rgb) } : {}),
      ...(item.line || item.stroke ? { line: chartLine(item.line || item.stroke, chart, `series[${index}].line`, rgb) } : {}),
      ...(item.marker ? { marker: chartMarker(item.marker, chart, `series[${index}].marker`, rgb) } : {}),
      ...(trendlines.length ? { trendlines } : {}),
      ...(errorBars ? { errorBars } : {}),
      ...(seriesType == null ? {} : { _comboType: seriesType, _comboAxisGroup: PRESENTATION_CHART_AXIS_GROUPS_TO_WIRE.get(axisGroup) }),
    };
  });
  if (combo && (!series.some((item) => item._comboType === SpreadsheetChartType.BAR) || !series.some((item) => item._comboType === SpreadsheetChartType.LINE))) throw new OfficeKitCodecError(`Presentation combo chart ${chart.id} requires at least one bar series and one line series.`, [], { code: "invalid_presentation_chart" });
  const comboBars = combo ? series.filter((item) => item._comboType === SpreadsheetChartType.BAR) : [];
  const comboLines = combo ? series.filter((item) => item._comboType === SpreadsheetChartType.LINE) : [];
  const hasSecondaryComboLine = comboLines.some((item) => item._comboAxisGroup === PresentationChartAxisGroup.SECONDARY);
  if (combo && comboBars.some((item) => item._comboAxisGroup !== PresentationChartAxisGroup.PRIMARY)) throw new OfficeKitCodecError(`Presentation combo chart ${chart.id} supports bars only on the primary axis pair.`, [], { code: "unsupported_presentation_features" });
  if (hasSecondaryComboLine && comboLines.some((item) => item._comboAxisGroup !== PresentationChartAxisGroup.SECONDARY)) throw new OfficeKitCodecError(`Presentation combo chart ${chart.id} cannot mix primary and secondary line plots.`, [], { code: "unsupported_presentation_features" });
  if (combo && originalChart?.type === SpreadsheetChartType.COMBO && originalSeries.some((item, index) =>
    item.type !== series[index]?._comboType ||
    (item.axisGroup === PresentationChartAxisGroup.SECONDARY ? PresentationChartAxisGroup.SECONDARY : PresentationChartAxisGroup.PRIMARY) !== series[index]?._comboAxisGroup,
  )) throw new OfficeKitCodecError(`Presentation chart ${chart.id} cannot change an imported combo series type or axis group.`, [], { code: "presentation_chart_topology_changed" });
  if (originalChart && originalSeries.some((item, index) => {
    const originalPayload = combo ? item.series : item;
    return (originalPayload?.trendlines?.length || 0) !== (series[index]?.trendlines?.length || 0);
  })) throw new OfficeKitCodecError(`Presentation chart ${chart.id} cannot change its imported trendline count.`, [], { code: "presentation_chart_topology_changed" });
  if (originalChart && originalSeries.some((item, index) => {
    const originalPayload = combo ? item.series : item;
    return Boolean(originalPayload?.errorBars) !== Boolean(series[index]?.errorBars);
  })) throw new OfficeKitCodecError(`Presentation chart ${chart.id} cannot add or remove imported error bars.`, [], { code: "presentation_chart_topology_changed" });
  const nativeSeries = series.map(({ _comboType, _comboAxisGroup, ...item }) => item);
  let xAxis = circular ? undefined : chartAxis(chart.axes?.category, chart, "xAxis", originalChart?.xAxis);
  let yAxis = circular ? undefined : chartAxis(chart.axes?.value, chart, "yAxis", originalChart?.yAxis);
  // DrawingML axis-bearing plots own a crossed pair. Preserve the historical
  // authoring convenience where callers configure only one visible axis while
  // still sending a complete, explicit pair to the strict native codec.
  if (!circular && Boolean(xAxis) !== Boolean(yAxis)) {
    xAxis ||= chartAxis({}, chart, "xAxis", {});
    yAxis ||= chartAxis({}, chart, "yAxis", {});
  }
  const secondaryXAxis = hasSecondaryComboLine ? chartAxis(chart.axes?.secondary?.category, chart, "secondaryXAxis", originalChart?.secondaryXAxis) : undefined;
  const secondaryYAxis = hasSecondaryComboLine ? chartAxis(chart.axes?.secondary?.value, chart, "secondaryYAxis", originalChart?.secondaryYAxis) : undefined;
  const dataLabels = chartDataLabels(chart.dataLabels, chart, originalChart?.dataLabels);
  return {
    id: original?.id || chart.id,
    name: chart.name || original?.name || chart.id,
    source: original?.source,
    content: {
      case: "chart",
      value: {
        leftEmu: sourceBoundFrameEmuFromPixels(position.left, `${chart.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(position.top, `${chart.id}.position.top`, original),
        widthEmu: emuFromPixels(position.width, `${chart.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${chart.id}.position.height`),
        type,
        title: String(chart.title || ""),
        hasLegend: Boolean(chart.legend?.visible ?? chart.hasLegend),
        categories: chart.categories.map((value) => String(value ?? "")),
        series: combo ? [] : nativeSeries,
        ...(combo ? { comboSeries: series.map(({ _comboType, _comboAxisGroup, ...item }) => ({ type: _comboType, axisGroup: _comboAxisGroup, series: item })) } : {}),
        ...(xAxis ? { xAxis } : {}),
        ...(yAxis ? { yAxis } : {}),
        ...(secondaryXAxis ? { secondaryXAxis } : {}),
        ...(secondaryYAxis ? { secondaryYAxis } : {}),
        ...(dataLabels ? { dataLabels } : {}),
      },
    },
  };
}

function modelChartColor(color) {
  if (!color) return undefined;
  if (color.source?.case !== "rgb") throw new OfficeKitCodecError("Presentation chart contains a non-RGB color outside the bounded chart slice.", [], { code: "invalid_presentation_chart" });
  return `#${color.source.value}`;
}

function modelChartLine(line) {
  if (!line) return undefined;
  const style = PRESENTATION_CHART_LINE_STYLES_FROM_WIRE.get(line.dashStyle);
  if (!style) throw new OfficeKitCodecError("Presentation chart contains an unsupported line style.", [], { code: "invalid_presentation_chart" });
  const presentationStyle = { dashed: "dash", dotted: "dot", "dash-dot": "dashDot", "dash-dot-dot": "longDashDotDot" }[style] || style;
  return {
    ...(line.color ? { fill: modelChartColor(line.color) } : {}),
    style: presentationStyle,
    ...(line.widthPoints === undefined ? {} : { width: line.widthPoints }),
  };
}

function modelChartMarker(marker) {
  if (!marker) return undefined;
  const symbol = PRESENTATION_CHART_MARKERS_FROM_WIRE.get(marker.symbol);
  if (!symbol) throw new OfficeKitCodecError("Presentation chart contains an unsupported series marker.", [], { code: "invalid_presentation_chart" });
  return {
    symbol,
    ...(marker.size === undefined ? {} : { size: marker.size }),
    ...(marker.fill ? { fill: modelChartColor(marker.fill) } : {}),
    ...(marker.line ? { line: modelChartLine(marker.line) } : {}),
  };
}

function modelChartTrendline(trendline) {
  const type = PRESENTATION_CHART_TRENDLINES_FROM_WIRE.get(trendline.type);
  if (!type) throw new OfficeKitCodecError("Presentation chart contains an unsupported trendline type.", [], { code: "invalid_presentation_chart" });
  return {
    type,
    ...(trendline.name ? { name: trendline.name } : {}),
    ...(trendline.polynomialOrder === undefined ? {} : { order: trendline.polynomialOrder }),
    ...(trendline.period === undefined ? {} : { period: trendline.period }),
    ...(trendline.forward === undefined ? {} : { forward: trendline.forward }),
    ...(trendline.backward === undefined ? {} : { backward: trendline.backward }),
    ...(trendline.intercept === undefined ? {} : { intercept: trendline.intercept }),
    displayEquation: Boolean(trendline.displayEquation),
    displayRSquared: Boolean(trendline.displayRSquared),
    ...(trendline.line ? { line: modelChartLine(trendline.line) } : {}),
  };
}

function modelChartErrorBarData(data, side) {
  if (!data) return {};
  return {
    ...(data.values?.length ? { [`${side}Values`]: [...data.values] } : {}),
    ...(data.formula ? { [`${side}Formula`]: data.formula } : {}),
    ...(data.formatCode ? { [`${side}FormatCode`]: data.formatCode } : {}),
  };
}

function modelChartErrorBars(errorBars) {
  if (!errorBars) return undefined;
  const direction = PRESENTATION_CHART_ERROR_BAR_DIRECTIONS_FROM_WIRE.get(errorBars.direction);
  const type = PRESENTATION_CHART_ERROR_BAR_TYPES_FROM_WIRE.get(errorBars.type);
  const valueType = PRESENTATION_CHART_ERROR_BAR_VALUE_TYPES_FROM_WIRE.get(errorBars.valueType);
  if (!direction || !type || !valueType) throw new OfficeKitCodecError("Presentation chart contains an unsupported error-bar direction, type, or value type.", [], { code: "invalid_presentation_chart" });
  return {
    direction,
    type,
    valueType,
    ...(errorBars.value === undefined ? {} : { value: errorBars.value }),
    ...modelChartErrorBarData(errorBars.plus, "plus"),
    ...modelChartErrorBarData(errorBars.minus, "minus"),
    noEndCap: Boolean(errorBars.noEndCap),
    ...(errorBars.line ? { line: modelChartLine(errorBars.line) } : {}),
  };
}

function modelChartAxis(axis, category) {
  if (!axis) return undefined;
  return {
    title: axis.title || "",
    ...(axis.numberFormatCode ? { numberFormatCode: axis.numberFormatCode } : {}),
    ...(category && axis.tickLabelInterval !== undefined ? { tickLabelInterval: axis.tickLabelInterval } : {}),
    ...(!category && axis.minimum !== undefined ? { min: axis.minimum } : {}),
    ...(!category && axis.maximum !== undefined ? { max: axis.maximum } : {}),
    ...(!category && axis.majorUnit !== undefined ? { majorUnit: axis.majorUnit } : {}),
  };
}

function modelChartDataLabels(labels) {
  if (!labels) return undefined;
  const position = labels.position === undefined ? undefined : PRESENTATION_CHART_LABEL_POSITIONS_FROM_WIRE.get(labels.position);
  if (labels.position !== undefined && !position) throw new OfficeKitCodecError("Presentation chart contains an unsupported data-label position.", [], { code: "invalid_presentation_chart" });
  return {
    showValue: Boolean(labels.showValue),
    showCategoryName: Boolean(labels.showCategoryName),
    ...(labels.showSeriesName === undefined ? {} : { showSeriesName: Boolean(labels.showSeriesName) }),
    ...(labels.showPercent === undefined ? {} : { showPercent: Boolean(labels.showPercent) }),
    ...(position ? { position } : {}),
  };
}

export function modelPresentationChartFromWire(source, emuPerPixel) {
  const chartType = PRESENTATION_CHART_TYPES_FROM_WIRE.get(source.type);
  if (!chartType) throw new OfficeKitCodecError("Presentation chart contains an unsupported chart type.", [], { code: "invalid_presentation_chart" });
  const combo = source.type === SpreadsheetChartType.COMBO;
  const numericX = PRESENTATION_NUMERIC_X_CHART_TYPES.has(source.type);
  const circular = PRESENTATION_CIRCULAR_CHART_TYPES.has(source.type);
  const sourceSeries = combo ? source.comboSeries || [] : source.series;
  if (combo && (!sourceSeries.some((item) => item.type === SpreadsheetChartType.BAR) || !sourceSeries.some((item) => item.type === SpreadsheetChartType.LINE))) throw new OfficeKitCodecError("Presentation combo chart must contain at least one bar and one line series.", [], { code: "invalid_presentation_chart" });
  const comboAxisGroup = (entry) => {
    if (entry.axisGroup === PresentationChartAxisGroup.UNSPECIFIED || entry.axisGroup === PresentationChartAxisGroup.PRIMARY) return "primary";
    const value = PRESENTATION_CHART_AXIS_GROUPS_FROM_WIRE.get(entry.axisGroup);
    if (!value) throw new OfficeKitCodecError("Presentation combo chart contains an unsupported axis group.", [], { code: "invalid_presentation_chart" });
    return value;
  };
  const comboBars = combo ? sourceSeries.filter((entry) => entry.type === SpreadsheetChartType.BAR) : [];
  const comboLines = combo ? sourceSeries.filter((entry) => entry.type === SpreadsheetChartType.LINE) : [];
  const hasSecondaryComboLine = comboLines.some((entry) => comboAxisGroup(entry) === "secondary");
  if (combo && (comboBars.some((entry) => comboAxisGroup(entry) !== "primary") || (hasSecondaryComboLine && comboLines.some((entry) => comboAxisGroup(entry) !== "secondary")))) {
    throw new OfficeKitCodecError("Presentation combo chart does not match the canonical primary-bar/secondary-line topology.", [], { code: "invalid_presentation_chart" });
  }
  if (hasSecondaryComboLine && (!source.secondaryXAxis || !source.secondaryYAxis)) throw new OfficeKitCodecError("Presentation secondary-axis combo chart is missing its paired secondary axes.", [], { code: "invalid_presentation_chart" });
  const axes = circular ? undefined : {
    category: modelChartAxis(source.xAxis, !numericX) || { title: "" },
    value: modelChartAxis(source.yAxis, false) || { title: "" },
    ...(hasSecondaryComboLine ? { secondary: { category: modelChartAxis(source.secondaryXAxis, true) || { title: "" }, value: modelChartAxis(source.secondaryYAxis, false) || { title: "" } } } : {}),
  };
  return {
    chartType,
    position: {
      left: Number(source.leftEmu) / emuPerPixel,
      top: Number(source.topEmu) / emuPerPixel,
      width: Number(source.widthEmu) / emuPerPixel,
      height: Number(source.heightEmu) / emuPerPixel,
    },
    title: source.title,
    categories: [...source.categories],
    series: sourceSeries.map((entry) => {
      const series = combo ? entry.series : entry;
      const seriesType = combo ? PRESENTATION_CHART_TYPES_FROM_WIRE.get(entry.type) : undefined;
      if (combo && !new Set(["bar", "line"]).has(seriesType)) throw new OfficeKitCodecError("Presentation combo chart contains an unsupported series type.", [], { code: "invalid_presentation_chart" });
      return {
        name: series.name,
        values: [...series.values],
        ...(numericX ? { xValues: [...series.xValues] } : {}),
        ...(source.type === SpreadsheetChartType.BUBBLE ? { bubbleSizes: [...series.bubbleSizes] } : {}),
        ...(series.fill ? { fill: modelChartColor(series.fill) } : {}),
        ...(series.line ? { line: modelChartLine(series.line) } : {}),
        ...(series.marker ? { marker: modelChartMarker(series.marker) } : {}),
        ...((series.trendlines || []).length ? { trendlines: series.trendlines.map(modelChartTrendline) } : {}),
        ...(series.errorBars ? { errorBars: modelChartErrorBars(series.errorBars) } : {}),
        ...(seriesType ? { chartType: seriesType } : {}),
        ...(combo && comboAxisGroup(entry) === "secondary" ? { axisGroup: "secondary" } : {}),
      };
    }),
    hasLegend: source.hasLegend,
    ...(axes ? { axes } : {}),
    ...(source.dataLabels ? { dataLabels: modelChartDataLabels(source.dataLabels) } : {}),
  };
}
