import { normalizeChartTrendlines } from "../shared/chart-trendlines.mjs";
import { normalizeSpreadsheetChartLineStyle } from "./chart-line-style.mjs";

export function normalizeSpreadsheetChartTrendlines(value, valueCount, chartType) {
  return normalizeChartTrendlines(value, {
    valueCount,
    chartType,
    normalizeLine: (line) => normalizeSpreadsheetChartLineStyle(line, "trendline.line"),
  });
}
