import { normalizeChartErrorBars } from "../shared/chart-error-bars.mjs";
import { normalizeSpreadsheetChartLineStyle } from "./chart-line-style.mjs";

export function normalizeSpreadsheetChartErrorBars(value, valueCount, chartType) {
  return normalizeChartErrorBars(value, {
    valueCount,
    chartType,
    normalizeLine: (line) => normalizeSpreadsheetChartLineStyle(line, "errorBars.line"),
  });
}
