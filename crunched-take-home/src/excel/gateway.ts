/* global Excel */

import { MAX_CELLS } from "../shared/limits";
import {
  cellCount,
  exceedsCellCap,
  hasFormula,
  JsonValue,
  padRows,
  tooLarge,
} from "./policy";

export type { JsonValue };

async function loadBoundedRange(
  context: Excel.RequestContext,
  range: Excel.Range,
  sheet: string
): Promise<Record<string, unknown>> {
  range.load(["address", "rowCount", "columnCount"]);
  await context.sync();

  const rows = range.rowCount;
  const columns = range.columnCount;
  if (exceedsCellCap(rows, columns)) {
    return {
      sheet,
      address: range.address,
      ...tooLarge(rows, columns),
      values: null,
      formulas: null,
    };
  }

  range.load(["values", "formulas"]);
  await context.sync();
  return {
    sheet,
    address: range.address,
    rows,
    columns,
    values: range.values,
    formulas: range.formulas,
  };
}

export async function listWorkbookStructure(): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const sheetMeta = sheets.items.map((sheet) => {
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["address", "rowCount", "columnCount"]);
      const charts = sheet.charts;
      charts.load("items/name,items/chartType,items/title/text");
      return { sheet, used, charts };
    });
    await context.sync();

    return {
      sheets: sheetMeta.map(({ sheet, used, charts }) => ({
        name: sheet.name,
        used_range: used.isNullObject
          ? null
          : {
              address: used.address,
              rows: used.rowCount,
              columns: used.columnCount,
            },
        charts: charts.items.map((chart) => ({
          name: chart.name,
          type: chart.chartType,
          title: chart.title.text || null,
        })),
      })),
    };
  });
}

function chartTypes(): Record<string, Excel.ChartType> {
  return {
    column: Excel.ChartType.columnClustered,
    bar: Excel.ChartType.barClustered,
    line: Excel.ChartType.line,
    pie: Excel.ChartType.pie,
  };
}

/** Default placement: right of a typical P&L stub, below header rows. */
const DEFAULT_CHART_TOP_LEFT = "D8";
const DEFAULT_CHART_BOTTOM_RIGHT = "L22";

function resolveChartType(raw: string): Excel.ChartType | null {
  const types = chartTypes();
  const key = raw.trim().toLowerCase();
  return types[key] ?? null;
}

function unknownChartType(chartType: string): Record<string, unknown> {
  return { error: "unknown_chart_type", chart_type: chartType, allowed: Object.keys(chartTypes()) };
}

export async function createChart(
  sheet: string,
  sourceA1: string,
  chartType: string,
  title?: string
): Promise<unknown> {
  const type = resolveChartType(chartType);
  if (!type) {
    return unknownChartType(chartType);
  }

  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const range = worksheet.getRange(sourceA1);
    const chart = worksheet.charts.add(type, range, Excel.ChartSeriesBy.auto);
    chart.setPosition(DEFAULT_CHART_TOP_LEFT, DEFAULT_CHART_BOTTOM_RIGHT);
    if (title) {
      chart.title.text = title;
    }
    chart.load(["name", "chartType"]);
    await context.sync();
    return {
      sheet,
      source: sourceA1,
      name: chart.name,
      type: chart.chartType,
    };
  });
}

export async function setChartType(sheet: string, chartType: string, name?: string): Promise<unknown> {
  const type = resolveChartType(chartType);
  if (!type) {
    return unknownChartType(chartType);
  }

  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const charts = worksheet.charts;
    charts.load("items/name,items/chartType");
    await context.sync();
    if (charts.items.length === 0) {
      return { error: "no_charts", sheet };
    }

    const chart = name
      ? worksheet.charts.getItem(name)
      : worksheet.charts.getItemAt(charts.items.length - 1);
    chart.chartType = type;
    chart.load(["name", "chartType"]);
    await context.sync();
    return { sheet, name: chart.name, type: chart.chartType };
  });
}

export async function getSelection(): Promise<unknown> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    sheet.load("name");
    await context.sync();
    return loadBoundedRange(context, range, sheet.name);
  });
}

export async function readRange(sheet: string, a1: string): Promise<unknown> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const range = worksheet.getRange(a1);
    return loadBoundedRange(context, range, sheet);
  });
}

export async function writeRange(
  sheet: string,
  startCell: string,
  values: JsonValue[][]
): Promise<unknown> {
  if (!Array.isArray(values) || values.length === 0) {
    return { error: "values must be a non-empty 2D array" };
  }

  const padded = padRows(values);
  const rows = padded.length;
  const columns = padded[0]?.length ?? 0;
  if (rows === 0 || columns === 0) {
    return { error: "values must be a non-empty 2D array" };
  }
  if (cellCount(rows, columns) > MAX_CELLS) {
    return tooLarge(rows, columns);
  }

  const wroteFormulas = hasFormula(padded);

  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const range = worksheet.getRange(startCell).getResizedRange(rows - 1, columns - 1);
    if (wroteFormulas) {
      range.formulas = padded as unknown as string[][];
    } else {
      range.values = padded as unknown as string[][];
    }
    await context.sync();
    range.load("address");
    await context.sync();
    return {
      sheet,
      address: range.address,
      rows,
      columns,
      wrote_formulas: wroteFormulas,
    };
  });
}
