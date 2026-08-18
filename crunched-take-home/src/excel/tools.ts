import { MAX_TOOL_RESULT_CHARS } from "../shared/limits";
import {
  createChart,
  getSelection,
  listWorkbookStructure,
  setChartType,
  writeRange,
  readRange,
  JsonValue,
} from "./gateway";

function truncate(payload: unknown): string {
  const text = JSON.stringify(payload);
  if (text.length <= MAX_TOOL_RESULT_CHARS) {
    return text;
  }
  return JSON.stringify({
    truncated: true,
    max_chars: MAX_TOOL_RESULT_CHARS,
    hint: "Result was truncated. Read a smaller range.",
    preview: text.slice(0, MAX_TOOL_RESULT_CHARS),
  });
}

function asString(value: unknown, fallback?: string): string {
  if (typeof value === "string" && value.trim()) {
    return value.trim();
  }
  return fallback ?? "";
}

function readString(args: Record<string, unknown>, name: string): string | undefined {
  const value = args[name];
  if (typeof value === "string" && value.trim()) {
    return value.trim();
  }
  return undefined;
}

function missingArg(name: string): string {
  return truncate({ error: "missing_arg", name });
}

function asGrid(value: unknown): JsonValue[][] {
  if (!Array.isArray(value)) {
    throw new Error("values must be a 2D array");
  }
  return value.map((row) => (Array.isArray(row) ? (row as JsonValue[]) : [row as JsonValue]));
}

function writeSnippet(args: Record<string, unknown>): string {
  const values = args.values;
  if (!Array.isArray(values)) {
    return "";
  }
  const cells: string[] = [];
  for (const row of values) {
    if (!Array.isArray(row)) {
      continue;
    }
    for (const cell of row) {
      if (cell === null || cell === undefined || cell === "") {
        continue;
      }
      cells.push(String(cell));
      if (cells.length >= 4) {
        break;
      }
    }
    if (cells.length >= 4) {
      break;
    }
  }
  if (cells.length === 0) {
    return "";
  }
  const text = cells.join(", ");
  return text.length > 48 ? `${text.slice(0, 47)}…` : text;
}

function toolTarget(name: string, args: Record<string, unknown>): string {
  if (name === "read_range") {
    return `${asString(args.sheet, "sheet")}!${asString(args.a1, "range")}`;
  }
  if (name === "write_range") {
    return `${asString(args.sheet, "sheet")}!${asString(args.start_cell, "A1")}`;
  }
  if (name === "create_chart") {
    return `${asString(args.sheet, "sheet")}!${asString(args.source_a1, "range")}`;
  }
  if (name === "get_selection") {
    return "selection";
  }
  return name;
}

export function parseToolError(content: string): string | null {
  try {
    const parsed = JSON.parse(content) as unknown;
    if (parsed && typeof parsed === "object" && "error" in parsed) {
      const error = (parsed as { error: unknown }).error;
      return typeof error === "string" ? error : "error";
    }
  } catch {
    return null;
  }
  return null;
}

export async function executeTool(name: string, args: Record<string, unknown>): Promise<string> {
  try {
    switch (name) {
      case "list_workbook_structure":
        return truncate(await listWorkbookStructure());
      case "get_selection":
        return truncate(await getSelection());
      case "read_range": {
        const sheet = readString(args, "sheet");
        const a1 = readString(args, "a1");
        if (!sheet) {
          return missingArg("sheet");
        }
        if (!a1) {
          return missingArg("a1");
        }
        return truncate(await readRange(sheet, a1));
      }
      case "write_range": {
        const sheet = readString(args, "sheet");
        const start = readString(args, "start_cell");
        if (!sheet) {
          return missingArg("sheet");
        }
        if (!start) {
          return missingArg("start_cell");
        }
        if (!("values" in args)) {
          return missingArg("values");
        }
        return truncate(await writeRange(sheet, start, asGrid(args.values)));
      }
      case "create_chart": {
        const sheet = readString(args, "sheet");
        const source = readString(args, "source_a1");
        const chartType = readString(args, "chart_type");
        if (!sheet) {
          return missingArg("sheet");
        }
        if (!source) {
          return missingArg("source_a1");
        }
        if (!chartType) {
          return missingArg("chart_type");
        }
        return truncate(
          await createChart(
            sheet,
            source,
            chartType,
            typeof args.title === "string" ? args.title : undefined
          )
        );
      }
      case "set_chart_type": {
        const sheet = readString(args, "sheet");
        const chartType = readString(args, "chart_type");
        if (!sheet) {
          return missingArg("sheet");
        }
        if (!chartType) {
          return missingArg("chart_type");
        }
        return truncate(
          await setChartType(sheet, chartType, typeof args.name === "string" ? args.name : undefined)
        );
      }
      default:
        return truncate({ error: "unknown_tool", name });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return truncate({ error: "excel_error", message });
  }
}

export function describeTool(name: string, args: Record<string, unknown>): string {
  switch (name) {
    case "list_workbook_structure":
      return "Inspected workbook structure";
    case "get_selection":
      return "Read current selection";
    case "read_range":
      return `Read ${toolTarget(name, args)}`;
    case "write_range": {
      const snippet = writeSnippet(args);
      const target = toolTarget(name, args);
      return snippet ? `Wrote ${target} · ${snippet}` : `Wrote ${target}`;
    }
    case "create_chart":
      return `Charted ${toolTarget(name, args)} · ${asString(args.chart_type, "chart")}`;
    case "set_chart_type":
      return `Retargeted chart · ${asString(args.chart_type, "type")}`;
    default:
      return name;
  }
}

export function describeToolFailure(name: string, args: Record<string, unknown>, content: string): string {
  const error = parseToolError(content) ?? "error";
  const target = toolTarget(name, args);
  if (error === "range_too_large") {
    return `Failed: ${target} too large`;
  }
  if (error === "too_many_reads") {
    return "Failed: too many reads";
  }
  if (error === "missing_arg") {
    try {
      const parsed = JSON.parse(content) as { name?: string };
      return parsed.name ? `Failed: missing ${parsed.name}` : "Failed: missing argument";
    } catch {
      return "Failed: missing argument";
    }
  }
  if (error === "unknown_tool") {
    return `Failed: unknown tool ${name}`;
  }
  if (error === "excel_error") {
    return `Failed: ${target}`;
  }
  return `Failed: ${target}`;
}
