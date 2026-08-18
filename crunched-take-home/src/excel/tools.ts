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

function asString(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() ? value : fallback;
}

function asGrid(value: unknown): JsonValue[][] {
  if (!Array.isArray(value)) {
    throw new Error("values must be a 2D array");
  }
  return value.map((row) => (Array.isArray(row) ? (row as JsonValue[]) : [row as JsonValue]));
}

export async function executeTool(name: string, args: Record<string, unknown>): Promise<string> {
  try {
    switch (name) {
      case "list_workbook_structure":
        return truncate(await listWorkbookStructure());
      case "get_selection":
        return truncate(await getSelection());
      case "read_range":
        return truncate(await readRange(asString(args.sheet, ""), asString(args.a1, "")));
      case "write_range":
        return truncate(
          await writeRange(asString(args.sheet, ""), asString(args.start_cell, "A1"), asGrid(args.values))
        );
      case "create_chart":
        return truncate(
          await createChart(
            asString(args.sheet, ""),
            asString(args.source_a1, ""),
            asString(args.chart_type, ""),
            typeof args.title === "string" ? args.title : undefined
          )
        );
      case "set_chart_type":
        return truncate(
          await setChartType(
            asString(args.sheet, ""),
            asString(args.chart_type, ""),
            typeof args.name === "string" ? args.name : undefined
          )
        );
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
      return `Read ${asString(args.sheet, "sheet")}!${asString(args.a1, "range")}`;
    case "write_range":
      return `Wrote ${asString(args.sheet, "sheet")}!${asString(args.start_cell, "A1")}`;
    case "create_chart":
      return `Charted ${asString(args.sheet, "sheet")}!${asString(args.source_a1, "range")} · ${asString(args.chart_type, "chart")}`;
    case "set_chart_type":
      return `Retargeted chart · ${asString(args.chart_type, "type")}`;
    default:
      return name;
  }
}
