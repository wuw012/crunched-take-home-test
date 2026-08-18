import { ToolCall } from "../api/types";

export const MAX_INSPECT_ROUNDS = 3;

export const INSPECT_TOOLS = new Set(["list_workbook_structure", "get_selection", "read_range"]);

export function toolRoundKey(calls: ToolCall[]): string {
  return JSON.stringify(calls.map((call) => ({ name: call.name, args: call.args })));
}

export function isInspectOnly(calls: ToolCall[]): boolean {
  return calls.length > 0 && calls.every((call) => INSPECT_TOOLS.has(call.name));
}

export function shouldBlockInspect(calls: ToolCall[], inspectRounds: number): boolean {
  return isInspectOnly(calls) && inspectRounds >= MAX_INSPECT_ROUNDS;
}
