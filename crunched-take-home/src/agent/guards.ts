import { ToolCall } from "../api/types";
import { MAX_INSPECT_ROUNDS } from "../shared/limits";

export { MAX_INSPECT_ROUNDS };

export const INSPECT_TOOLS = new Set(["list_workbook_structure", "get_selection", "read_range"]);

export function toolRoundKey(calls: ToolCall[]): string {
  return JSON.stringify(calls.map((call) => ({ name: call.name, args: call.args })));
}

export function isInspectOnly(calls: ToolCall[]): boolean {
  return calls.length > 0 && calls.every((call) => INSPECT_TOOLS.has(call.name));
}

export function isStructureOnly(calls: ToolCall[]): boolean {
  return calls.length > 0 && calls.every((call) => call.name === "list_workbook_structure");
}

/** Inspect rounds that read cells. Structure-only listing does not increment the cap. */
export function isCellInspect(calls: ToolCall[]): boolean {
  return isInspectOnly(calls) && !isStructureOnly(calls);
}

export function shouldBlockInspect(calls: ToolCall[], inspectRounds: number): boolean {
  return isInspectOnly(calls) && inspectRounds >= MAX_INSPECT_ROUNDS;
}
