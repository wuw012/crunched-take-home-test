import assert from "node:assert/strict";
import { test } from "node:test";
import { ToolCall } from "../api/types";
import {
  INSPECT_TOOLS,
  isCellInspect,
  isInspectOnly,
  isStructureOnly,
  MAX_INSPECT_ROUNDS,
  shouldBlockInspect,
  toolRoundKey,
} from "./guards";

function call(name: string, args: Record<string, unknown> = {}): ToolCall {
  return { id: "1", name, args };
}

test("toolRoundKey is stable for the same calls", () => {
  const a = [call("read_range", { sheet: "P&L", a1: "A1:B6" })];
  const b = [call("read_range", { sheet: "P&L", a1: "A1:B6" })];
  assert.equal(toolRoundKey(a), toolRoundKey(b));
});

test("toolRoundKey changes when args change", () => {
  const a = [call("read_range", { sheet: "P&L", a1: "A1:B6" })];
  const b = [call("read_range", { sheet: "P&L", a1: "A1:B7" })];
  assert.notEqual(toolRoundKey(a), toolRoundKey(b));
});

test("isInspectOnly is true only when every call is an inspect tool", () => {
  assert.equal(isInspectOnly([call("read_range"), call("get_selection")]), true);
  assert.equal(isInspectOnly([call("list_workbook_structure")]), true);
  assert.equal(isInspectOnly([call("read_range"), call("write_range")]), false);
  assert.equal(isInspectOnly([]), false);
  assert.ok(INSPECT_TOOLS.has("read_range"));
});

test("shouldBlockInspect after MAX_INSPECT_ROUNDS", () => {
  const reads = [call("read_range", { sheet: "P&L", a1: "A1" })];
  assert.equal(shouldBlockInspect(reads, MAX_INSPECT_ROUNDS - 1), false);
  assert.equal(shouldBlockInspect(reads, MAX_INSPECT_ROUNDS), true);
  assert.equal(shouldBlockInspect([call("write_range")], MAX_INSPECT_ROUNDS), false);
});

test("isStructureOnly does not count as cell inspect", () => {
  assert.equal(isStructureOnly([call("list_workbook_structure")]), true);
  assert.equal(isCellInspect([call("list_workbook_structure")]), false);
  assert.equal(isCellInspect([call("read_range")]), true);
  assert.equal(isCellInspect([call("get_selection"), call("read_range")]), true);
  assert.equal(isCellInspect([call("list_workbook_structure"), call("read_range")]), true);
});
