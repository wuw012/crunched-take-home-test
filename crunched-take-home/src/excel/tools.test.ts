import assert from "node:assert/strict";
import { test } from "node:test";
import { describeTool, describeToolFailure, executeTool, parseToolError } from "./tools";

test("describeTool includes a formula snippet on write_range", () => {
  const line = describeTool("write_range", {
    sheet: "P&L",
    start_cell: "B4",
    values: [["=B2-B3"]],
  });
  assert.equal(line, "Wrote P&L!B4 · =B2-B3");
});

test("describeToolFailure names a range that is too large", () => {
  const line = describeToolFailure(
    "read_range",
    { sheet: "P&L", a1: "A1:B6" },
    JSON.stringify({ error: "range_too_large", rows: 50, columns: 50 })
  );
  assert.equal(line, "Failed: P&L!A1:B6 too large");
});

test("parseToolError reads the error key", () => {
  assert.equal(parseToolError(JSON.stringify({ error: "missing_arg", name: "sheet" })), "missing_arg");
  assert.equal(parseToolError(JSON.stringify({ sheet: "P&L" })), null);
});

test("executeTool returns missing_arg without calling Excel", async () => {
  const content = await executeTool("read_range", {});
  assert.equal(JSON.parse(content).error, "missing_arg");
  assert.equal(JSON.parse(content).name, "sheet");
});
