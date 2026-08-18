import assert from "node:assert/strict";
import { test } from "node:test";
import { displayUserText, lastUserText, visibleMessages } from "./viewModel";

test("displayUserText strips the Excel selection suffix", () => {
  assert.equal(displayUserText("fix GP"), "fix GP");
  assert.equal(displayUserText("fix GP\n\nExcel selection: P&L!A1:B6 · 6×2"), "fix GP");
});

test("visibleMessages hides the selection suffix on user bubbles", () => {
  const list = visibleMessages([
    { role: "user", content: "fix GP\n\nExcel selection: P&L!A1:B6 · 6×2" },
  ]);
  assert.equal(list[0]?.kind, "user");
  if (list[0]?.kind === "user") {
    assert.equal(list[0].text, "fix GP");
  }
});

test("lastUserText strips the selection suffix", () => {
  assert.equal(
    lastUserText([{ role: "user", content: "fix GP\n\nExcel selection: P&L!A1:B6 · 6×2" }]),
    "fix GP"
  );
});

test("pending tool calls are present tense, not Wrote", () => {
  const list = visibleMessages([
    { role: "user", content: "fix GP" },
    {
      role: "assistant",
      content: null,
      tool_calls: [
        { id: "1", name: "write_range", args: { sheet: "P&L", start_cell: "B4", values: [["=B2-B3"]] } },
      ],
    },
  ]);
  assert.equal(list[1]?.kind, "steps");
  if (list[1]?.kind === "steps") {
    assert.equal(list[1].lines[0]?.pending, true);
    assert.equal(list[1].lines[0]?.failed, false);
    assert.equal(list[1].lines[0]?.text, "Writing P&L!B4 · =B2-B3");
  }
});

test("cancelled tool results render as failed, not wrote", () => {
  const list = visibleMessages([
    {
      role: "assistant",
      content: null,
      tool_calls: [{ id: "1", name: "write_range", args: { sheet: "P&L", start_cell: "B4" } }],
    },
    { role: "tool", tool_call_id: "1", content: JSON.stringify({ error: "cancelled" }) },
  ]);
  assert.equal(list[0]?.kind, "steps");
  if (list[0]?.kind === "steps") {
    assert.equal(list[0].lines[0]?.pending, false);
    assert.equal(list[0].lines[0]?.failed, true);
    assert.equal(list[0].lines[0]?.text, "Failed: stopped");
  }
});
