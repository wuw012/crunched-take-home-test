import assert from "node:assert/strict";
import { test } from "node:test";
import { formatSelectionChip } from "./selection";

test("formatSelectionChip uses sheet!address and size", () => {
  assert.equal(
    formatSelectionChip({ sheet: "P&L", address: "A1:B6", rows: 6, columns: 2 }),
    "P&L!A1:B6 · 6×2"
  );
});

test("formatSelectionChip keeps an address that already includes the sheet", () => {
  assert.equal(
    formatSelectionChip({ sheet: "P&L", address: "P&L!A1:B6", rows: 6, columns: 2 }),
    "P&L!A1:B6 · 6×2"
  );
});

test("formatSelectionChip still shows size when values were refused", () => {
  assert.equal(
    formatSelectionChip({
      sheet: "Exports",
      address: "A1:Z200",
      rows: 200,
      columns: 26,
      error: "range_too_large",
      values: null,
    }),
    "Exports!A1:Z200 · 200×26"
  );
});

test("formatSelectionChip returns null for empty payloads", () => {
  assert.equal(formatSelectionChip(null), null);
  assert.equal(formatSelectionChip({}), null);
});
