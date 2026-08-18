import assert from "node:assert/strict";
import { test } from "node:test";
import { MAX_CELLS } from "../shared/limits";
import {
  cellCount,
  exceedsCellCap,
  hasFormula,
  isFormulaCell,
  padRows,
  tooLarge,
} from "./policy";

test("cellCount multiplies rows by columns", () => {
  assert.equal(cellCount(50, 50), 2500);
  assert.equal(cellCount(1, 1), 1);
});

test("exceedsCellCap matches MAX_CELLS", () => {
  assert.equal(exceedsCellCap(44, 45), false);
  assert.equal(exceedsCellCap(50, 50), true);
  assert.ok(50 * 50 > MAX_CELLS);
});

test("tooLarge is refuse-before-load JSON", () => {
  assert.deepEqual(tooLarge(50, 40), {
    error: "range_too_large",
    rows: 50,
    columns: 40,
    cells: 2000,
    max_cells: MAX_CELLS,
    hint: "Read or write a smaller A1 slice (for example A1:H50).",
  });
});

test("padRows pads short rows with null", () => {
  assert.deepEqual(padRows([[1], [2, 3]]), [
    [1, null],
    [2, 3],
  ]);
});

test("isFormulaCell and hasFormula detect leading =", () => {
  assert.equal(isFormulaCell("=B2-B3"), true);
  assert.equal(isFormulaCell("600"), false);
  assert.equal(isFormulaCell(600), false);
  assert.equal(hasFormula([["Gross Profit", "=B2-B3"]]), true);
  assert.equal(hasFormula([["Gross Profit", 600]]), false);
});
