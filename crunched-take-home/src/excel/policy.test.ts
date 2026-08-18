import assert from "node:assert/strict";
import { test } from "node:test";
import { MAX_CELLS } from "../shared/limits";
import {
  cellCount,
  chartAlias,
  exceedsCellCap,
  isA1Cell,
  isFormulaCell,
  plannedWrites,
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

test("isA1Cell accepts a cell and rejects a range", () => {
  assert.equal(isA1Cell("B4"), true);
  assert.equal(isA1Cell("$B$4"), true);
  assert.equal(isA1Cell("AA10"), true);
  assert.equal(isA1Cell("A1:B2"), false);
  assert.equal(isA1Cell("P&L!B4"), false);
});

test("plannedWrites skips omitted cells instead of padding null", () => {
  assert.deepEqual(plannedWrites([["Revenue"], ["COGS", 400]]), [
    { row: 0, column: 0, value: "Revenue", formula: false },
    { row: 1, column: 0, value: "COGS", formula: false },
    { row: 1, column: 1, value: 400, formula: false },
  ]);
});

test("plannedWrites splits labels and formulas in one block", () => {
  const writes = plannedWrites([["Gross Profit", "=B2-B3"], ["Operating Profit", " =B4-B5"]]);
  assert.deepEqual(writes, [
    { row: 0, column: 0, value: "Gross Profit", formula: false },
    { row: 0, column: 1, value: "=B2-B3", formula: true },
    { row: 1, column: 0, value: "Operating Profit", formula: false },
    { row: 1, column: 1, value: "=B4-B5", formula: true },
  ]);
});

test("isFormulaCell trims a leading space before =", () => {
  assert.equal(isFormulaCell("=B2-B3"), true);
  assert.equal(isFormulaCell(" =B2-B3"), true);
  assert.equal(isFormulaCell("600"), false);
  assert.equal(isFormulaCell(600), false);
});

test("chartAlias maps Office.js enums and tool literals", () => {
  assert.equal(chartAlias("column"), "column");
  assert.equal(chartAlias("ColumnClustered"), "column");
  assert.equal(chartAlias("barClustered"), "bar");
  assert.equal(chartAlias("Line"), "line");
  assert.equal(chartAlias("pie"), "pie");
  assert.equal(chartAlias("radar"), null);
});
