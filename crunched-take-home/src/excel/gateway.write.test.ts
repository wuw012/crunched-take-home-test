import assert from "node:assert/strict";
import { afterEach, test } from "node:test";
import { writeRange } from "./gateway";

type Cell = { formulas?: unknown; values?: unknown };

function installExcelMock(address = "P&L!A4:B5"): Map<string, Cell> {
  const cells = new Map<string, Cell>();
  (globalThis as { Excel?: unknown }).Excel = {
    run: async (fn: (context: unknown) => Promise<unknown>) => {
      const origin = {
        getOffsetRange(row: number, column: number) {
          const key = `${row},${column}`;
          if (!cells.has(key)) {
            cells.set(key, {});
          }
          return cells.get(key);
        },
        getResizedRange() {
          return {
            load() {
              return undefined;
            },
            address,
          };
        },
      };
      return fn({
        workbook: {
          worksheets: {
            getItem() {
              return { getRange() { return origin; } };
            },
          },
        },
        async sync() {
          return undefined;
        },
      });
    },
  };
  return cells;
}

afterEach(() => {
  delete (globalThis as { Excel?: unknown }).Excel;
});

test("mixed write assigns values and formulas on different cells", async () => {
  const cells = installExcelMock();
  const result = (await writeRange("P&L", "A4", [["Gross Profit", "=B2-B3"]])) as {
    wrote_formulas: boolean;
    cells: number;
  };
  assert.equal(result.wrote_formulas, true);
  assert.equal(result.cells, 2);
  assert.deepEqual(cells.get("0,0")?.values, [["Gross Profit"]]);
  assert.equal(cells.get("0,0")?.formulas, undefined);
  assert.deepEqual(cells.get("0,1")?.formulas, [["=B2-B3"]]);
  assert.equal(cells.get("0,1")?.values, undefined);
});

test("ragged write does not assign the omitted neighbor", async () => {
  const cells = installExcelMock("P&L!A1:B2");
  await writeRange("P&L", "A1", [["Revenue"], ["COGS", 400]]);
  assert.equal(cells.has("0,1"), false);
  assert.deepEqual(cells.get("0,0")?.values, [["Revenue"]]);
  assert.deepEqual(cells.get("1,0")?.values, [["COGS"]]);
  assert.deepEqual(cells.get("1,1")?.values, [[400]]);
});

test("leading space still writes a formula, not a value", async () => {
  const cells = installExcelMock("P&L!B4");
  await writeRange("P&L", "B4", [[" =B2-B3"]]);
  assert.deepEqual(cells.get("0,0")?.formulas, [["=B2-B3"]]);
  assert.equal(cells.get("0,0")?.values, undefined);
});
