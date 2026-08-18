import { MAX_CELLS } from "../shared/limits";

export type JsonValue = string | number | boolean | null;

export function cellCount(rows: number, columns: number): number {
  return rows * columns;
}

export function tooLarge(rows: number, columns: number): Record<string, unknown> {
  return {
    error: "range_too_large",
    rows,
    columns,
    cells: cellCount(rows, columns),
    max_cells: MAX_CELLS,
    hint: "Read or write a smaller A1 slice (for example A1:H50).",
  };
}

export function padRows(values: JsonValue[][]): JsonValue[][] {
  const columns = values.reduce((max, row) => Math.max(max, row.length), 0);
  return values.map((row) => {
    if (row.length === columns) {
      return row;
    }
    return [...row, ...Array(columns - row.length).fill(null)];
  });
}

export function isFormulaCell(cell: JsonValue): boolean {
  return typeof cell === "string" && cell.startsWith("=");
}

export function hasFormula(values: JsonValue[][]): boolean {
  return values.some((row) => row.some(isFormulaCell));
}

export function exceedsCellCap(rows: number, columns: number): boolean {
  return cellCount(rows, columns) > MAX_CELLS;
}
