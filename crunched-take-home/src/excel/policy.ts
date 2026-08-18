import { MAX_CELLS } from "../shared/limits";

export type JsonValue = string | number | boolean | null;

export type CellWrite = {
  row: number;
  column: number;
  value: JsonValue;
  formula: boolean;
};

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

/** Single cell like B4 or $B$4. Rejects A1:B2. */
export function isA1Cell(address: string): boolean {
  return /^\$?[A-Za-z]+\$?[1-9]\d*$/.test(address.trim());
}

export function isFormulaCell(cell: JsonValue): boolean {
  return typeof cell === "string" && cell.trimStart().startsWith("=");
}

/** Skip null/omitted cells so a ragged row does not clear neighbors. */
export function plannedWrites(values: JsonValue[][]): CellWrite[] {
  const writes: CellWrite[] = [];
  for (let row = 0; row < values.length; row += 1) {
    const line = values[row];
    for (let column = 0; column < line.length; column += 1) {
      const cell = line[column];
      if (cell === null || cell === undefined) {
        continue;
      }
      const formula = isFormulaCell(cell);
      writes.push({
        row,
        column,
        value: formula && typeof cell === "string" ? cell.trim() : cell,
        formula,
      });
    }
  }
  return writes;
}

export function exceedsCellCap(rows: number, columns: number): boolean {
  return cellCount(rows, columns) > MAX_CELLS;
}

const CHART_ALIASES = ["column", "bar", "line", "pie"] as const;
export type ChartAlias = (typeof CHART_ALIASES)[number];

/** Map tool literals and Office.js enum names (ColumnClustered) to column|bar|line|pie. */
export function chartAlias(raw: string): ChartAlias | null {
  const key = raw.trim().toLowerCase();
  for (const alias of CHART_ALIASES) {
    if (key === alias || key.startsWith(alias)) {
      return alias;
    }
  }
  return null;
}

export const CHART_ALIAS_LIST: ChartAlias[] = [...CHART_ALIASES];
