from typing import Any, Literal
import re

from pydantic import BaseModel, Field, field_validator, model_validator

from app.limits import MAX_CELLS

A1_CELL = re.compile(r"^\$?[A-Za-z]+\$?[1-9]\d*$")


class list_workbook_structure(BaseModel):
    """List worksheets, used-range dimensions, and existing charts (name, type, title). Includes charts the user created in Excel. Never return cell values."""


class get_selection(BaseModel):
    """Read the current Excel selection as values (computed) and formulas. A cell is hardcoded when its formulas entry does not start with =. If it is larger than 2000 cells, return dimensions only."""


class read_range(BaseModel):
    """Read values (computed) and formulas from an A1 range. A cell is hardcoded when its formulas entry does not start with =. If the range is larger than 2000 cells, do not read cells — return dimensions so you can retry a smaller slice."""

    sheet: str = Field(description="Worksheet name")
    a1: str = Field(description="A1 notation such as A1:D20")


class write_range(BaseModel):
    """Write a 2D array starting at start_cell. Strings that start with = are Excel formulas, mixed with values in the same block. Omitted cells are left unchanged. Prefer formulas for totals and calculated lines. If the block is larger than 2000 cells, do not write — use a smaller range."""

    sheet: str = Field(description="Worksheet name")
    start_cell: str = Field(description="Top-left cell only, e.g. B5 — not a range")
    values: list[list[Any]] = Field(description="Rows of cell values or formulas")

    @field_validator("start_cell")
    @classmethod
    def start_cell_is_single(cls, value: str) -> str:
        cell = value.strip()
        if not A1_CELL.match(cell):
            raise ValueError("start_cell must be a single cell like B4, not a range")
        return cell

    @model_validator(mode="after")
    def cap_cells(self) -> "write_range":
        cells = sum(1 for row in self.values for cell in row if cell is not None)
        if cells == 0:
            raise ValueError("values must be a non-empty 2D array")
        if cells > MAX_CELLS:
            raise ValueError(
                f"Range too large: {cells} cells. Max is {MAX_CELLS}. Write a smaller block."
            )
        return self


class create_chart(BaseModel):
    """Create an Excel chart from a source range. Allowed chart_type values: column, bar, line, pie. The chart is a live Excel object, not an image."""

    sheet: str = Field(description="Worksheet name")
    source_a1: str = Field(description="Data range in A1 notation, e.g. A1:B6")
    chart_type: Literal["column", "bar", "line", "pie"] = Field(
        description="column, bar, line, or pie"
    )
    title: str | None = Field(default=None, description="Optional chart title")


class set_chart_type(BaseModel):
    """Change an existing chart's type, including charts the user created. Allowed: column, bar, line, pie. Prefer this over create_chart when list_workbook_structure already shows a chart. Pass name from that list; if omitted, updates the last chart on the sheet. Do not screenshot the chart."""

    sheet: str = Field(description="Worksheet name")
    chart_type: Literal["column", "bar", "line", "pie"] = Field(
        description="column, bar, line, or pie"
    )
    name: str | None = Field(default=None, description="Chart name from list_workbook_structure")


TOOLS = [
    list_workbook_structure,
    get_selection,
    read_range,
    write_range,
    create_chart,
    set_chart_type,
]
TOOLS_BY_NAME = {cls.__name__: cls for cls in TOOLS}
