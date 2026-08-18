from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator

from app.limits import MAX_CELLS


class list_workbook_structure(BaseModel):
    """List worksheets, used-range dimensions, and existing charts (name, type, title). Includes charts the user created in Excel. Never return cell values."""


class get_selection(BaseModel):
    """Read the current Excel selection as values (computed) and formulas. A cell is hardcoded when its formulas entry does not start with =. If it is larger than 2000 cells, return dimensions only."""


class read_range(BaseModel):
    """Read values (computed) and formulas from an A1 range. A cell is hardcoded when its formulas entry does not start with =. If the range is larger than 2000 cells, do not read cells — return dimensions so you can retry a smaller slice."""

    sheet: str = Field(description="Worksheet name")
    a1: str = Field(description="A1 notation such as A1:D20")


class write_range(BaseModel):
    """Write a 2D array starting at start_cell. Strings that start with = are Excel formulas, not hardcoded values. Prefer formulas for totals and calculated lines. If the block is larger than 2000 cells, do not write — use a smaller range."""

    sheet: str = Field(description="Worksheet name")
    start_cell: str = Field(description="Top-left cell, e.g. B5")
    values: list[list[Any]] = Field(description="Rows of cell values or formulas")

    @model_validator(mode="after")
    def cap_cells(self) -> "write_range":
        rows = len(self.values)
        columns = max((len(row) for row in self.values), default=0)
        if rows == 0 or columns == 0:
            raise ValueError("values must be a non-empty 2D array")
        if rows * columns > MAX_CELLS:
            raise ValueError(
                f"Range too large: {rows}x{columns} cells. Max is {MAX_CELLS}. Write a smaller block."
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
