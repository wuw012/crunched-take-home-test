from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    workbook = Workbook()

    pnl = workbook.active
    pnl.title = "P&L"
    pnl["A1"] = "Line"
    pnl["B1"] = "FY24"
    pnl["A2"] = "Revenue"
    pnl["B2"] = 1000  # hardcoded; should be Assumptions!B2*B3 (12*100 = 1200)
    pnl["A3"] = "COGS"
    pnl["B3"] = 400
    pnl["A4"] = "Gross Profit"
    pnl["B4"] = 500  # does not foot: 1000-400 = 600
    pnl["A5"] = "Opex"
    pnl["B5"] = 150
    pnl["A6"] = "Operating Profit"
    pnl["B6"] = 350  # follows the wrong GP (500-150)
    pnl["A8"] = "Notes"
    pnl["B8"] = "Revenue is not linked to Assumptions. GP/OP are hardcoded."
    pnl["D1"] = "FY25"
    pnl["E1"] = "FY26"
    pnl["F1"] = "FY27"
    # D2:F6 left empty on purpose — forecast scenario writes formulas here

    assumptions = workbook.create_sheet("Assumptions")
    assumptions["A1"] = "Driver"
    assumptions["B1"] = "Value"
    assumptions["A2"] = "Price"
    assumptions["B2"] = 12
    assumptions["A3"] = "Units"
    assumptions["B3"] = 100
    assumptions["A4"] = "YoY growth"
    assumptions["B4"] = 0.05
    assumptions["A6"] = "Note"
    assumptions["B6"] = "Price * Units should drive P&L Revenue."

    data = workbook.create_sheet("Exports")
    data["A1"] = "Month"
    data["B1"] = "Region"
    data["C1"] = "Amount"
    regions = ["Nordics", "UK", "DACH", "US"]
    for i in range(2, 82):
        data[f"A{i}"] = f"2024-{((i - 2) % 12) + 1:02d}"
        data[f"B{i}"] = regions[(i - 2) % 4]
        data[f"C{i}"] = 80 + ((i * 17) % 40)

    out = Path(__file__).resolve().parents[2] / "samples" / "demo.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
