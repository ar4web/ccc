"""Create a manual-recalc full workbook and an instant, static Quick View.

The automatic full workbook is created separately by polish_ccc.py. The manual
copy keeps every formula and saved result; the Quick View uses cached sample
values only and intentionally contains no formulas.
"""
from __future__ import annotations

import argparse
from copy import copy
from datetime import date, datetime
from pathlib import Path
import tempfile

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from polish_ccc import OUTPUT_SHEETS, formula_cells, transfer_formula_caches

DARK = "163239"
DARK_TEXT = "193940"
WHITE = "FFFFFF"
MINT = "D6FFF0"
GREEN = "E2F3EB"
GREEN_TEXT = "176244"
RED = "FBE7E1"
RED_TEXT = "A62C35"
AMBER = "FFF3C4"
AMBER_TEXT = "856000"
PALE = "F3F7F5"
PALE_BLUE = "EBF1F4"
LINE = "D8E1E0"


def safe_number(value) -> float:
    return float(value) if isinstance(value, (int, float)) else 0.0


def date_text(value) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%d %b %Y")
    return "Not set"


def make_manual_copy(source: Path, output: Path) -> None:
    book = openpyxl.load_workbook(source, data_only=False)
    if book.sheetnames != OUTPUT_SHEETS:
        raise ValueError("Manual copy expects the complete CCC workbook with its Action Center.")
    dash = book["Dashboard"]
    dash["B3"] = "Workforce & payroll  /  SYNTHETIC DATA  /  MANUAL CALC — PRESS F9"
    font = copy(dash["B3"].font)
    font.sz = 15
    dash["B3"].font = font
    dash["B3"].alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
    book.calculation.calcMode = "manual"
    book.calculation.fullCalcOnLoad = False
    book.calculation.forceFullCalc = False
    book.calculation.calcOnSave = False
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ccc-manual-") as temp_dir:
        draft = Path(temp_dir) / "manual.xlsx"
        book.save(draft)
        transfer_formula_caches(source, draft, output, overrides={})
    book.close()

    check = openpyxl.load_workbook(output, data_only=False)
    assert check.calculation.calcMode == "manual"
    assert not check.calculation.fullCalcOnLoad and not check.calculation.forceFullCalc
    assert len(formula_cells(check)) > 100_000
    check.close()


def style_merged(ws, ref: str, fill: str, font: Font, alignment: Alignment) -> None:
    ws.merge_cells(ref)
    for row in ws[ref]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = copy(font)
            cell.alignment = copy(alignment)


def quick_view(source: Path, output: Path) -> None:
    cached = openpyxl.load_workbook(source, data_only=True)
    dash = cached["Dashboard"]
    salary = cached["Salarysheet"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CCC Quick View"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    for col, width in {
        "A": 19, "B": 17, "C": 16, "D": 14,
        "E": 16, "F": 12, "G": 16, "H": 16,
    }.items():
        ws.column_dimensions[col].width = width

    title_font = Font(name="Calibri", size=20, bold=True, color=WHITE)
    subtitle_font = Font(name="Calibri", size=10, bold=True, color=MINT)
    label_font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
    value_font = Font(name="Calibri", size=24, bold=True, color=DARK_TEXT)
    note_font = Font(name="Calibri", size=10, color=DARK_TEXT)
    header_font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    table_header_font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    body_font = Font(name="Calibri", size=10, color=DARK_TEXT)
    left = Alignment(horizontal="left", vertical="center", indent=1)
    center = Alignment(horizontal="center", vertical="center")
    thin_gray = Side(style="thin", color=LINE)
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    style_merged(ws, "A1:H1", DARK, title_font, left)
    ws["A1"] = "CCC  /  COMPANY COMMAND CENTER"
    style_merged(ws, "A2:H2", DARK, subtitle_font, left)
    ws["A2"] = "SYNTHETIC DATA  ·  QUICK VIEW  ·  STATIC CACHED VALUES — NOT RECALCULATED"
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 24

    ws["A3"] = "Report month"
    ws["B3"] = dash["B9"].value
    ws["E3"] = "Compliance as of"
    ws["F3"] = dash["N12"].value
    for coord in ("A3", "E3"):
        ws[coord].font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
    for coord in ("B3", "F3"):
        ws[coord].font = Font(name="Calibri", size=10, color=DARK_TEXT)
        ws[coord].number_format = "mmmm yyyy" if coord == "B3" else "dd mmm yyyy"
    ws.merge_cells("B3:D3")
    ws.merge_cells("F3:H3")
    ws.row_dimensions[3].height = 23

    cards = [
        ("A", "B", "Active employees", dash["B15"].value, "Roster in scope", GREEN, GREEN_TEXT),
        ("C", "D", "Expired Iqamas", dash["F15"].value, "ACTION REQUIRED", RED, RED_TEXT),
        ("E", "F", "Due within 15 days", dash["J15"].value, "DUE SOON · includes today", AMBER, AMBER_TEXT),
        ("G", "H", "Open findings", dash["N15"].value, "Review before release", RED, RED_TEXT),
    ]
    for start, end, label, value, note, fill, color in cards:
        style_merged(ws, f"{start}5:{end}5", fill, label_font, left)
        ws[f"{start}5"] = label
        style_merged(
            ws,
            f"{start}6:{end}7",
            fill,
            Font(name="Calibri", size=25, bold=True, color=color),
            center,
        )
        ws[f"{start}6"] = value if value is not None else "—"
        style_merged(ws, f"{start}8:{end}8", fill, Font(name="Calibri", size=9, bold=True, color=color), center)
        ws[f"{start}8"] = note
    for row, height in ((5, 24), (6, 30), (7, 16), (8, 23)):
        ws.row_dimensions[row].height = height

    style_merged(
        ws,
        "A10:H10",
        PALE,
        Font(name="Calibri", size=10, bold=True, color=DARK_TEXT),
        left,
    )
    ws["A10"] = "STATUS KEY   ·   GREEN = ACTIVE   /   AMBER = DUE SOON   /   RED = ACTION OR REVIEW REQUIRED"
    ws.row_dimensions[10].height = 23

    style_merged(ws, "A12:H12", DARK, header_font, left)
    ws["A12"] = "PAYROLL SNAPSHOT"
    ready = safe_number(salary["B947"].value)
    held = safe_number(salary["F947"].value)
    style_merged(ws, "A13:D13", GREEN, label_font, left)
    ws["A13"] = "Ready estimate"
    style_merged(ws, "E13:H13", RED, label_font, left)
    ws["E13"] = "Held estimate"
    style_merged(ws, "A14:D14", GREEN, Font(name="Calibri", size=17, bold=True, color=GREEN_TEXT), center)
    ws["A14"] = f"SAR {ready:,.2f}"
    style_merged(ws, "E14:H14", RED, Font(name="Calibri", size=17, bold=True, color=RED_TEXT), center)
    ws["E14"] = f"SAR {held:,.2f}"
    ws.row_dimensions[12].height = 25
    ws.row_dimensions[13].height = 24
    ws.row_dimensions[14].height = 31

    status = str(dash["B17"].value or "No status summary saved.")
    style_merged(ws, "A16:H16", PALE_BLUE, note_font, left)
    ws["A16"] = status
    ws["A16"].alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[16].height = 30

    action_kpis = str(dash["B22"].value or "Action KPIs not available in saved values.")
    style_merged(
        ws,
        "A17:H17",
        GREEN,
        Font(name="Calibri", size=9, bold=True, color=DARK_TEXT),
        left,
    )
    ws["A17"] = f"ACTION KPIs  ·  {action_kpis}"
    ws["A17"].alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[17].height = 27

    style_merged(ws, "A18:H18", DARK, header_font, left)
    ws["A18"] = "RENEWAL WORKLIST  /  SAVED SAMPLE ORDER"
    headers = [("A19", "Employee"), ("C19", "Client"), ("E19", "Expiry date"), ("F19", "Days left"), ("G19", "Next step")]
    for coord, value in headers:
        ws[coord] = value
        ws[coord].font = table_header_font
        ws[coord].fill = PatternFill("solid", fgColor="24505A")
        ws[coord].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws[coord].border = thin_border
    ws.merge_cells("A19:B19")
    ws.merge_cells("C19:D19")
    ws.merge_cells("G19:H19")
    ws.row_dimensions[18].height = 25
    ws.row_dimensions[19].height = 23

    entries = []
    for row in range(67, 77):
        person = dash[f"B{row}"].value
        if person in (None, ""):
            continue
        entries.append((person, dash[f"F{row}"].value, dash[f"J{row}"].value, dash[f"L{row}"].value, dash[f"N{row}"].value))
    for index, (person, client, expiry, days_left, next_step) in enumerate(entries, start=20):
        fill = WHITE if index % 2 == 0 else PALE
        for ref in (f"A{index}:B{index}", f"C{index}:D{index}", f"G{index}:H{index}"):
            ws.merge_cells(ref)
        vals = {f"A{index}": person, f"C{index}": client or "Not recorded", f"E{index}": expiry, f"F{index}": days_left, f"G{index}": next_step}
        for coord, value in vals.items():
            cell = ws[coord]
            cell.value = value
            cell.font = body_font
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            cell.border = thin_border
        ws[f"E{index}"].number_format = "dd mmm yyyy"
        ws[f"F{index}"].number_format = "0"
        if isinstance(days_left, (int, float)) and days_left <= 15:
            ws[f"F{index}"].font = Font(name="Calibri", size=10, bold=True, color=RED_TEXT if days_left < 0 else AMBER_TEXT)
        if next_step and ("overdue" in str(next_step).lower() or "renew next" in str(next_step).lower()):
            ws[f"G{index}"].font = Font(name="Calibri", size=10, bold=True, color=RED_TEXT if "overdue" in str(next_step).lower() else AMBER_TEXT)
        ws.row_dimensions[index].height = 22

    end_row = max(20, 19 + len(entries))
    detail = dash["B77"].value
    note_row = end_row + 2
    style_merged(ws, f"A{note_row}:H{note_row}", PALE, note_font, left)
    ws[f"A{note_row}"] = str(detail or "Renewal list reflects saved source values.")
    ws.row_dimensions[note_row].height = 22
    foot_row = note_row + 1
    style_merged(
        ws,
        f"A{foot_row}:H{foot_row}",
        PALE_BLUE,
        Font(name="Calibri", size=9, italic=True, color=DARK_TEXT),
        left,
    )
    ws[f"A{foot_row}"] = (
        "STATIC SNAPSHOT: values come from saved formula caches, not a fresh calculation. "
        "Use the full CCC workbook for data entry and formula updates."
    )
    ws.row_dimensions[foot_row].height = 24

    for row in ws.iter_rows(min_row=1, max_row=foot_row, min_col=1, max_col=8):
        for cell in row:
            if cell.value is not None and cell.border == Border():
                cell.border = thin_border
    ws.freeze_panes = "A20"
    ws.print_area = f"A1:H{foot_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.properties.title = "CCC Quick View | Synthetic cached snapshot"
    wb.properties.subject = "Static cached sample dashboard"
    wb.calculation.calcMode = "manual"
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.forceFullCalc = False
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    wb.close()
    cached.close()

    check = openpyxl.load_workbook(output, data_only=False)
    assert check.sheetnames == ["CCC Quick View"]
    assert sum(c.data_type == "f" for s in check for c in s._cells.values()) == 0
    assert "STATIC CACHED VALUES" in check["CCC Quick View"]["A2"].value
    check.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Make fast/manual CCC download variants from the automatic workbook.")
    parser.add_argument("source", nargs="?", type=Path, default=Path("CCC_Company_Command_Center.xlsx"))
    parser.add_argument("--manual-output", type=Path, default=Path("CCC_Manual_Recalc.xlsx"))
    parser.add_argument("--quick-output", type=Path, default=Path("CCC_Quick_View.xlsx"))
    args = parser.parse_args()
    source = args.source.expanduser()
    if not source.is_file():
        raise SystemExit(f"Automatic CCC workbook not found: {source}")
    make_manual_copy(source, args.manual_output.expanduser())
    quick_view(source, args.quick_output.expanduser())
    print(f"Created manual full workbook: {args.manual_output}")
    print(f"Created static quick view: {args.quick_output}")


if __name__ == "__main__":
    main()
