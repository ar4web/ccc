"""Create a presentation-focused CCC copy from an existing five-tab workbook.

Usage: python polish_ccc.py /path/to/CCC-V1.xlsx [output.xlsx]
The source is never overwritten. Formula caches are preserved, not recalculated.
"""
from __future__ import annotations

import argparse
import datetime
import posixpath
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formula import Tokenizer
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from copy import copy

REQUIRED_SHEETS = [
    "Dashboard",
    "Main Dataset",
    "Timesheet_Input",
    "Salarysheet",
    "Operations Control",
]
OUTPUT_SHEETS = REQUIRED_SHEETS + ["Action Center"]
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"x": SHEET_NS}
Q = lambda name: f"{{{SHEET_NS}}}{name}"


def formula_cells(book):
    return {
        (ws.title, cell.coordinate): cell.value
        for ws in book
        for cell in ws._cells.values()
        if cell.data_type == "f"
    }


def table_refs(book):
    return {
        (ws.title, name): ws.tables[name].ref
        for ws in book
        for name in ws.tables
    }


def validation_signature(book):
    return {
        ws.title: tuple(
            (str(v.sqref), v.type, v.formula1, v.formula2, v.allow_blank)
            for v in ws.data_validations.dataValidation
        )
        for ws in book
    }


def safe_number(value):
    return float(value) if isinstance(value, (int, float)) else 0.0


def cache_overrides(values, action_inputs=None):
    action_inputs = action_inputs or {}
    dash = values["Dashboard"]
    input_ws = values["Timesheet_Input"]
    live_rows = sum(input_ws[f"BE{row}"].value == 1 for row in range(6, 426))
    expired = int(safe_number(dash["F15"].value))
    due_soon = int(safe_number(dash["J15"].value))
    finding_count = int(safe_number(dash["B29"].value))
    missing_dates = int(safe_number(dash["N29"].value))
    ready = safe_number(values["Salarysheet"]["B947"].value)
    held = safe_number(values["Salarysheet"]["F947"].value)
    metrics = [expired, due_soon, finding_count, live_rows, held, missing_dates]
    is_open = [expired > 0, due_soon > 0, finding_count > 0, live_rows == 0, held > 0, missing_dates > 0]
    open_count = sum(is_open)
    summary = (
        f"{'Live timesheet rows: ' + format(live_rows, ',') if live_rows else 'No live timesheet inputs'}"
        f"  ·  Payroll ready/held SAR {ready:,.2f}/{held:,.2f}"
        f"  ·  {finding_count:,} open findings  ·  Open actions: {open_count}"
    )
    review_date = values["Operations Control"]["B1770"].value
    if not isinstance(review_date, (datetime.datetime, datetime.date)):
        review_date = dash["N12"].value
    review_day = review_date.date() if isinstance(review_date, datetime.datetime) else review_date

    unassigned = 0
    overdue = 0
    for idx, row in enumerate(range(6, 12)):
        if not is_open[idx]:
            continue
        owner = action_inputs.get((row, 5))
        due_date = action_inputs.get((row, 6))
        action_status = action_inputs.get((row, 7), "")
        unassigned += not bool(owner)
        if isinstance(due_date, datetime.datetime):
            due_date = due_date.date()
        overdue += bool(due_date and review_day and due_date < review_day and action_status != "Done")

    action = {
        "B3": ("n", str(open_count)),
        "D3": ("n", str(unassigned)),
        "F3": ("n", str(overdue)),
        "H3": ("str", review_day.strftime("%d %b %Y") if review_day else "Not set"),
    }
    priorities = ["URGENT" if is_open[0] else "CLEAR", "DUE SOON" if is_open[1] else "CLEAR",
                  "URGENT" if is_open[2] else "CLEAR", "INPUT NEEDED" if is_open[3] else "IN PROGRESS",
                  "HOLD" if is_open[4] else "CLEAR", "REVIEW" if is_open[5] else "CLEAR"]
    actions = ["Renew expired Iqamas", "Schedule renewals due within 15 days",
               "Review findings before payroll release", "Load current-period timesheets",
               "Resolve held payroll before release", "Complete missing Iqama dates"]
    clear_actions = ["No expired Iqamas", "No renewals due within 15 days", "No open findings",
                     "Timesheets entered; continue payroll checks", "No held payroll", "No missing Iqama dates"]
    statuses = ["OPEN" if flag else "CLEAR" for flag in is_open]
    link_texts = ["Open employee master", "Open expiry worklist", "Review exceptions",
                  "Open timesheets", "Open payroll", "Open employee master"]
    for idx, row in enumerate(range(6, 12)):
        action[f"A{row}"] = ("str", priorities[idx])
        action[f"B{row}"] = ("str", actions[idx] if is_open[idx] else clear_actions[idx])
        action[f"C{row}"] = ("n", str(metrics[idx]))
        action[f"D{row}"] = ("str", statuses[idx])
        action[f"I{row}"] = ("str", link_texts[idx])

    return {
        "Dashboard": {
            "B17": ("str", summary),
            "N15": ("n", str(finding_count)),
            "N16": ("str", "Open action center →"),
        },
        "Action Center": action,
    }


def sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result = {}
    for sheet in workbook.findall(f".//{{{SHEET_NS}}}sheet"):
        target = targets[sheet.get(f"{{{DOC_REL_NS}}}id")]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.get("name")] = part
    return result


def set_cached_result(cell, result_type: str, value: str) -> None:
    cell.set("t", result_type)
    cached = cell.find(Q("v"))
    if cached is None:
        cached = ET.SubElement(cell, Q("v"))
    cached.text = value


def transfer_formula_caches(source_path: Path, draft_path: Path, output_path: Path, overrides) -> None:
    with zipfile.ZipFile(source_path) as source_zip, zipfile.ZipFile(draft_path) as draft_zip:
        source_parts = sheet_parts(source_zip)
        draft_parts = sheet_parts(draft_zip)
        if not source_parts.keys() <= draft_parts.keys():
            raise AssertionError("A source worksheet was removed while polishing the workbook.")
        output_parts = {}
        for sheet_name, draft_part in draft_parts.items():
            draft_xml = ET.fromstring(draft_zip.read(draft_part))
            if sheet_name in source_parts:
                source_xml = ET.fromstring(source_zip.read(source_parts[sheet_name]))
                source_formulas = {
                    cell.get("r"): cell
                    for cell in source_xml.iter(Q("c"))
                    if cell.find(Q("f")) is not None
                }
            else:
                source_formulas = {}
            sheet_overrides = overrides.get(sheet_name, {})
            seen_overrides = set()
            for cell in draft_xml.iter(Q("c")):
                if cell.find(Q("f")) is None:
                    continue
                coordinate = cell.get("r")
                old = source_formulas.get(coordinate)
                if old is None:
                    if coordinate in sheet_overrides:
                        result_type, result_value = sheet_overrides[coordinate]
                        set_cached_result(cell, result_type, result_value)
                        seen_overrides.add(coordinate)
                        continue
                    raise AssertionError(f"Formula lacks a source cache: {sheet_name}!{coordinate}")
                new_formula = cell.find(Q("f"))
                old_formula = old.find(Q("f"))
                # Source shared-formula followers have no formula text; openpyxl
                # expands them on save, while each original result cache remains valid.
                if old_formula.text is not None and new_formula.text != old_formula.text:
                    if coordinate not in sheet_overrides:
                        raise AssertionError(f"Unexpected formula change: {sheet_name}!{coordinate}")
                if coordinate in sheet_overrides:
                    result_type, result_value = sheet_overrides[coordinate]
                    set_cached_result(cell, result_type, result_value)
                    seen_overrides.add(coordinate)
                    continue
                if old.get("t") is None:
                    cell.attrib.pop("t", None)
                else:
                    cell.set("t", old.get("t"))
                old_value = old.find(Q("v"))
                new_value = cell.find(Q("v"))
                if old_value is None:
                    if new_value is not None:
                        cell.remove(new_value)
                else:
                    if new_value is None:
                        new_value = ET.SubElement(cell, Q("v"))
                    new_value.text = old_value.text
            if seen_overrides != set(sheet_overrides):
                raise AssertionError(
                    f"Missing cache overrides on {sheet_name}: "
                    f"{set(sheet_overrides) - seen_overrides}"
                )
            output_parts[draft_part] = ET.tostring(
                draft_xml, encoding="UTF-8", xml_declaration=True
            )

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as out_zip:
            for info in draft_zip.infolist():
                data = output_parts.get(info.filename)
                if data is None:
                    data = draft_zip.read(info.filename)
                out_zip.writestr(info, data)


def load_action_inputs(path: Path):
    if not path.is_file():
        return {}
    try:
        previous = openpyxl.load_workbook(path, data_only=False)
    except Exception:
        return {}
    try:
        if "Action Center" not in previous.sheetnames:
            return {}
        ws = previous["Action Center"]
        return {
            (row, col): ws.cell(row, col).value
            for row in range(6, 12)
            for col in range(5, 9)
            if ws.cell(row, col).value not in (None, "")
            and ws.cell(row, col).data_type != "f"
        }
    finally:
        previous.close()


def add_action_center(book, action_inputs=None):
    action_inputs = action_inputs or {}
    ws = book.create_sheet("Action Center")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = "0F6B5A"
    widths = {"A": 17, "B": 44, "C": 14, "D": 14, "E": 18, "F": 16, "G": 18, "H": 32, "I": 22}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    dark = "163239"
    white = "FFFFFF"
    light = "F3F7F5"
    input_fill = "EDF5FA"
    dark_font = Font(name="Calibri", size=10, color="193940")
    white_header = Font(name="Calibri", size=11, bold=True, color=white)
    title_font = Font(name="Calibri", size=18, bold=True, color=white)
    hint_font = Font(name="Calibri", size=10, color="30474B")
    label_font = Font(name="Calibri", size=10, bold=True, color="30474B")
    metric_font = Font(name="Calibri", size=16, bold=True, color="176244")
    thin = Side(style="thin", color="D8E1E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:I1")
    ws["A1"] = "CCC  /  ACTION CENTER"
    ws["A1"].fill = PatternFill("solid", fgColor=dark)
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A2:I2")
    ws["A2"] = "Live alerts follow Dashboard filters. Assign an owner, follow-up date, action status, and remark in the blue cells. Fix the source issue to clear its alert."
    ws["A2"].fill = PatternFill("solid", fgColor=light)
    ws["A2"].font = hint_font
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[2].height = 32

    summary_cells = [
        ("A3", "Open alerts"), ("B3", '=COUNTIF($D$6:$D$11,"OPEN")'),
        ("C3", "Unassigned"), ("D3", '=COUNTIFS($D$6:$D$11,"OPEN",$E$6:$E$11,"")'),
        ("E3", "Overdue follow-ups"),
        ("F3", '=COUNTIFS($D$6:$D$11,"OPEN",$F$6:$F$11,">0",$F$6:$F$11,"<"&ReviewDate,$G$6:$G$11,"<>Done")'),
        ("G3", "As of"), ("H3", '=TEXT(ReviewDate,"dd mmm yyyy")'),
    ]
    for coord, value in summary_cells:
        cell = ws[coord]
        cell.value = value
        cell.fill = PatternFill("solid", fgColor="E2F3EB" if coord in ("B3", "D3", "F3", "H3") else "DDECF3")
        cell.font = metric_font if coord in ("B3", "D3", "F3", "H3") else label_font
        cell.alignment = Alignment(horizontal="center" if coord in ("B3", "D3", "F3", "H3") else "left", vertical="center", indent=0 if coord in ("B3", "D3", "F3", "H3") else 1)
        cell.border = border
    ws.row_dimensions[3].height = 30

    headers = ["Priority", "Suggested action", "Metric", "Alert", "Owner", "Follow-up by", "Action status", "Remarks", "Source"]
    for col, heading in enumerate(headers, start=1):
        cell = ws.cell(5, col, heading)
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = white_header
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        cell.border = border
        cell.comment = Comment(
            {
                5: "Set the person responsible for this follow-up.",
                6: "Enter the target date. Overdue follow-ups are counted against the Dashboard reference date.",
                7: "Choose Not started, In progress, or Done. The system alert clears when its source metric is resolved.",
                8: "Add a short context note or next step.",
            }.get(col, "Formula-driven alert; use the source link to review the underlying data."),
            "CCC",
        )
    ws.row_dimensions[5].height = 32

    formula_sets = {
        "A": [
            '=IF(C6=0,"CLEAR","URGENT")',
            '=IF(C7=0,"CLEAR","DUE SOON")',
            '=IF(C8=0,"CLEAR","URGENT")',
            '=IF(C9=0,"INPUT NEEDED","IN PROGRESS")',
            '=IF(C10=0,"CLEAR","HOLD")',
            '=IF(C11=0,"CLEAR","REVIEW")',
        ],
        "B": [
            '=IF(C6=0,"No expired Iqamas","Renew expired Iqamas")',
            '=IF(C7=0,"No renewals due within 15 days","Schedule renewals due within 15 days")',
            '=IF(C8=0,"No open findings","Review findings before payroll release")',
            '=IF(C9=0,"Load current-period timesheets","Timesheets entered; continue payroll checks")',
            '=IF(C10=0,"No held payroll","Resolve held payroll before release")',
            '=IF(C11=0,"No missing Iqama dates","Complete missing Iqama dates")',
        ],
        "C": [
            "=Dashboard!$F$15", "=Dashboard!$J$15", "=Dashboard!$N$15",
            "=COUNTIF(Timesheet_Input!$BE$6:$BE$425,1)", "=Salarysheet!$F$947", "=Dashboard!$N$29",
        ],
        "D": [
            '=IF(C6=0,"CLEAR","OPEN")', '=IF(C7=0,"CLEAR","OPEN")',
            '=IF(C8=0,"CLEAR","OPEN")', '=IF(C9=0,"OPEN","CLEAR")',
            '=IF(C10=0,"CLEAR","OPEN")', '=IF(C11=0,"CLEAR","OPEN")',
        ],
        "I": [
            '=HYPERLINK("#\'Main Dataset\'!A5","Open employee master")',
            '=HYPERLINK("#\'Dashboard\'!B64","Open expiry worklist")',
            '=HYPERLINK("#\'Operations Control\'!A69","Review exceptions")',
            '=HYPERLINK("#\'Timesheet_Input\'!A5","Open timesheets")',
            '=HYPERLINK("#\'Salarysheet\'!A5","Open payroll")',
            '=HYPERLINK("#\'Main Dataset\'!D5","Open employee master")',
        ],
    }
    for idx, row in enumerate(range(6, 12)):
        for col, formulas in formula_sets.items():
            cell = ws[f"{col}{row}"]
            cell.value = formulas[idx]
            cell.font = dark_font
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws[f"C{row}"].number_format = '"SAR "#,##0.00' if row == 10 else "#,##0"
        for col in ("E", "F", "G", "H"):
            cell = ws[f"{col}{row}"]
            cell.value = action_inputs.get((row, openpyxl.utils.column_index_from_string(col)))
            cell.fill = PatternFill("solid", fgColor=input_fill)
            cell.font = dark_font
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws[f"F{row}"].number_format = "dd mmm yyyy"
        ws.row_dimensions[row].height = 34
        for col in range(1, 10):
            ws.cell(row, col).border = border

    dv_status = DataValidation(type="list", formula1='"Not started,In progress,Done"', allow_blank=True)
    dv_status.promptTitle = "Action status"
    dv_status.prompt = "Choose the current follow-up status."
    dv_status.errorTitle = "Choose a listed status"
    dv_status.error = "Use Not started, In progress, or Done."
    ws.add_data_validation(dv_status)
    dv_status.add("G6:G11")
    dv_date = DataValidation(type="date", operator="greaterThan", formula1="0", allow_blank=True)
    dv_date.promptTitle = "Follow-up date"
    dv_date.prompt = "Enter the target date for this action."
    ws.add_data_validation(dv_date)
    dv_date.add("F6:F11")

    red_fill = PatternFill("solid", fgColor="FBE7E1")
    amber_fill = PatternFill("solid", fgColor="FFF3C4")
    green_fill = PatternFill("solid", fgColor="E2F3EB")
    red_font = Font(name="Calibri", size=10, bold=True, color="A62C35")
    amber_font = Font(name="Calibri", size=10, bold=True, color="856000")
    green_font = Font(name="Calibri", size=10, bold=True, color="176244")
    ws.conditional_formatting.add("A6:A11", FormulaRule(formula=['$A6="URGENT"'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add("A6:A11", FormulaRule(formula=['$A6="HOLD"'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add("A6:A11", FormulaRule(formula=['$A6="DUE SOON"'], fill=amber_fill, font=amber_font))
    ws.conditional_formatting.add("A6:A11", FormulaRule(formula=['$A6="INPUT NEEDED"'], fill=amber_fill, font=amber_font))
    ws.conditional_formatting.add("D6:D11", FormulaRule(formula=['$D6="OPEN"'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add("D6:D11", FormulaRule(formula=['$D6="CLEAR"'], fill=green_fill, font=green_font))
    ws.conditional_formatting.add("F6:F11", FormulaRule(formula=['AND(ISNUMBER($F6),$F6<ReviewDate,$D6="OPEN",$G6<>"Done")'], fill=red_fill, font=red_font))

    table = Table(displayName="CCCActionTracker", ref="A5:I11")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = "A5:I11"
    ws.print_area = "A1:I11"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def polish(source: Path, output: Path) -> dict[str, int]:
    if not source.is_file():
        raise FileNotFoundError(f"Workbook not found: {source}")
    if source.resolve() == output.resolve():
        raise ValueError("Choose a separate output path; the source workbook is left unchanged.")

    action_inputs = load_action_inputs(output)
    original = openpyxl.load_workbook(source, data_only=False)
    source_values = openpyxl.load_workbook(source, data_only=True)
    if original.sheetnames != REQUIRED_SHEETS:
        raise ValueError(
            "Expected a built CCC workbook with these five tabs: "
            + ", ".join(REQUIRED_SHEETS)
        )
    if original._external_links:
        raise ValueError("External workbook links are present; review them before delivery.")

    overrides = cache_overrides(source_values, action_inputs)
    formulas_before = formula_cells(original)
    formula_cache_before = {
        (ws.title, cell.coordinate): source_values[ws.title][cell.coordinate].value
        for ws in original
        for cell in ws._cells.values()
        if cell.data_type == "f"
    }

    dashboard = original["Dashboard"]
    dashboard["B2"] = "CCC  /  COMPANY COMMAND CENTER"
    dashboard["B3"] = "Workforce & payroll  /  SYNTHETIC DATA"
    dashboard["N14"] = "Open findings"
    dashboard["N15"] = "=B29"
    dashboard["N16"] = '=HYPERLINK("#\'Action Center\'!A1","Open action center →")'
    dashboard["B17"] = (
        '=IF(COUNTIF(Timesheet_Input!$BE$6:$BE$425,1)=0,"No live timesheet inputs",'
        '"Live timesheet rows: "&COUNTIF(Timesheet_Input!$BE$6:$BE$425,1))'
        '&"  ·  Payroll ready/held SAR "'
        '&TEXT(Salarysheet!$B$947,"#,##0.00")&"/"'
        '&TEXT(Salarysheet!$F$947,"#,##0.00")'
        '&"  ·  "&N15&" open findings  ·  Open actions: "&\'Action Center\'!$B$3'
    )
    action_ws = add_action_center(original, action_inputs)
    tables_before = table_refs(original)
    validations_before = validation_signature(original)

    if any(r.min_row <= 10 <= r.max_row for r in dashboard.merged_cells.ranges):
        raise ValueError("Unexpected merged cells in Dashboard row 10.")
    if any(dashboard.cell(10, col).value is not None for col in range(2, 17)):
        raise ValueError("Dashboard row 10 is not empty; refusing to overwrite it.")
    dashboard.merge_cells("B10:P10")
    cue = dashboard["B10"]
    cue.value = "ADVANCED FILTERS  /  Expand the grouped rows below for nationality, contract and ranking controls."
    cue._style = copy(dashboard["B8"]._style)
    cue.font = Font(name="Calibri", size=10, color="516C70")
    cue.fill = PatternFill("solid", fgColor="F3F7F5")
    cue.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    dashboard.row_dimensions[10].height = 19
    dashboard.row_dimensions.group(11, 12, outline_level=1, hidden=True)
    dashboard.sheet_properties.outlinePr.summaryBelow = False

    # Keep the headline, KPI cards, a compact payroll/input strip, and the
    # renewal worklist visible. Detailed exception and ranking tables remain
    # in the workbook and can be expanded from the outline controls.
    dashboard.row_dimensions.group(24, 62, outline_level=1, hidden=True)
    dashboard.row_dimensions.group(78, 142, outline_level=1, hidden=True)
    dashboard._charts = []
    dashboard.print_area = "A1:Q77"
    dashboard.page_setup.orientation = "landscape"
    dashboard.page_setup.fitToWidth = 1
    dashboard.page_setup.fitToHeight = 1
    dashboard.sheet_properties.pageSetUpPr.fitToPage = True
    dashboard.row_breaks.brk = []
    dashboard.sheet_view.zoomScale = 95
    dashboard.freeze_panes = "B8"
    dashboard.oddFooter.left.text = "CCC  /  SYNTHETIC DEMO DATA"

    for sheet in original:
        sheet.sheet_view.tabSelected = sheet.title == "Dashboard"
    original.active = original.sheetnames.index("Dashboard")
    original.properties.title = "CCC | Company Command Center"
    original.properties.subject = "Workforce and payroll operations (synthetic data)"
    original.properties.description = "Focused workforce, payroll, compliance and review dashboard."
    original.calculation.calcMode = "auto"
    # Don't force a full 117k-formula recalculation on every open. Excel will
    # still calculate changed formulas automatically, and may calculate once
    # on first open because many source formulas have no saved result cache.
    original.calculation.fullCalcOnLoad = False
    original.calculation.forceFullCalc = False
    original.calculation.calcOnSave = True

    for coordinate in ("B17", "N15", "N16"):
        value = dashboard[coordinate].value
        if not isinstance(value, str) or not value.startswith("="):
            raise AssertionError(f"Expected a formula in Dashboard!{coordinate}")
        Tokenizer(value)

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ccc-polish-") as temp_dir:
        draft = Path(temp_dir) / "polished.xlsx"
        original.save(draft)
        transfer_formula_caches(source, draft, output, overrides)

    polished = openpyxl.load_workbook(output, data_only=False)
    polished_values = openpyxl.load_workbook(output, data_only=True)
    assert polished.sheetnames == OUTPUT_SHEETS
    assert polished.active.title == "Dashboard"
    assert polished.properties.title == "CCC | Company Command Center"
    assert len(polished["Dashboard"]._charts) == 0
    assert len(polished["Salarysheet"]._charts) == 3
    assert table_refs(polished) == tables_before
    assert validation_signature(polished) == validations_before
    expected_formulas = formulas_before.copy()
    expected_formulas["Dashboard", "B17"] = dashboard["B17"].value
    expected_formulas["Dashboard", "N15"] = "=B29"
    expected_formulas["Dashboard", "N16"] = dashboard["N16"].value
    expected_formulas.update({
        ("Action Center", cell.coordinate): cell.value
        for cell in action_ws._cells.values()
        if cell.data_type == "f"
    })
    assert formula_cells(polished) == expected_formulas
    assert not polished._external_links
    assert polished.calculation.calcMode == "auto"
    assert not polished.calculation.fullCalcOnLoad
    assert not polished.calculation.forceFullCalc

    # Ensure all source formula results remain intact except the two visible
    # dashboard summaries whose cached strings/numbers are intentionally updated.
    formula_cache_after = {
        (ws.title, cell.coordinate): polished_values[ws.title][cell.coordinate].value
        for ws in polished
        for cell in ws._cells.values()
        if cell.data_type == "f"
    }
    for key, value in formula_cache_before.items():
        if key in (("Dashboard", "B17"), ("Dashboard", "N15")):
            continue
        assert formula_cache_after[key] == value, (key, value, formula_cache_after[key])
    for sheet_name, cells in overrides.items():
        for coordinate, (result_type, expected) in cells.items():
            actual = polished_values[sheet_name][coordinate].value
            if result_type == "n":
                assert float(actual) == float(expected), (sheet_name, coordinate, actual, expected)
            else:
                assert actual == expected, (sheet_name, coordinate, actual, expected)

    formula_errors = [
        (ws.title, cell.coordinate)
        for ws in polished_values
        for cell in ws._cells.values()
        if cell.data_type == "e"
    ]
    assert not formula_errors, formula_errors[:10]
    for start, end in ((11, 12), (24, 62), (78, 142)):
        assert all(polished["Dashboard"].row_dimensions[r].hidden for r in range(start, end + 1))

    result = {
        "formula_count": len(formula_cache_after),
        "formula_cache_errors": len(formula_errors),
        "tables": sum(len(ws.tables) for ws in polished),
        "dashboard_charts": len(polished["Dashboard"]._charts),
        "salary_charts": len(polished["Salarysheet"]._charts),
    }
    original.close()
    source_values.close()
    polished.close()
    polished_values.close()
    return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a focused CCC dashboard copy without changing source data."
    )
    parser.add_argument("source", type=Path, help="Existing five-tab CCC workbook")
    parser.add_argument(
        "output", nargs="?", type=Path, default=Path("CCC_Company_Command_Center.xlsx")
    )
    args = parser.parse_args()
    result = polish(args.source.expanduser(), args.output.expanduser())
    print(f"Created {args.output}; formula caches preserved (not recalculated).")
    print("Audit:", result)
    print("Excel automatic calculation is enabled without forcing a full recalculation on every open.")
    print("The first Excel open may still calculate formulas with missing saved results; save once after it finishes.")


if __name__ == "__main__":
    main()
