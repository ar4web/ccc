"""Create a presentation-focused CCC copy from an existing five-tab workbook.

Usage: python polish_ccc.py /path/to/CCC-V1.xlsx [output.xlsx]
The source is never overwritten. Formula caches are preserved, not recalculated.
"""
from __future__ import annotations

import argparse
import posixpath
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.styles import Alignment, Font, PatternFill
from copy import copy

REQUIRED_SHEETS = [
    "Dashboard",
    "Main Dataset",
    "Timesheet_Input",
    "Salarysheet",
    "Operations Control",
]
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


def cache_overrides(values):
    dash = values["Dashboard"]
    input_ws = values["Timesheet_Input"]
    live_rows = sum(input_ws[f"BN{row}"].value == 1 for row in range(6, 426))
    finding_count = int(safe_number(dash["B29"].value))
    ready = safe_number(values["Salarysheet"]["B947"].value)
    held = safe_number(values["Salarysheet"]["F947"].value)

    if live_rows:
        entry_text = f"Live timesheet rows: {live_rows:,}"
    else:
        entry_text = "No live timesheet inputs"
    summary = (
        f"{entry_text}  ·  Payroll ready/held SAR "
        f"{ready:,.2f}/{held:,.2f}  ·  {finding_count:,} open findings"
    )
    return {
        "Dashboard": {
            "B17": ("str", summary),
            "N15": ("n", str(finding_count)),
        }
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
        if source_parts.keys() != draft_parts.keys():
            raise AssertionError("Worksheet names changed while polishing the workbook.")
        output_parts = {}
        for sheet_name, draft_part in draft_parts.items():
            draft_xml = ET.fromstring(draft_zip.read(draft_part))
            source_xml = ET.fromstring(source_zip.read(source_parts[sheet_name]))
            source_formulas = {
                cell.get("r"): cell
                for cell in source_xml.iter(Q("c"))
                if cell.find(Q("f")) is not None
            }
            sheet_overrides = overrides.get(sheet_name, {})
            seen_overrides = set()
            for cell in draft_xml.iter(Q("c")):
                if cell.find(Q("f")) is None:
                    continue
                coordinate = cell.get("r")
                old = source_formulas.get(coordinate)
                if old is None:
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


def polish(source: Path, output: Path) -> dict[str, int]:
    if not source.is_file():
        raise FileNotFoundError(f"Workbook not found: {source}")
    if source.resolve() == output.resolve():
        raise ValueError("Choose a separate output path; the source workbook is left unchanged.")

    original = openpyxl.load_workbook(source, data_only=False)
    source_values = openpyxl.load_workbook(source, data_only=True)
    if original.sheetnames != REQUIRED_SHEETS:
        raise ValueError(
            "Expected a built CCC workbook with these five tabs: "
            + ", ".join(REQUIRED_SHEETS)
        )
    if original._external_links:
        raise ValueError("External workbook links are present; review them before delivery.")

    overrides = cache_overrides(source_values)
    formulas_before = formula_cells(original)
    tables_before = table_refs(original)
    validations_before = validation_signature(original)
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
    dashboard["N16"] = "Review before payroll release"
    dashboard["B17"] = (
        '=IF(COUNTIF(Timesheet_Input!$BN$6:$BN$425,1)=0,"No live timesheet inputs",'
        '"Live timesheet rows: "&COUNTIF(Timesheet_Input!$BN$6:$BN$425,1))'
        '&"  ·  Payroll ready/held SAR "'
        '&TEXT(Salarysheet!$B$947,"#,##0.00")&"/"'
        '&TEXT(Salarysheet!$F$947,"#,##0.00")'
        '&"  ·  "&N15&" open findings"'
    )

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

    for coordinate in ("B17", "N15"):
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
    assert polished.sheetnames == REQUIRED_SHEETS
    assert polished.active.title == "Dashboard"
    assert polished.properties.title == "CCC | Company Command Center"
    assert len(polished["Dashboard"]._charts) == 0
    assert len(polished["Salarysheet"]._charts) == 3
    assert table_refs(polished) == tables_before
    assert validation_signature(polished) == validations_before
    expected_formulas = formulas_before.copy()
    expected_formulas["Dashboard", "B17"] = dashboard["B17"].value
    expected_formulas["Dashboard", "N15"] = "=B29"
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
                assert actual == int(expected), (sheet_name, coordinate, actual, expected)
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
