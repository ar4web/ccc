# CCC — Company Command Center

The full five-tab CCC demo workbook is available in this branch. The sample records are synthetic, as confirmed by the user.

## Download the workbook

[Download the full CCC Company Command Center workbook](CCC_Company_Command_Center.xlsx)

The Dashboard is visibly labelled **SYNTHETIC DATA**. Open the file in Excel to recalculate formulas and check the displayed period and totals before using it as a template.

## Rebuild or refresh the workbook

`polish_ccc.py` creates a focused Dashboard copy from an existing five-tab workbook. It does not overwrite the source:

```bash
python -m pip install -r requirements.txt
python polish_ccc.py /path/to/CCC-V1.xlsx
```

The default output is `CCC_Company_Command_Center.xlsx`. Formula caches are preserved and Excel is asked to recalculate on open; this script does **not** calculate formulas itself. The committed demo workbook is tracked on this branch; regenerated changes can be committed if desired.

`CCC-V1.xlsx` is already built. It is not the raw multi-sheet input expected by the separate `build_sheet.py` pipeline. Keep real employee and payroll workbooks in approved, access-controlled storage, not Git.
