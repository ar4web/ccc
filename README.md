# CCC — Company Command Center

The sample records are synthetic. The Dashboard and files below are clearly labelled **SYNTHETIC DATA**.

## Choose the version you need

- [Full workbook — automatic calculation](CCC_Company_Command_Center.xlsx): five tabs and all formulas. Excel recalculates changed cells automatically. It no longer forces a full recalculation every time the file opens. Since many source formulas have no saved results, the first Excel open may still take time; let it finish and save once.
- [Full workbook — manual calculation](CCC_Manual_Recalc.xlsx): same five tabs and formulas, but Excel does not recalculate automatically. Press **F9** to update. Formula cells without a saved result may appear blank until you calculate.
- [Quick View — static snapshot](CCC_Quick_View.xlsx): one color-coded dashboard sheet for a fast open. It has no formulas and cannot refresh; it shows saved sample values only.

## Regenerate the files

Install dependencies and use the existing five-tab V1 workbook as input:

```bash
python -m pip install -r requirements.txt
python polish_ccc.py /path/to/CCC-V1.xlsx
python make_ccc_variants.py CCC_Company_Command_Center.xlsx
```

`polish_ccc.py` preserves source formula caches; it does not calculate formulas. `make_ccc_variants.py` produces the manual-calculation and static Quick View copies. Check periods and totals in Excel before operational use.

`CCC-V1.xlsx` is already built. It is not the raw multi-sheet input expected by the separate `build_sheet.py` pipeline. Keep real employee and payroll workbooks in approved, access-controlled storage, not Git.
