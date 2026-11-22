# CCC — Company Command Center

The original V1 workbook contains five data and operations tabs. The CCC refresh adds a sixth, formula-driven **Action Center**. The sample data are synthetic and clearly labelled.

## Download a version

- [Full workbook — automatic calculation](CCC_Company_Command_Center.xlsx): all six tabs, dynamic KPIs and alerts. Changed formulas recalculate automatically. It no longer forces a full recalculation on every open. The first Excel open may still take time because many original formulas have no saved results; let it finish and save once.
- [Full workbook — manual calculation](CCC_Manual_Recalc.xlsx): same six tabs and alerts, but Excel does not recalculate automatically. Press **F9** to refresh. Formula cells with no saved result may appear blank until calculated.
- [Quick View — static snapshot](CCC_Quick_View.xlsx): one color-coded sheet for quick viewing. It has no formulas and cannot refresh; it shows saved sample values only.

## Using the Action Center

The Dashboard links to **Action Center**, which lists formula-driven alerts for expired Iqamas, renewals due within 15 days, open findings, missing current-period timesheets, held payroll, and missing Iqama dates. Each alert links to its source tab. Fill in the blue **Owner**, **Follow-up by**, **Action status**, and **Remarks** cells to manage work. The summary shows open, unassigned, and overdue follow-ups. Alerts clear when the underlying source metric is resolved; entering a “Done” action status does not hide an unresolved data alert.

## Regenerate the files

```bash
python -m pip install -r requirements.txt
python polish_ccc.py /path/to/CCC-V1.xlsx
python make_ccc_variants.py CCC_Company_Command_Center.xlsx
```

`polish_ccc.py` preserves original formula caches; it does not calculate formulas. It also preserves Action Center owner, date, status, and remark inputs when refreshing an existing CCC output. `make_ccc_variants.py` creates the manual-calculation and static Quick View copies. Confirm the period and totals in Excel before operational use.

`CCC-V1.xlsx` is already built. It is not the raw multi-sheet input expected by the separate `build_sheet.py` pipeline. Keep real employee and payroll workbooks in approved, access-controlled storage, not Git.
