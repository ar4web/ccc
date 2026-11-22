# CCC — Company Command Center

The original V1 workbook has five data and operations tabs. The CCC refresh adds a formula-driven **Action Center** and **Kanban Board** for a seven-tab full workbook. The sample records are synthetic and visibly labelled.

## Download a version

- [Full workbook — automatic calculation](CCC_Company_Command_Center.xlsx): all seven tabs, interactive Dashboard KPIs, live alerts, and Kanban. Changed formulas calculate automatically. A full recalculation is no longer forced every time the file opens. The first Excel open may still take time because many original formulas have no saved results; let it finish and save once.
- [Full workbook — manual calculation](CCC_Manual_Recalc.xlsx): same seven tabs and formulas, but Excel does not recalculate automatically. Press **F9** to refresh. Formula cells without a saved result may appear blank until calculated.
- [Quick View — static snapshot](CCC_Quick_View.xlsx): one color-coded sheet for quick viewing, including the saved action-flow KPI strip. It has no formulas and cannot refresh; all values are saved sample values only.

## Dashboard and Kanban

The Dashboard keeps the core cards for active staff, expired Iqamas, items due within 15 days, and open findings. A second KPI band shows submitted timesheets, flagged entries, total workforce, and timesheet coverage. A compact action strip adds unassigned actions, overdue follow-ups, in-progress actions, and done/resolved actions without repeating the payroll figures already shown above.

The **Action Center** lists formula-driven alerts for expired Iqamas, renewals due within 15 days, open findings, missing current-period timesheets, held payroll, and missing Iqama dates. Each alert links to its source tab. Fill in the blue **Owner**, **Follow-up by**, **Action status**, and **Remarks** cells to manage work.

The **Kanban Board** sorts each active alert into Open, In Progress, or Done/Resolved. Change an alert’s Action status in the Action Center and its card moves automatically. Alerts remain open until their underlying source metric is resolved. The Dashboard opens the Kanban board; the board links back to the Action Center.

## Regenerate the files

```bash
python -m pip install -r requirements.txt
python polish_ccc.py /path/to/CCC-V1.xlsx
python make_ccc_variants.py CCC_Company_Command_Center.xlsx
```

`polish_ccc.py` preserves original formula caches; it does not calculate formulas. It also preserves Action Center owner, date, status, and remark inputs when refreshing an existing CCC output. `make_ccc_variants.py` creates the manual-calculation and static Quick View copies. Confirm the period and totals in Excel before operational use.

`CCC-V1.xlsx` is already built. It is not the raw multi-sheet input expected by the separate `build_sheet.py` pipeline. Keep real employee and payroll workbooks in approved, access-controlled storage, not Git.
