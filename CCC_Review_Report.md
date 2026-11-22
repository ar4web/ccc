# CCC V1 review

**Status:** The full five-tab demo workbook is committed in this branch as `CCC_Company_Command_Center.xlsx`. The user confirmed the sample records are synthetic; the dashboard is labelled **SYNTHETIC DATA**. The original V1 source workbook is unchanged.

## Dashboard changes

- Branded the workbook **CCC — Company Command Center** and labelled it **SYNTHETIC DATA**.
- Kept four lead KPIs: active employees, expired Iqamas, Iqamas due within 15 days, and open findings.
- Added a concise strip for live timesheet input count and ready/held payroll amounts. An empty timesheet period is described as “No live timesheet inputs,” not as overdue.
- Collapsed advanced filters and detailed exception/ranking sections; the renewal worklist remains visible. Removed two Dashboard charts; the three Salarysheet charts and all underlying tables remain available.
- Made Dashboard the opening tab and set its print layout to one landscape page.

## Verification completed

- Five expected sheets and ten Excel tables are present; eight Dashboard data validations are retained.
- Checked 117,796 formulas. Formula count is unchanged, no external workbook links were found, and there are zero cached formula errors or `#REF!`/`DUMMYFUNCTION` formulas.
- Only two formulas were changed: the compact Dashboard status line and the “Open findings” KPI. All other formula caches were preserved. Input values were unchanged.
- The source contains no live timesheet input rows or numeric payroll results. The displayed cached values (147 active staff, 75 expired Iqamas, 8 due within 15 days, 17 findings, and SAR 0 ready/held) are sample/cached values—not evidence of a live payroll run.
- The workbook was not recalculated in this environment. LibreOffice is unavailable; openpyxl does not calculate formulas. Excel is set to perform full automatic recalculation on open, but the resulting totals still need to be checked.

## Remaining improvements before operational use

1. Open the copy in Excel, confirm the report period and totals, and test a fictional monthly timesheet through payroll. Reconcile overtime, absences, deductions, duplicate IDs, and release holds to independently calculated expected results.
2. The separate `build_sheet.py` pipeline expects a different raw workbook; `CCC-V1.xlsx` is already built and is not a compatible raw input. Reconcile the builder’s historical assumptions with V1 before attempting a rebuild.
3. In the builder logic, replace deductions keyed by employee name and adjustments parsed from formula text with a reviewed adjustment table tied to stable source records, reasons, and evidence.
4. Replace row-position pairing between attendance and payroll with stable IDs and key-based lookups. Add an explicit timesheet deadline and period-close status before classifying missing submissions as late.
5. Keep the scope honest: this workbook covers workforce, attendance, payroll, document expiry, and related controls. It does not provide reliable company-wide P&L, cash, sales, inventory, or project KPIs without those source datasets.

The Dashboard uses September 2026 as its selected report month, matching the original saved view. Confirm or change the period in Excel after recalculation.
