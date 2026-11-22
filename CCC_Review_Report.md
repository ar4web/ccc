# CCC V1 review

**Status:** The refreshed CCC workbook adds a dynamic Action Center and Kanban Board to the original five-tab model. The user confirmed the sample records are synthetic; all outputs are labelled accordingly. The source workbook is unchanged.

## What changed

- The Dashboard has four lead cards: active staff, expired Iqamas, items due within 15 days, and open findings. A second KPI band shows submitted timesheets, flagged entries, total workforce, and timesheet coverage. A compact third strip shows payroll ready/held amounts plus unassigned and overdue follow-ups.
- The live status line shows filtered timesheet rows, payroll estimates, open findings, and open actions. A Dashboard link opens the Kanban board.
- **Action Center** has six formula-driven alerts: expired Iqamas, renewals due soon, open findings, no current-period timesheet input, held payroll, and missing Iqama dates. Each links to its source tab and has editable owner, follow-up date, action status, and remarks. Summary formulas count open, unassigned, and overdue items.
- **Kanban Board** sorts alert cards into Open, In Progress, and Done/Resolved. Changing Action status in the Action Center moves the card; resolving the source metric clears the alert. A test with a temporary owner, due date, status, and remark confirmed that the card and Dashboard counts update as intended.
- Alert labels and colors distinguish urgent, due soon, input needed, in-progress, and clear states. Detailed tables remain available; the renewal worklist stays visible on the Dashboard.
- Three outputs are included: full automatic, full manual-calculation, and static Quick View.

## Why the workbook felt slow

The full workbook now has **117,854 formulas**: 117,796 in V1, plus 34 Action Center formulas, 22 Kanban formulas, and two new Dashboard formulas. **103,946 original formulas have no saved result cache**, and V1 was set to force a full recalculation on every open. The file size is not the main cause; Excel formula calculation is. Operations Control also has 841 `TODAY()` formulas; centralizing their date logic may reduce volatility, but should be tested against the intended business rule.

The automatic workbook no longer forces a full recalculation on every open. Excel may still calculate missing results the first time; save after it finishes. The manual workbook avoids automatic recalculation until **F9**. The static Quick View opens quickly but cannot refresh.

## Verification

- The full automatic and manual workbooks have seven tabs, 11 Excel tables, 10 data validations, and three Salarysheet charts. The two Dashboard charts were removed.
- The automatic copy preserves source input values and original formula caches except for the updated Dashboard summaries. New Action Center/Kanban formulas have sample caches derived from V1's saved values.
- Structural checks found zero cached formula errors. The source contains no live current-period timesheet rows or numeric payroll results. The displayed 147 active staff, 75 expired Iqamas, 8 due within 15 days, 17 findings, and SAR 0 ready/held are sample/cached values—not a validated payroll run.
- No Excel calculation engine is available here, so full formula recalculation and actual Excel speed were not benchmarked. Validate totals in Excel before operational use.

## Further improvements before operational use

1. Test a fictional monthly timesheet and reconcile overtime, absences, deductions, duplicate IDs, and release holds against independent expected results.
2. `build_sheet.py` expects a different raw workbook; `CCC-V1.xlsx` is already built and is not a compatible raw input. Reconcile its historical builder assumptions with V1 before attempting a rebuild.
3. Replace deductions keyed by employee name and adjustments parsed from formula text with a reviewed adjustment table tied to stable records, reasons, and evidence.
4. Replace row-position pairing between attendance and payroll with stable IDs and key-based lookups. Add a timesheet deadline and period-close status before classifying missing submissions as late.
5. The workbook covers workforce, attendance, payroll, document expiry, and related controls; it does not provide company-wide P&L, cash, sales, inventory, or project KPIs without those source datasets.

The selected Dashboard period remains September 2026, matching the original saved view. Confirm or change it in Excel.
