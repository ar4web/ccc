# CCC V1 review

**Status:** The refreshed CCC workbook now adds a dynamic Action Center to the original five-tab model. The user confirmed that the sample records are synthetic; all outputs are labelled accordingly. The source workbook is unchanged.

## What changed

- The Dashboard is branded **CCC — Company Command Center** and remains focused on active staff, expired Iqamas, items due within 15 days, and open findings.
- A live status line shows filtered timesheet rows, ready/held payroll estimates, findings, and open actions. The Dashboard links directly to the Action Center.
- The new **Action Center** adds six formula-driven alerts: expired Iqamas, renewals due soon, open findings, no current-period timesheet input, held payroll, and missing Iqama dates. Each has a source link plus editable owner, follow-up date, action status, and remarks. Summary counts show open, unassigned, and overdue follow-ups. Source issues clear their alerts when the underlying metric is resolved.
- Red/amber/green labels and formatting distinguish urgent items, due-soon items, input needed, and clear items. Detailed sections remain available in the other tabs; the renewal worklist stays visible on the Dashboard.
- Three files are included: automatic full workbook, manual F9 full workbook, and a fast static Quick View.

## Why the workbook felt slow

The full workbook has **117,831 formulas** (117,796 in V1, plus 35 for the Action Center and Dashboard link). **103,946 original formulas have no saved result cache**, and V1 was set to force a full recalculation on every open. The file size is not the main cause; Excel formula calculation is. There are also 841 `TODAY()` calls on Operations Control; centralizing their date logic may help further, but should be tested against the intended business rule.

The automatic workbook no longer forces a full recalculation on every open. Excel may still calculate missing results the first time; save after it finishes. The manual workbook avoids automatic recalculation until **F9**. The static Quick View opens quickly but cannot refresh.

## Verification

- The automatic and manual full workbooks have six tabs, 11 Excel tables, 10 data validations, and three Salarysheet charts. The two Dashboard charts are removed.
- The automatic copy retains all original formulas and caches except the updated Dashboard status/KPI formulas; the Action Center adds formula-driven alert cells with sample caches. No existing source input values were changed.
- Structural checks found zero cached formula errors. The original data contain no live current-period timesheet rows or numeric payroll results. The displayed 147 active staff, 75 expired Iqamas, 8 due within 15 days, 17 findings, and SAR 0 ready/held are sample/cached values—not a validated payroll run.
- No Excel calculation engine is available here, so actual recalculation speed and transaction results were not benchmarked. Validate the formulas in Excel before operational use.

## Further improvements before operational use

1. Test a fictional monthly timesheet and reconcile overtime, absences, deductions, duplicate IDs, and release holds against independent expected results.
2. `build_sheet.py` expects a different raw workbook; `CCC-V1.xlsx` is already built and is not a compatible raw input. Reconcile its historical builder assumptions with V1 before attempting a rebuild.
3. Replace deductions keyed by employee name and adjustments parsed from formula text with a reviewed adjustment table tied to stable records, reasons, and evidence.
4. Replace row-position pairing between attendance and payroll with stable IDs and key-based lookups. Add a timesheet deadline and period-close status before classifying missing submissions as late.
5. The workbook covers workforce, attendance, payroll, document expiry, and related controls; it does not provide company-wide P&L, cash, sales, inventory, or project KPIs without those source datasets.

The selected Dashboard period remains September 2026, matching the original saved view. Confirm or change it in Excel.
