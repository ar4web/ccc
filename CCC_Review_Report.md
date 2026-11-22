# CCC V1 review

**Status:** Three labelled outputs are available in this branch: the full automatic workbook, the full manual-calculation workbook, and a fast static Quick View. The user confirmed that the sample records are synthetic. The original V1 source workbook is unchanged.

## Why the workbook feels slow

- The full workbook contains **117,796 formulas**; **103,946 have no saved result cache**. It was configured to force a full recalculation on every open. The 2.4 MB file size is not the main cause—the formula recalculation is.
- The Operations Control sheet contains **841 `TODAY()` formulas**, which are volatile. They add recalculation work and are a possible later optimization, subject to confirming the intended date logic.
- The no-forced-recalc setting on the automatic version prevents an unnecessary full recalculation on every subsequent open. Excel may still need to calculate missing results on the first open. Save the file after calculation so Excel can retain those results.

## Choose a version

- `CCC_Company_Command_Center.xlsx` — full workbook, automatic calculation for changed formulas; no forced full recalc on every open. First-open calculation may still be slow because source caches are incomplete.
- `CCC_Manual_Recalc.xlsx` — full workbook, manual calculation. Press F9 when you want to update. Cells with no saved result can appear blank until calculated.
- `CCC_Quick_View.xlsx` — one-sheet, formula-free snapshot for quick viewing. It is explicitly marked static and shows saved sample values only; it cannot update or replace the full workbook.

The Quick View makes categories more distinct: green for active, amber for due soon, and red for action/review. The full Dashboard also has distinct cards for active employees, expired Iqamas, items due within 15 days, and open findings.

## Verification

- Full versions retain five expected sheets, ten tables, eight Dashboard data validations, and three Salarysheet charts. The Dashboard charts were removed to reduce clutter.
- The automatic full workbook retains 117,796 formulas and has zero cached formula errors. All non-Dashboard formula caches were preserved; the compact Dashboard status line and “Open findings” KPI are updated.
- The sample contains no live timesheet entries or numeric payroll results. Cached Dashboard values (147 active staff, 75 expired Iqamas, 8 due within 15 days, 17 findings, and SAR 0 ready/held) are sample values, not a validated payroll run.
- No Excel calculation engine is available here, so actual recalculation time and formula results were not benchmarked. Confirm all totals in Excel before operational use.

## Remaining logic improvements before operational use

1. Test a fictional monthly timesheet and reconcile overtime, absences, deductions, duplicate IDs, and release holds against independent expected results.
2. `build_sheet.py` expects a different raw workbook; `CCC-V1.xlsx` is already built and is not a compatible raw input. Reconcile historical builder assumptions with V1 before rebuilding.
3. Replace deductions keyed by employee name and adjustments parsed from formula text with a reviewed adjustment table tied to stable records, reasons, and evidence.
4. Replace row-position pairing between attendance and payroll with stable IDs and key-based lookups. Add a clear timesheet deadline and period-close status before classifying missing submissions as late.
5. The workbook covers workforce, attendance, payroll, document expiry, and related controls; it does not provide company-wide P&L, cash, sales, inventory, or project KPIs without those source datasets.

The selected Dashboard period remains September 2026, matching the original saved view. Confirm or change it in Excel.
