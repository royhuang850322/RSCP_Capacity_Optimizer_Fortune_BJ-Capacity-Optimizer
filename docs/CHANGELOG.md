# Fortune BJ Capacity Optimizer Changelog

## v0.1.2 - 2026-07-24

- Split ModeB optimized allocation load across reporting weeks/months in the order-operation detail, period capacity report, workgroup heatmap, optimization overview, and capacity recommendation sheets.
- Add ModeB detail fields for allocation-period split rows and period-level original/optimized load hours, so detail totals reconcile with heatmap totals.
- Keep ModeB OR-Tools optimization at period-level integer product allocation; the new split only changes report load attribution and does not add solver variables.
- Update default input paths for current workbook-based demand, work center, optional operation, and calendar templates.
- Update packaging to include `.csv`, `.xlsx`, and `.xls` input templates in the packaged application folder.

## v0.1.1 - 2026-07-13

- Rename demand input file to `订单交期数量_产能分析输入模板.csv`.
- Replace urgent order flag with strict `紧急类型` priority values: `越库`, `T-14`, `RTM`, `临时加急`.
- Block analysis when `紧急类型` contains non-standard values and report exact rows in the precheck workbook.
- Keep ModeA overdue priority without moving original due dates; move ModeB overdue orders into the optimization start period.
- Change ModeA/ModeB heatmap report from work-center view to resource-group workgroup view.
- Add application version `0.1.1` to the GUI title and report runtime sheet.
- Add microsecond-level report filenames to avoid collisions during quick repeated runs.
