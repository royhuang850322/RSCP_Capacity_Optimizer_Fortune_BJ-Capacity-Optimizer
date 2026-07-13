# Fortune BJ Capacity Optimizer Changelog

## v0.1.1 - 2026-07-13

- Rename demand input file to `订单交期数量_产能分析输入模板.csv`.
- Replace urgent order flag with strict `紧急类型` priority values: `越库`, `T-14`, `RTM`, `临时加急`.
- Block analysis when `紧急类型` contains non-standard values and report exact rows in the precheck workbook.
- Keep ModeA overdue priority without moving original due dates; move ModeB overdue orders into the optimization start period.
- Change ModeA/ModeB heatmap report from work-center view to resource-group workgroup view.
- Add application version `0.1.1` to the GUI title and report runtime sheet.
- Add microsecond-level report filenames to avoid collisions during quick repeated runs.
