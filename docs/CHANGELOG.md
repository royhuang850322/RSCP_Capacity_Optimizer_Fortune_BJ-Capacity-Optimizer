# Fortune BJ Capacity Optimizer Changelog

## v0.1.6 - 2026-07-29

- Normalize GUI path display to Windows backslashes after browsing for files or folders.
- Normalize manually pasted path fields before starting an analysis run.
- Keep report, license, and fingerprint path messages consistent with Windows path separators.

## v0.1.5 - 2026-07-29

- Change forecast-demand routing to use `工艺路线-北京.xlsx`, `工艺路线-沈阳.xlsx`, and `工艺路线-南通.xlsx` instead of production-order history.
- Match forecast routes by priority: Beijing, then Shenyang, then Nantong.
- Skip unmatched forecast materials without blocking the run and list them in the data-quality report.
- Add forecast route source fields to ModeA and ModeB order-operation detail reports.
- Change workgroup heatmap coloring to four fixed bands: 0-25% dark green, 25-75% light green, 75-100% light red, and above 100% dark red.
- Remove `模拟数据导入` from packaged application output.

## v0.1.4 - 2026-07-29

- Add optional demand forecast import from `需求预测_产能分析输入模板.xlsx`, generating weekly Sunday virtual forecast orders with `FCST-YYYYMMDD-物料号` order numbers.
- Add forecast precheck blocking when forecast materials do not have matching historical production-operation routes.
- Add operation transfer modes for ModeA and ModeB: full-batch, half-batch, single-piece flow, and due-date-forced load.
- Switch optimization start handling to exact start date and keep weekly/monthly reports aligned to the selected analysis granularity.
- Add forecast source fields and period split fields to order-operation detail reports.
- Move the forecast import toggle into the main parameter panel and disable mouse-wheel value changes on parameter controls.
- Update heatmap colors: 0-100% uses dark-to-light green, and loads above 100% use light-to-dark red.

## v0.1.3 - 2026-07-27

- Align ModeA and ModeB to use the same optimization start period handling for overdue or early-start order chains.
- Keep original order due dates available for reporting while moving affected unfinished orders into the selected analysis start period.
- Add optional-operation unit-hour inheritance: blank non-outsourced optional hours can inherit the matching production operation unit hours by material and activity.
- Block analysis with a precheck report when a blank optional operation has no matching production operation unit hour to inherit.
- Add unit-hour source fields to ModeB optional split and order-operation detail reports.
- Refresh the customer usage and report-reading guide for v0.1.3.

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
