from __future__ import annotations

import html
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs"
OUT_PATH = OUT_DIR / "Fortune_BJ_Capacity_Allocation_Logic_CN_V02.docx"


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def run(text: str, *, bold: bool = False) -> str:
    props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f"<w:r>{props}<w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r>"


def paragraph(text: str = "", *, style: str | None = None, bold: bool = False) -> str:
    ppr = f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>" if style else ""
    return f"<w:p>{ppr}{run(text, bold=bold) if text else ''}</w:p>"


def bullets(items: list[str]) -> str:
    parts = []
    for item in items:
        parts.append(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"ListBullet\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            f"{run(item)}</w:p>"
        )
    return "".join(parts)


def table(headers: list[str], rows: list[list[object]]) -> str:
    total_width = 9360
    base_width = total_width // max(len(headers), 1)
    widths = [base_width for _ in headers]
    if widths:
        widths[-1] = total_width - sum(widths[:-1])

    def cell(value: object, width: int, *, header: bool = False) -> str:
        shade = "<w:shd w:fill=\"D9EAF7\"/>" if header else ""
        return (
            f"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/>"
            "<w:vAlign w:val=\"center\"/>"
            f"{shade}</w:tcPr>{paragraph(str(value), bold=header)}</w:tc>"
        )

    grid = "".join(f"<w:gridCol w:w=\"{width}\"/>" for width in widths)
    xml = [
        "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/>"
        "<w:tblW w:w=\"9360\" w:type=\"dxa\"/><w:tblInd w:w=\"120\" w:type=\"dxa\"/>"
        "<w:tblLayout w:type=\"fixed\"/></w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
    ]
    xml.append("<w:tr>" + "".join(cell(h, widths[i], header=True) for i, h in enumerate(headers)) + "</w:tr>")
    for row in rows:
        padded = list(row) + [""] * max(len(headers) - len(row), 0)
        xml.append("<w:tr>" + "".join(cell(v, widths[i]) for i, v in enumerate(padded[: len(headers)])) + "</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


def section(title: str, body: list[str]) -> str:
    return paragraph(title, style="Heading1") + "".join(body)


def formula(text: str) -> str:
    return paragraph(text, style="Formula")


def build_body() -> str:
    parts: list[str] = []
    parts.append(paragraph("Fortune BJ 产能优化工具计算逻辑说明", style="Title"))
    parts.append(paragraph("适用范围：数据导入、分析前校验、ModeA 无限产能分析、ModeB 100% 产能优化建议、可选工序、外包、热处/表处专用逻辑"))
    parts.append(paragraph("版本：V02 | 生成日期：2026-07-05"))
    parts.append(paragraph("核心结论", style="Heading1"))
    parts.append(paragraph("本工具不是车间 APS 执行排产工具，而是周期级产能优化与建议工具。它先把订单需求、完整工序路线、工作中心能力、日历、可选工序和外包路径转成可计算的负荷，再在 ModeB 中用 OR-Tools 按整数产品数量选择原工作中心、可选工作中心或外包路径。报告重点解释每个周期的负荷、缺口、分流、外包释放和补产能建议。"))
    parts.append(bullets([
        "ModeA：按订单交期倒排，形成无限产能压力视图，用于识别瓶颈和原始负荷。",
        "ModeB：以 ModeA 倒排后的周期为基线，在同一周期内用 OR-Tools 做整数产品数量分配。",
        "同一道工序可以按整数件拆到多个路径，但不能按非整数小时拆分产品。",
        "外包路径视为无限产能，不占用厂内工作中心；当前固定按 7 天日历返回。",
        "热处/表处专用逻辑当前是周期级容量折算，不输出每炉或每条流水线的执行顺序。",
    ]))

    parts.append(section("1. 当前订单数量和工序行数怎么看", [
        paragraph("工具会读取订单数量和工序行数，但这两个指标只用于确认需求范围、路线完整性和数据规模。真正进入产能判断的是每道工序换算后的负荷小时，以及该负荷小时落在哪个工作中心、哪个分析周期。"),
        formula("普通工序单位工时 = 标准值1 + 标准值2 + 标准值3"),
        formula("普通工序负荷小时 = 订单数量 × 普通工序单位工时"),
        table(["订单", "数量", "单位工时", "工序负荷"], [
            ["A", 10, "2 小时/件", "10 × 2 = 20 小时"],
            ["B", 2, "15 小时/件", "2 × 15 = 30 小时"],
        ]),
        paragraph("补充说明：为什么不能只看订单数量或工序行数。这个例子中，订单 B 只有 2 件，但负荷是 30 小时，高于订单 A 的 20 小时。因此订单数、件数、工序行数只能帮助理解数据规模，不能直接判断产能压力。"),
    ]))

    parts.append(section("2. 符号和变量解释", [
        table(["符号", "中文含义", "说明"], [
            ["o", "订单", "来自订单交期数量中的富创单据号，与生产订单工序的订单字段对齐。"],
            ["a", "活动/工序号", "同一订单内按活动号排序。"],
            ["m", "物料", "用于匹配可选工序表。"],
            ["w", "工作中心", "由生产订单工序的工序短文本匹配工作中心表。"],
            ["q(o)", "订单数量", "同一富创单据号的华创需求数量累加。"],
            ["h(o,a)", "单位工时", "标准值1+标准值2+标准值3。"],
            ["load(o,a,w)", "工序负荷小时", "普通工序为 q(o)×h(o,a)。"],
            ["cap(p,w)", "周期产能小时", "工作中心数量×平均每日小时×周期天数。"],
            ["x(o,a,r)", "ModeB 分配数量", "工序 o,a 分配到路径 r 的整数产品数量。"],
            ["short(p,w)", "周期缺口小时", "优化后负荷超过周期产能的部分。"],
        ]),
    ]))

    parts.append(section("3. 输入文件如何变成计算对象", [
        table(["输入文件", "关键字段", "进入模型后的作用"], [
            ["订单交期数量_产能分析输入模板.csv", "富创单据号、华创需求数量、供给日期、紧急类型", "形成订单需求、数量、交期和紧急类型优先级。"],
            ["生产订单工序_排产输入模板.csv", "订单、活动、物料、工序短文本、标准值1/2/3", "形成订单工序路线、工序排序、工作中心映射和单位工时。"],
            ["工作中心_排产输入模板.csv", "工作中心、资源组分类、数量、日历名称", "形成工作中心资源能力；热处/表处字段可覆盖容量计算方式。"],
            ["工作日历_排产输入模板.csv", "日历名称、每日工作小时、每周工作天数", "形成平均每日可用小时。"],
            ["可选工序_排产输入模板.csv", "物料、活动、可选工作中心、可选资源组分类、可选单位工时", "形成 ModeB 可选路径；可选资源组分类为外包时视为外包路径。"],
        ]),
        paragraph("例子：订单 100000001，需求数量 5，交期 2026-07-20；生产订单工序有活动 10、20、30。活动 20 的工序短文本匹配到工作中心 A，标准值为 1、0.5、0，则活动 20 单位工时为 1.5 小时/件，负荷为 5×1.5=7.5 小时。"),
    ]))

    parts.append(section("4. 分析前硬校验逻辑", [
        paragraph("正式计算前，工具先做硬校验。只要有效需求订单缺工序，或参与分析的工序短文本缺工作中心映射，正式产能分析就停止，并输出分析前数据校验报告。"),
        table(["校验项", "失败条件", "原因"], [
            ["订单是否都有工序", "订单交期数量中有订单，但生产订单工序中找不到该订单", "没有工序路线，无法计算工作中心和工时。"],
            ["工序是否都有工作中心", "参与分析的工序短文本不在工作中心表中，且不是外协", "无法判断工序占用哪个资源。"],
            ["必填列", "输入文件缺少必填表头", "无法构建模型字段。"],
        ]),
        paragraph("例子：订单交期数量中有富创单据号 100000006038，数量 3，供给日期 2026-07-06；但生产订单工序文件没有订单=100000006038 的任何行。工具会报告“订单缺失工序”，并停止正式分析。"),
    ]))

    parts.append(section("5. 订单需求和交期处理", [
        formula("同一订单数量 = Σ 华创需求数量"),
        paragraph("同一富创单据号如果出现多行，工具会把数量累加，并取最早供给日期作为该订单的交期。"),
        paragraph("如果 ModeB 设置了优化开始周期，而订单交期早于优化开始周期，工具会把该订单转入优化开始周期计算，并自动视作紧急订单。这是为了处理历史遗留未生产订单。"),
        table(["输入行", "数量", "供给日期", "汇总结果"], [
            ["订单 A 第1行", 2, "2026-07-10", "数量累计"],
            ["订单 A 第2行", 3, "2026-07-20", "数量累计"],
            ["汇总", 5, "2026-07-10", "数量=5，交期取最早 2026-07-10"],
        ]),
    ]))

    parts.append(section("6. 工作中心周期产能计算", [
        formula("平均每日小时 = 每日工作小时 × 每周工作天数 / 7"),
        formula("周期产能小时 = 工作中心数量 × 平均每日小时 × 周期天数"),
        paragraph("如果优化粒度选择周，周期天数通常为 7；如果选择月，周期天数为该月自然天数。"),
        table(["工作中心", "数量", "每日小时", "每周工作天数", "周期", "周期产能"], [
            ["北京立加", 2, 24, 7, "1 周", "2×24×7 = 336 小时"],
            ["焊接", 1, 16, 5, "1 周", "1×(16×5/7)×7 = 80 小时"],
        ]),
    ]))

    parts.append(section("7. ModeA 无限产能分析逻辑", [
        paragraph("ModeA 按订单交期倒排同一订单内的工序。它不检查工作中心是否同一时间被多道工序占用，因此叫无限产能分析。它的作用是识别周期压力和瓶颈，而不是生成可执行排程。"),
        formula("每道工序结束时间 = 后序工序开始时间或订单交期"),
        formula("每道工序开始时间 = 工序结束时间 - 工序负荷小时"),
        table(["订单", "活动", "负荷小时", "倒排结果"], [
            ["A", 30, 4, "交期 7月10日 00:00 结束，7月9日 20:00 开始"],
            ["A", 20, 6, "接活动30开始点，7月9日 14:00 开始"],
            ["A", 10, 10, "接活动20开始点，7月9日 04:00 开始"],
        ]),
        paragraph("ModeA 产能分析按周或月汇总工作中心负荷。若某工作中心周期负荷小时大于周期产能小时，则识别为瓶颈。"),
    ]))

    parts.append(section("8. ModeB 100% 产能优化建议逻辑", [
        paragraph("ModeB 先执行 ModeA 形成基线，然后按每个工序落入的周期做 OR-Tools 整数产品数量优化。工具不是把单个工序按小时切开，而是把产品数量作为整数变量。"),
        formula("对每个工序：Σ x(o,a,r) = q(o)"),
        paragraph("r 表示可选路径，包括原工作中心、可选工作中心和外包。每个 x 都是整数产品数量。"),
        table(["路径", "变量含义", "是否占厂内产能"], [
            ["原工作中心", "保留在原工作中心的产品数量", "是"],
            ["可选工作中心", "转到可选工作中心的产品数量", "是"],
            ["外包", "转外包的产品数量", "否"],
        ]),
        paragraph("例子：某工序订单数量 5 件，原工作中心 A 单位工时 10 小时，可选工作中心 B 单位工时 12 小时，外包可用。OR-Tools 可以求出 A=1、B=2、外包=2。总数量仍为 1+2+2=5，且没有 1.5 件这种不可生产拆分。"),
    ]))

    parts.append(section("9. ModeB 工作中心约束和缺口变量", [
        paragraph("每个周期、每个工作中心都有产能约束。为了保证即使超产能也能输出建议，模型使用缺口变量 short，而不是强制所有工作中心必须小于等于 100%。"),
        formula("Σ 厂内负荷小时(p,w) <= 周期产能小时(p,w) + short(p,w)"),
        formula("short(p,w) = max(优化后负荷 - 周期产能, 0)"),
        paragraph("如果优化后仍超 100%，报告会显示实际负荷率，并把缺口转成外包、加班、每日增加小时/台、新增设备数等建议。"),
        table(["周期", "工作中心", "产能小时", "优化后负荷", "缺口", "负荷率"], [
            ["W28", "A", 100, 130, 30, "130%"],
            ["W28", "B", 120, 90, 0, "75%"],
        ]),
    ]))

    parts.append(section("10. OR-Tools 目标函数如何选择路径", [
        paragraph("ModeB 的目标函数优先减少产能缺口，其次减少超载工作中心周期数，再惩罚外包、路径变更和额外工时。"),
        formula("Minimize 1,000,000×Σshort + 100,000×Σoverloaded + 1,000×Σoutsource_hours + 100×Σroute_change + 10×Σextra_hours"),
        table(["惩罚项", "含义", "业务效果"], [
            ["short", "优化后仍超过产能的小时", "最高优先级，尽量减少超 100% 缺口。"],
            ["overloaded", "是否存在超载工作中心周期", "同等缺口下倾向减少超载点数量。"],
            ["outsource_hours", "转外包释放的原工时", "外包可用但不是无成本优先。"],
            ["route_change", "是否从原路径转移", "没有必要时尽量不改路径。"],
            ["extra_hours", "可选路径比原路径多出的小时", "同等条件下倾向工时更低路径。"],
        ]),
        paragraph("例子：A 工作中心本周多出 20 小时。若某工序可以转到 B 后释放 A 20 小时，即使 B 单位工时多 2 小时，模型通常仍会选择转移，因为减少 short 的收益远高于额外工时惩罚。"),
    ]))

    parts.append(section("11. 可选工序和外包逻辑", [
        paragraph("可选工序按物料 + 活动匹配。非外包路径必须在工作中心表中存在；可选资源组分类等于“外包”时，工具视为外包路径。"),
        paragraph("外包不占用厂内工作中心产能，当前固定按 7 天日历返回。外包单位工时即使填写，也会被忽略。"),
        table(["字段", "规则"], [
            ["物料 + 活动", "决定可选工序是否匹配当前工序。"],
            ["可选工作中心", "非外包时必须存在于工作中心表。"],
            ["可选资源组分类=外包", "视为外包无限产能路径。"],
            ["可选单位工时", "非外包路径用于计算负荷；外包路径忽略。"],
        ]),
    ]))

    parts.append(section("12. 热处/表处同机加逻辑", [
        paragraph("当 GUI 选择“同机加逻辑”，或工序没有被识别为热处/表处，或工作中心没有维护专用参数时，热处/表处工序仍按普通工时计算。"),
        formula("热处/表处同机加负荷 = 订单数量 × (标准值1+标准值2+标准值3)"),
        paragraph("例子：某热处理工序数量 20 件，标准值合计 0.5 小时/件，则同机加逻辑下负荷为 20×0.5=10 小时。"),
    ]))

    parts.append(section("13. 热处/表处批量处理逻辑（当前代码口径）", [
        paragraph("当选择“热处/表处专用逻辑”，且工作中心产能计算类型为批量处理时，工具按容量占用折算炉次。当前代码口径是周期级折算，不做每炉装载清单。"),
        formula("容量占用 = 产品数量 × 单件容量占用"),
        formula("炉次 = ceil(容量占用 / 单炉容量)"),
        formula("批量处理负荷小时 = 炉次 × (单炉周期小时 + 装卸/准备小时)"),
        table(["项目", "数值"], [
            ["产品数量", "25 件"],
            ["单件容量占用", "1 容量单位/件"],
            ["单炉容量", "10 容量单位"],
            ["单炉周期小时", "8 小时"],
            ["装卸/准备小时", "0 小时"],
            ["计算结果", "ceil(25/10)=3 炉；负荷=3×8=24 小时"],
        ]),
        paragraph("注意：近期讨论中已确认后续批量炉应改为投影面积与装载组逻辑，即容器/堆叠产品不能简单按单件有效面积均摊；该细化属于下一步规则，不在当前代码口径中完全实现。"),
    ]))

    parts.append(section("14. 热处/表处流水线处理逻辑（当前代码口径）", [
        paragraph("流水线处理不按炉次计算，而按吞吐率计算周期负荷。单件在炉时间用于报告展示，不作为独占产能负荷。"),
        formula("流水线容量占用 = 产品数量 × 单件容量占用"),
        formula("流水线负荷小时 = 流水线容量占用 / 流水线吞吐率 + 换型时间小时"),
        table(["项目", "数值"], [
            ["产品数量", "200 件"],
            ["单件容量占用", "1 件"],
            ["流水线吞吐率", "50 件/小时"],
            ["换型时间", "0 小时"],
            ["计算结果", "200/50=4 小时"],
        ]),
        paragraph("如果使用面积口径，吞吐率和单件容量占用必须保持同一单位。例如吞吐率为 20 平方米/小时，则单件容量占用也应维护为平方米。"),
    ]))

    parts.append(section("15. ModeA 和 ModeB 的区别", [
        table(["项目", "ModeA", "ModeB"], [
            ["定位", "无限产能压力分析", "100%产能优化建议"],
            ["是否使用 OR-Tools", "不使用", "使用 CP-SAT 整数优化"],
            ["是否考虑可选工序", "不优化选择", "原路径、可选路径、外包一起优化"],
            ["是否允许超过100%", "允许，用于识别瓶颈", "优化后仍可超过100%，但报告输出缺口和建议"],
            ["输出重点", "瓶颈、负荷率、无限产能缺口", "优化后负荷、分流、外包、加班/设备建议、订单工序分配"],
        ]),
        paragraph("例子：某周 A 工作中心产能 100 小时，ModeA 倒排负荷 180 小时，负荷率 180%。ModeB 发现其中 50 小时可转到 B，20 小时可外包，则优化后 A 负荷为 110 小时，仍超 10 小时。报告会显示 A 优化后负荷率 110%，并给出 10 小时的补产能建议。"),
    ]))

    parts.append(section("16. 报告口径和检查方法", [
        table(["报告页", "用途"], [
            ["仪表板", "汇总模式、周期、瓶颈、缺口和运行信息。"],
            ["周度/月度产能分析", "按工作中心和周期展示原始/优化后负荷、产能、缺口和建议。"],
            ["瓶颈分析", "基于负荷率识别瓶颈工作中心。"],
            ["工作组热力图", "以周期为列，按资源组分类展示负荷率和负荷小时。"],
            ["订单工序分配明细", "按订单、物料、活动展示每道工序如何被分配。"],
            ["可选工序分流分析", "只展示发生路径变化或外包的分流行。"],
            ["OR-Tools本次求解规模", "展示变量数、周期数、工作中心数、求解状态和耗时。"],
            ["热处表处产能分析", "专用逻辑下展示炉次、容量占用、吞吐率等周期级指标。"],
        ]),
        paragraph("检查某个工作中心结果是否合理时，可以手工用同周期、同工作中心的优化后负荷小时除以周期产能小时。如果结果大于 1，说明报告中的负荷率超过 100%，缺口应等于优化后负荷小时减周期产能小时。"),
        formula("优化后负荷率 = 优化后负荷小时 / 周期产能小时"),
        formula("优化后缺口小时 = max(优化后负荷小时 - 周期产能小时, 0)"),
    ]))

    parts.append(section("17. 当前模型没有做什么", [
        bullets([
            "不输出车间执行排程，不给出每台设备、每一炉、每一件产品的具体开始结束时间。",
            "不做真实二维装炉摆放，也不判断形状是否一定能摆下。",
            "不做跨周期自动合炉；当前周期内尾数按当前周期负荷体现。",
            "不自动根据 Excel 行顺序决定优先级。",
            "不在当前代码中完整实现升温、保温、降温曲线兼容判断；目前热处/表处兼容更多依赖工艺组或工艺时间字段的后续维护。",
            "不模拟外包供应商产能、外包成本上限或供应商日历。",
        ]),
        paragraph("这些限制并不是错误，而是为了保持工具定位在“产能优化及建议”，避免变成复杂 APS 执行排程系统。"),
    ]))

    return "".join(parts)


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="202020"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:spacing w:after="280"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="46"/><w:color w:val="16365C"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="360" w:after="180"/><w:outlineLvl w:val="0"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="D9EAF7"/></w:pBdr></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="34"/><w:color w:val="1F4E79"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:ind w:left="540" w:hanging="270"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="202020"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Formula">
    <w:name w:val="Formula"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="80" w:after="120"/><w:ind w:left="360"/><w:shd w:fill="F4F6F9"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Microsoft YaHei"/><w:sz w:val="21"/><w:color w:val="404040"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>
  </w:style>
</w:styles>"""


def numbering_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs>
        <w:ind w:left="540" w:hanging="270"/>
      </w:pPr>
      <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>"""


def document_xml(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def write_docx(path: Path) -> None:
    body = build_body()
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>""",
        "word/document.xml": document_xml(body),
        "word/styles.xml": styles_xml(),
        "word/numbering.xml": numbering_xml(),
        "word/settings.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>""",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)


def main() -> None:
    write_docx(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
