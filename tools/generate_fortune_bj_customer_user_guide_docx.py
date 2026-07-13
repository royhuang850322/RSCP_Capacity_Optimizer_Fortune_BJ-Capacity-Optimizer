from __future__ import annotations

import html
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs"
OUT_PATH = OUT_DIR / "Fortune_BJ_客户使用与报告阅读说明_CN_V03.docx"


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
            "<w:pPr><w:pStyle w:val=\"ListBullet\"/>"
            "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            f"{run(item)}</w:p>"
        )
    return "".join(parts)


def table(headers: list[str], rows: list[list[object]], widths: list[int] | None = None) -> str:
    total_width = 9360
    if widths is None:
        base_width = total_width // max(len(headers), 1)
        widths = [base_width for _ in headers]
        if widths:
            widths[-1] = total_width - sum(widths[:-1])
    else:
        if len(widths) != len(headers):
            raise ValueError("width count must match headers")
        if sum(widths) != total_width:
            widths = widths[:]
            widths[-1] += total_width - sum(widths)

    def cell(value: object, width: int, *, header: bool = False) -> str:
        shade = "<w:shd w:fill=\"E8EEF5\"/>" if header else ""
        return (
            f"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/>"
            "<w:vAlign w:val=\"center\"/>"
            f"{shade}</w:tcPr>{paragraph(str(value), style='TableText', bold=header)}</w:tc>"
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


def subsection(title: str, body: list[str]) -> str:
    return paragraph(title, style="Heading2") + "".join(body)


def formula(text: str) -> str:
    return paragraph(text, style="Formula")


def callout(text: str) -> str:
    return paragraph(text, style="Callout")


def build_body() -> str:
    parts: list[str] = []
    parts.append(paragraph("Fortune BJ 产能优化工具客户使用与报告阅读说明", style="Title"))
    parts.append(paragraph("面向对象：计划、生产、工艺、设备和管理人员。用途：维护源数据、运行 ModeA/ModeB、阅读产能报告、追溯订单工序分配结果。", style="Subtitle"))
    parts.append(paragraph("版本：V03 | 生成日期：2026-07-06", style="Subtitle"))

    parts.append(section("1. 工具用途和阅读顺序", [
        paragraph("本工具用于把订单需求、工序路线、工作中心能力、工作日历和可选工序转换成周期级产能分析结果。客户使用时，建议按以下顺序阅读报告。"),
        table(["步骤", "先看哪里", "要判断什么", "正例", "需关注例"], [
            ["1", "仪表板", "确认本次运行模式、分析粒度、订单数、瓶颈数量和剩余缺口。", "ModeB求解状态为成功，剩余缺口为0或较小。", "剩余缺口较大时，继续看超产能解决建议。"],
            ["2", "周度/ 月度产能分析", "确认每个工作中心在每个周期的负荷、产能和缺口。", "负荷率0.85表示85%，产能有余量。", "负荷率1.30表示130%，需要管理动作补足30%的缺口。"],
            ["3", "工作组热力图", "快速找出哪些周期、哪些工作组压力最高。", "某工作组连续多周小于1.00，产能相对平稳。", "某工作组某周大于1.00，应打开产能分析页看具体工作中心缺口小时。"],
            ["4", "订单工序分配明细", "追溯每个订单、每个活动被分到哪个工作中心或外包。", "同一活动5件被分成2件原工作中心、3件可选工作中心。", "某活动全部保留原工作中心，表示没有更优可选路径或可选路径已满。"],
            ["5", "超产能解决建议", "查看仍超过100%的工作中心该补多少产能。", "建议加班20小时，可通过本周加班覆盖。", "建议新增设备数大于0且连续出现时，适合评估设备投资。"],
        ], widths=[700, 1700, 2500, 2200, 2260]),
        callout("客户使用口径：本工具输出产能分析和优化建议，不作为车间逐小时执行排程单。报告中的开始/完成时间用于计算周期负荷和追溯逻辑。"),
    ]))

    parts.append(section("2. 常用概念说明", [
        table(["概念", "客户应如何理解", "正例", "需关注例"], [
            ["周期", "工具按周或月汇总产能。选择周时，报告列出W23等周标签和日期跨度；选择月时，报告按月份汇总。", "W23，日期跨度为2026-06-01至2026-06-07。", "只看W23不看日期跨度时，容易忽略它具体覆盖哪几天。"],
            ["工作中心", "实际消耗产能的资源，例如一类机床、热处理炉或表处线。", "三轴立式加工中心HCMC-1682(北京)。", "工序短文本在工作中心表中找不到时，分析前校验会停止。"],
            ["单位工时", "每件产品在该工序上的标准小时数。普通工序取标准值1、标准值2、标准值3之和。", "标准值1=1.5，标准值2=0.5，标准值3=0，单位工时=2小时/件。", "单位工时为空或维护成0时，该工序负荷会异常偏低。"],
            ["负荷小时", "订单数量转换成工作中心占用小时。", "5件产品，每件2小时，负荷=10小时。", "同样1行工序，数量和单位工时不同，负荷小时会差很多。"],
            ["周期产能小时", "该工作中心在该周期内可用的小时数。", "2台设备、每日24小时、每周7天，周产能=336小时。", "日历名称填错时，工作中心无法正确取得产能。"],
            ["负荷率", "负荷小时除以产能小时。0.80表示80%，1.20表示120%。", "80小时负荷/100小时产能=0.80。", "1.20不是20%，是120%，表示超出产能20小时。"],
            ["缺口小时", "负荷超过产能的部分。", "130小时负荷、100小时产能，缺口=30小时。", "80小时负荷、100小时产能，缺口=0小时。"],
            ["可选工序", "某个物料的某个活动可以转到其他工作中心或外包。", "物料A活动20可从WC1转到WC2。", "物料或活动号对不上时，工具不会把这条可选工序用于该订单。"],
            ["外包", "可选资源组分类为外包时，视为厂外处理，不占用厂内工作中心产能，当前按7天日历返回。", "活动20转外包后，厂内负荷小时为0，并释放原工作中心工时。", "外包单位工时即使填写，也不参与厂内产能计算。"],
        ], widths=[1300, 3000, 2500, 2560]),
    ]))

    parts.append(section("3. 数据导入文件夹中的源数据", [
        paragraph("数据导入文件夹是当前工具的正式数据源。每次运行前，客户应维护以下5个CSV文件。"),
        table(["文件", "主要作用", "关键字段", "正例", "需关注例"], [
            ["订单交期数量_产能分析输入模板.csv", "告诉工具有哪些订单、每个订单多少件、客户需要日期是哪一天。", "富创单据号、华创需求数量、供给日期、紧急类型", "订单1001数量5，供给日期2026-07-15，紧急类型填写越库。", "富创单据号被Excel保存成科学计数法时，工具会停止并提示恢复文本订单号。"],
            ["生产订单工序_排产输入模板.csv", "告诉工具每个订单有哪些活动、物料、工序短文本和标准工时。", "订单、活动、物料、工序短文本、标准值1、标准值2、标准值3", "订单1001有活动10、20、30，每行都维护工序短文本。", "订单在订单交期数量表中存在，但这里没有工序时，分析前校验会报告缺失工序。"],
            ["工作中心_排产输入模板.csv", "告诉工具每个工序短文本对应的工作中心能力。", "工作中心、资源组分类、数量、日历名称", "工作中心A数量2，日历名称为7天24小时。", "生产订单工序中的工序短文本没有在这里维护时，校验报告会列出缺失工作中心。"],
            ["工作日历_排产输入模板.csv", "告诉工具不同日历每天工作多少小时、每周工作几天。", "日历名称、每日工作小时、每周工作天数", "7天24小时：每日24，每周7天。", "日历名称与工作中心表不一致时，对应工作中心无法参与计算。"],
            ["可选工序_排产输入模板.csv", "告诉ModeB某些物料活动可转移到哪个工作中心或外包。", "物料、活动、可选工作中心、可选资源组分类、工序优先级、可选单位工时", "物料A活动20可选WC2，单位工时12小时/件。", "外包行以可选资源组分类=外包判断，可选单位工时可空且会被忽略。"],
        ], widths=[2100, 2300, 2300, 1400, 1260]),
    ]))

    parts.append(subsection("3.1 订单交期数量表的读取规则", [
        paragraph("订单交期数量表的供给日期作为客户需求日期。同一富创单据号出现多行时，工具会合并为一个订单需求。"),
        formula("同一订单数量 = 相同富创单据号的华创需求数量合计"),
        formula("同一订单需求日期 = 相同富创单据号中最早的供给日期"),
        table(["输入示例", "数量", "供给日期", "工具处理结果"], [
            ["订单1001第1行", 2, "2026-07-10", "参与合并"],
            ["订单1001第2行", 3, "2026-07-20", "参与合并"],
            ["合并后", 5, "2026-07-10", "数量=5，需求日期取最早供给日期"],
        ], widths=[1800, 1100, 1800, 4660]),
        table(["场景", "工具处理", "客户应检查"], [
            ["供给日期为2026-07-10", "作为正常需求日期进入分析。", "确认该日期是否为客户要求的满足日期。"],
            ["供给日期为2049年或以后", "视为占位交期，不进入正式产能分析。", "在“占位交期订单”页检查并决定是否修正日期。"],
            ["ModeB中供给日期早于优化开始周期", "转入优化开始周期，并自动标记为紧急订单。", "确认历史未完订单是否需要从本期开始消化。"],
        ], widths=[2300, 3400, 3660]),
    ]))

    parts.append(subsection("3.2 生产订单工序表的读取规则", [
        paragraph("生产订单工序表决定订单有哪些工序，以及每道工序的基础负荷。"),
        formula("普通工序单位工时 = 标准值1 + 标准值2 + 标准值3"),
        formula("普通工序负荷小时 = 订单数量 × 普通工序单位工时"),
        table(["示例", "订单数量", "标准值1", "标准值2", "标准值3", "计算结果"], [
            ["活动20", 5, 1.5, 0.5, 0, "单位工时=2；负荷=5×2=10小时"],
            ["活动30", 2, 8, 0, 0, "单位工时=8；负荷=2×8=16小时"],
        ], widths=[1200, 1200, 1200, 1200, 1200, 3360]),
        table(["数据状态", "工具处理", "客户应检查"], [
            ["订单1001在订单表和工序表都存在", "进入正式分析。", "检查活动号顺序、物料和工序短文本是否正确。"],
            ["订单1002在订单表存在，工序表没有任何行", "分析前校验报告列为缺失工序，正式分析停止。", "补充该订单完整工序，或从订单需求中移除不参与分析的订单。"],
        ], widths=[2700, 3400, 3260]),
    ]))

    parts.append(subsection("3.3 工作中心和日历的产能计算", [
        paragraph("工作中心表提供设备数量，工作日历表提供工作时间。工具把两者合并成周期产能小时。"),
        formula("平均每日小时/台 = 每日工作小时 × 每周工作天数 / 7"),
        formula("周期产能小时 = 设备数量 × 平均每日小时/台 × 周期天数"),
        table(["示例", "设备数量", "每日工作小时", "每周工作天数", "周期", "周期产能小时"], [
            ["7天24小时", 2, 24, 7, "1周7天", "2×24×7=336"],
            ["5天16小时", 1, 16, 5, "1周7天", "1×(16×5/7)×7=80"],
        ], widths=[1500, 1100, 1500, 1500, 1400, 2360]),
        callout("报告中的“平均每日小时/台”已经把每周工作天数折算进去。客户核对产能时，应使用报告中的平均每日小时/台和周期天数。"),
    ]))

    parts.append(subsection("3.4 可选工序和外包维护规则", [
        paragraph("ModeB会根据可选工序表，为同一物料、同一活动寻找可转移路径。匹配口径是“物料 + 活动”。"),
        table(["可选工序维护", "工具处理", "示例"], [
            ["可选资源组分类不是外包", "可选工作中心必须在工作中心表中存在；可选单位工时用于计算新负荷。", "物料A活动20从WC1转到WC2，5件×12小时/件=60小时加到WC2。"],
            ["可选资源组分类=外包", "视作无限厂外产能；厂内负荷为0；释放原工作中心工时；当前按7天日历返回。", "物料A活动20原负荷50小时，转外包后WC1释放50小时。"],
            ["同一活动有多个可选路径", "ModeB在周期内按整数产品数量分配。", "5件中2件留WC1、2件转WC2、1件外包。"],
        ], widths=[2500, 3900, 2960]),
    ]))

    parts.append(section("4. 运行模式怎么选", [
        table(["模式", "客户使用场景", "主要输出", "正例", "需关注例"], [
            ["ModeA 无限产能分析", "先看真实需求按交期倒排后，哪些工作中心压力最大。", "瓶颈、无限产能负荷、热力图、订单工序明细。", "用于月度/周度产能压力盘点。", "负荷率超过100%是压力信号，不代表已完成有限产能优化。"],
            ["ModeB 100%产能优化建议", "在ModeA压力基础上，使用可选工序和外包做周期级优化建议。", "优化后负荷、分流明细、外包释放、超产能解决建议。", "用于比较可选路径、外包和补产能动作。", "优化后仍超过100%时，报告会保留实际负荷率并给出解决建议。"],
        ], widths=[1800, 2500, 2300, 1400, 1360]),
        paragraph("ModeB参数中，优化粒度可选周或月。选择周时，报告按周展示；选择月时，报告按月展示。优化开始周期用于处理历史未完订单：早于开始周期的订单会转入开始周期计算。"),
    ]))

    parts.append(section("5. 正常运行前的数据校验", [
        paragraph("工具正式计算前会先做数据完整性校验。校验通过后才开始产能分析；校验不通过时，会生成分析前数据校验报告。"),
        table(["校验项", "通过状态", "未通过时的报告提示", "客户处理方法"], [
            ["订单是否都有工序", "订单交期数量表中的有效订单，都能在生产订单工序表中找到。", "缺失工序订单数量、示例订单、源文件行号。", "补充该订单工序，或确认该订单不参与分析后从订单表移除。"],
            ["工序是否都有工作中心", "参与分析的工序短文本，都能在工作中心表中找到。", "缺失工作中心工序短文本、涉及订单数、示例物料。", "在工作中心表补充对应工作中心、资源组、数量和日历。"],
            ["工作中心是否有日历", "工作中心表中的日历名称，都能在工作日历表中找到。", "日历名称缺失或无效。", "统一日历名称，或在工作日历表新增该日历。"],
        ], widths=[1800, 3100, 2400, 2060]),
    ]))

    parts.append(section("6. 报告页作用说明", [
        paragraph("以下说明按当前工具生成的Excel报告页编写。部分页面只在ModeB、热处/表处专用逻辑或存在异常数据时出现。"),
        table(["报告页", "出现模式", "主要作用", "客户怎么读", "需关注例"], [
            ["仪表板", "ModeA/ModeB", "汇总本次运行的核心指标。", "先确认模式、优化粒度、瓶颈数量、剩余缺口、数据质量记录数。", "ModeB求解状态不是成功时，应看OR-Tools本次求解规模。"],
            ["周度/ 月度产能分析", "ModeA/ModeB", "按工作中心和周期展示负荷、产能、负荷率和缺口。", "优先筛选负荷率大于1或缺口小时大于0的行。", "ModeB中优化后仍大于1时，应联动看超产能解决建议。"],
            ["瓶颈分析", "ModeA/ModeB", "列出周期内负荷率较高的工作中心。", "按负荷率从高到低看瓶颈排序。", "一个工作中心只在某一周超载时，应结合日期跨度判断。"],
            ["工作组热力图", "ModeA/ModeB", "横向查看工作组在多个周期的压力变化。", "找出连续高负荷或尖峰周期。", "只看颜色或数字时，要同时看指标行是负荷率、负荷小时还是缺口小时；具体工作中心仍看周度/月度产能分析。"],
            ["订单工序分配明细", "ModeA/ModeB", "追溯每个订单每个活动的工作中心、数量、工时和优化动作。", "按订单筛选，看活动是否被保留、转可选工作中心或转外包。", "同一订单活动出现多行时，表示按整数件拆到多个路径。"],
            ["占位交期订单", "有2049年及以后供给日期时", "列出未进入产能分析的占位交期订单。", "检查这些订单是否需要改成真实供给日期。", "占位需求数量较大时，会影响真实产能判断，应先修正。"],
            ["ModeB优化周期明细", "ModeB", "查看每个优化周期的订单数、工序数、外包数量、总负荷和总缺口。", "确认每个周期是否都有数据，以及求解状态是否正常。", "某周期工序数很大时，求解耗时可能增加。"],
            ["100%产能优化总览", "ModeB", "按周期和工作中心对比优化前后负荷、缺口和状态。", "看缺口改善小时和优化后缺口小时。", "状态仍超100%时，需要用建议页制定动作。"],
            ["可选工序分流分析", "ModeB且有分流时", "只列出发生可选工作中心或外包变化的工序。", "看原工作中心减少多少小时、建议工作中心增加多少小时。", "建议工作中心也接近满负荷时，应看产能分析页确认是否形成新瓶颈。"],
            ["超产能解决建议", "ModeB", "把优化后仍超过100%的缺口转换成外包、加班、设备建议。", "按工作中心和周期落实临时或长期动作。", "新增设备建议适合连续多周期缺口，不适合只看单次尖峰。"],
            ["OR-Tools本次求解规模", "ModeB", "记录本次优化引擎处理的数据规模、状态和耗时。", "看求解状态、候选方案数、求解耗时秒、总短缺小时。", "状态失败时，本次结果应按提示谨慎使用。"],
            ["热处表处产能分析", "热处/表处专用逻辑", "汇总批量处理或流水线处理的容量、炉次、吞吐率和负荷。", "批量炉看容量占用和折算炉次；流水线看吞吐率和负荷小时。", "该页是周期容量分析，不是每炉执行顺序。"],
            ["输入字段维护说明", "ModeA/ModeB", "说明热处/表处相关扩展字段如何维护。", "用于检查输入文件是否缺少关键字段。", "专用逻辑下批量炉或流水线字段为空时，结果会回到普通或不完整口径。"],
            ["缺失映射报告", "存在缺失工作中心时", "汇总生产订单工序中找不到工作中心映射的工序短文本。", "按工序短文本补维护工作中心表。", "缺失映射未处理时，正式分析不完整。"],
            ["数据质量报告", "存在数据调整或异常时", "列出过期订单转入优化开始周期、可选工序异常等记录。", "逐行确认数据是否符合业务预期。", "过期订单数量大时，会集中进入优化开始周期。"],
            ["运行信息", "ModeA/ModeB", "记录本次使用的输入文件路径、模式、参数、license信息和统计信息。", "用于复盘本次报告的来源和参数。", "不同报告对比时，应先确认运行信息一致。"],
        ], widths=[1700, 1450, 2150, 2200, 1860]),
    ]))

    parts.append(section("7. 订单工序分配明细怎么读", [
        paragraph("订单工序分配明细是客户追溯最重要的页面。ModeA中它展示原始倒排后的工序负荷；ModeB中它展示优化后每个订单活动被分配到哪里。"),
        callout("ModeB阅读重点：同一订单、同一物料、同一活动如果出现多行，代表该活动按整数件分配到多个路径。每行的“本行分配产品数量”相加，应等于“原工序总产品数量”。"),
    ]))

    parts.append(subsection("7.1 ModeB订单工序分配明细字段", [
        table(["字段", "含义和计算口径", "示例"], [
            ["订单", "来自订单交期数量表的富创单据号，并与生产订单工序表中的订单匹配。", "100000123456"],
            ["物料", "该工序对应的物料，用于匹配可选工序表。", "MAT-A"],
            ["活动", "工序活动号。ModeB用物料+活动匹配可选路径。", "20"],
            ["工序短文本", "生产订单工序中的工序说明，也用于匹配工作中心表。", "三轴加工"],
            ["源文件行号", "该工序在生产订单工序源文件中的行号，便于回查。", "358"],
            ["需求日期", "订单交期数量表中该订单最早供给日期；过期订单在ModeB中可能转到优化开始周期。", "2026-07-15"],
            ["是否紧急", "订单本身标记为紧急，或ModeB中过期订单转入优化开始周期后自动标记为紧急。", "是"],
            ["周期", "该工序进入分析的周或月。", "2026-W28"],
            ["周期日期跨度", "周期覆盖的具体日期范围。", "2026-07-06 至 2026-07-12"],
            ["原工序总产品数量", "该工序对应的订单产品数量。", "5"],
            ["本行分配产品数量", "本行实际分配到某个路径的整数件数。", "2"],
            ["同工序分配行", "同一工序被拆成几行，以及当前是第几行。", "1/3表示该工序共有3条分配行。"],
            ["原工作中心", "ModeA原始路径上的工作中心。", "WC1"],
            ["原资源组分类", "原工作中心所属资源组。", "机加工"],
            ["原单位工时", "原路径每件产品的单位小时数。", "10小时/件"],
            ["原路径负荷小时(按本行数量)", "本行数量按原路径计算会占用的小时数。计算：本行分配产品数量×原单位工时。", "2件×10=20小时"],
            ["优化动作", "ModeB给出的路径动作：保留原工作中心、转可选工作中心、转外包。", "转可选工作中心"],
            ["优化后工作中心", "本行产品最终建议占用的工作中心；外包时显示外包。", "WC2"],
            ["优化后资源组分类", "优化后工作中心所属资源组；外包时显示外包。", "机加工备用"],
            ["优化后单位工时", "本行优化后路径的单位小时数；外包行按0显示。", "12小时/件"],
            ["产能计算类型", "普通工时、批量处理、流水线处理或外包。", "普通工时"],
            ["热处/表处类型", "热处、表处或普通，用于热处表处报告分组。", "热处"],
            ["工艺兼容组", "热处/表处专用逻辑下的工艺分组字段。", "HRC45"],
            ["单件容量占用", "热处/表处专用逻辑下，每件产品占用的容量单位。", "0.2平方米"],
            ["容量占用", "本行数量折算后的容量占用。普通工时下可为空或0。", "10件×0.2=2平方米"],
            ["单炉容量", "批量处理工作中心每炉可用容量。", "5平方米"],
            ["单炉周期小时", "批量处理每炉占用的周期小时。", "8小时"],
            ["折算炉次", "批量处理下按容量占用向上取整得到的炉次数。", "容量占用12、单炉容量5，折算炉次=3"],
            ["流水线吞吐率", "流水线处理下每小时可处理的容量或件数。", "50件/小时"],
            ["单件在炉时间小时", "单件从进入流水线到出来的时间，用于提示和报告展示。", "2小时"],
            ["优化后厂内负荷小时", "本行优化后实际占用厂内工作中心的小时。外包行为0。", "2件×12=24小时；外包=0小时"],
            ["原工作中心释放小时", "本行从原工作中心释放出来的小时。保留原工作中心时为0。", "原路径20小时转走，则释放20小时"],
            ["外包释放本厂工时", "本行转外包时释放的厂内原工时。非外包为0。", "5件×10=50小时"],
            ["额外工时", "可选工作中心比原路径多出的厂内小时。计算：max(优化后厂内小时-原路径小时,0)。", "原20小时，优化后24小时，额外工时=4"],
            ["是否外包", "本行是否被分配到外包路径。", "是/否"],
            ["外包返回日历天", "外包路径当前固定按7天日历返回。", "7"],
            ["说明", "提示该工序是否被按整数件拆到多个路径。", "最小分配单位为1件产品；同一工序可按整数件拆到多个路径"],
        ], widths=[1900, 5200, 2260]),
    ]))

    parts.append(subsection("7.2 ModeB订单工序分配明细阅读例子", [
        table(["场景", "订单工序分配明细显示", "客户解读"], [
            ["全部保留原工作中心", "原工序总产品数量5，本行分配产品数量5，优化动作=保留原工作中心，优化后工作中心=WC1。", "该工序5件全部留在原工作中心。"],
            ["部分转可选工作中心", "同一订单活动出现2行：3件保留WC1，2件转WC2。", "该工序按整数件拆分，WC1保留3件，WC2承接2件。"],
            ["部分外包", "同一订单活动出现2行：4件保留WC1，1件外包；外包行优化后厂内负荷小时=0。", "1件不占用厂内产能，并释放原工作中心对应工时。"],
            ["优化后仍超产能", "某工作中心仍有优化后缺口小时。", "该工作中心的可选路径和外包已经使用后仍有缺口，需看超产能解决建议。"],
        ], widths=[2200, 4200, 2960]),
    ]))

    parts.append(subsection("7.3 ModeA订单工序分配明细字段", [
        paragraph("ModeA中的订单工序分配明细用于查看原始倒排后的工序负荷，不包含可选工序优化动作。"),
        table(["字段", "含义", "示例"], [
            ["订单/活动/物料/工序短文本", "来自生产订单工序表，用于定位工序。", "订单1001，活动20，物料A。"],
            ["工作中心/资源组分类", "该工序对应的原工作中心和资源组。", "WC1，机加工。"],
            ["订单数量", "该订单合并后的需求数量。", "5"],
            ["单位工时(小时/pcs)", "标准值1+标准值2+标准值3。", "2小时/件"],
            ["工序生产时间(小时)", "订单数量×单位工时。", "5×2=10小时"],
            ["开始时间/完成时间", "ModeA倒排计算出的时间，用于确定负荷落入哪个周期。", "完成时间接近需求日期。"],
            ["需求日期", "订单供给日期。", "2026-07-15"],
            ["是否紧急/是否热处/表处/是否外协", "用于筛选订单和工艺类型。", "是否紧急=是。"],
            ["分析口径/分析来源/窗口编号/窗口类型/说明", "说明该行来自哪种分析逻辑。", "ModeA无限产能倒排。"],
        ], widths=[2400, 4700, 2260]),
    ]))

    parts.append(section("8. 产能分析页的关键字段", [
        table(["字段", "含义和计算", "正例", "需关注例"], [
            ["周期/周期日期跨度", "表示该行统计的是哪一周或哪一月。", "W28，2026-07-06至2026-07-12。", "跨月周应按日期跨度判断，不只看周号。"],
            ["周期产能小时", "该工作中心本周期可用小时。", "100小时。", "为0时通常是工作中心或日历维护异常。"],
            ["原始负荷小时或无限产能负荷小时", "ModeA倒排得到的原始压力。", "180小时。", "超过产能表示该周期压力大。"],
            ["原始负荷率或无限产能负荷率", "原始负荷小时÷周期产能小时。", "180÷100=1.80。", "1.80表示180%。"],
            ["优化后负荷小时", "ModeB分流和外包后仍占用该工作中心的小时。", "优化后110小时。", "仍大于产能时，应看优化后缺口小时。"],
            ["优化后负荷率", "优化后负荷小时÷周期产能小时。", "110÷100=1.10。", "ModeB中该值允许超过1，超过部分会进入建议页。"],
            ["优化后缺口小时", "max(优化后负荷小时-周期产能小时,0)。", "110-100=10小时。", "缺口为0表示该周期该工作中心已在100%以内。"],
            ["缺口改善小时", "原始缺口小时-优化后缺口小时。", "原缺口80，优化后缺口10，改善70小时。", "为0表示本周期没有可用分流或分流效果有限。"],
        ], widths=[2200, 3400, 1800, 1960]),
    ]))

    parts.append(section("9. ModeB优化结果的业务含义", [
        paragraph("ModeB会在同一个周期内，根据原工作中心、可选工作中心和外包路径，为每个工序分配整数件产品。"),
        formula("每道工序：各路径分配产品数量合计 = 原工序总产品数量"),
        formula("优化后缺口小时 = max(优化后负荷小时 - 周期产能小时, 0)"),
        table(["业务场景", "优化结果", "客户动作"], [
            ["可选工作中心有余量", "部分产品从原工作中心转到可选工作中心，原工作中心释放小时。", "确认可选工作中心工艺可执行，并按建议评估分流。"],
            ["可选工作中心已接近满负荷", "只有部分产品转移，剩余产品保留原工作中心或转外包。", "查看可选工作中心是否形成新瓶颈。"],
            ["外包可用", "多余产能可转外包，厂内负荷为0，释放原工作中心小时。", "评估外包供应、质量和交付风险。"],
            ["优化后仍超100%", "报告保留实际负荷率和缺口小时。", "按建议页落实加班、外包、临时班次或设备投资。"],
        ], widths=[2600, 3800, 2960]),
    ]))

    parts.append(section("10. 热处/表处专用逻辑怎么读", [
        table(["处理方式", "计算口径", "例子", "报告查看点"], [
            ["同机加逻辑", "按订单数量×单位工时计算。", "20件×0.5小时/件=10小时。", "订单工序分配明细和产能分析页。"],
            ["批量处理", "按容量占用折算炉次，再按每炉周期小时计算负荷。", "容量占用25，单炉容量10，炉次=3；每炉8小时，负荷=24小时。", "热处表处产能分析页看容量占用、折算炉次合计、负荷小时。"],
            ["流水线处理", "按容量占用÷流水线吞吐率，再加换型时间。", "200件，吞吐率50件/小时，负荷=4小时。", "热处表处产能分析页看吞吐率、单件在炉时间、负荷小时。"],
        ], widths=[1700, 3000, 2600, 2060]),
        callout("热处表处产能分析页是周期容量分析，用于看本周期炉子或流水线产能是否足够；它不是每炉装炉清单，也不是每件产品的执行顺序。"),
    ]))

    parts.append(section("11. 常见检查场景", [
        table(["客户问题", "先看报告页", "判断方法", "处理建议"], [
            ["某个工作组本周压力是否过大", "工作组热力图、周度产能分析", "先看工作组热力图，再到周度产能分析筛选组内工作中心。", "缺口大于0时，进入超产能解决建议。"],
            ["某个订单被分到哪里", "订单工序分配明细", "筛选订单号，按活动号查看优化动作和优化后工作中心。", "对转外包或转可选工作中心的行做工艺确认。"],
            ["可选工序是否发挥作用", "可选工序分流分析", "看原工作中心减少负荷小时和建议工作中心增加负荷小时。", "若无数据，检查可选工序表物料和活动是否匹配。"],
            ["历史过期订单怎么处理", "数据质量报告、订单工序分配明细", "看是否出现过期订单转入优化开始周期。", "确认优化开始周期设置是否符合本次分析目的。"],
            ["报告中有订单缺失", "分析前数据校验报告", "看缺失工序订单或缺失工作中心工序短文本。", "先补源数据，再重新运行。"],
            ["需要解释新增设备建议", "超产能解决建议、100%产能优化总览", "看该工作中心优化后缺口小时和负荷率。", "单次尖峰可优先用外包或加班；连续多周期再评估设备。"],
        ], widths=[2300, 2300, 2700, 2060]),
    ]))

    parts.append(section("12. 使用前检查清单", [
        bullets([
            "订单交期数量表中的富创单据号保持文本格式，供给日期为真实客户需求日期。",
            "生产订单工序表中，每个有效订单都有完整活动和工序短文本。",
            "工作中心表中，每个工序短文本对应的工作中心、数量、资源组和日历都已维护。",
            "工作日历表中的日历名称与工作中心表完全一致。",
            "可选工序表中的物料和活动号与生产订单工序表一致。",
            "外包行使用“可选资源组分类=外包”，外包单位工时可不填。",
            "ModeB运行前确认优化粒度为周或月，并确认优化开始周期。",
            "运行完成后先看仪表板，再看产能分析和订单工序分配明细。",
        ]),
    ]))

    return "".join(parts)


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="202020"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="160"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="42"/><w:color w:val="1F4D78"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="100" w:line="280" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="21"/><w:color w:val="555555"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="360" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="32"/><w:color w:val="2E74B5"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="26"/><w:color w:val="2E74B5"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:ind w:left="540" w:hanging="270"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="202020"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Formula">
    <w:name w:val="Formula"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="80" w:after="100"/><w:ind w:left="240"/><w:shd w:fill="F4F6F9"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Microsoft YaHei"/><w:sz w:val="21"/><w:color w:val="404040"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Callout">
    <w:name w:val="Callout"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="80" w:after="140"/><w:shd w:fill="F4F6F9"/><w:ind w:left="160" w:right="160"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="1F3A5F"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TableText">
    <w:name w:val="Table Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="40" w:line="280" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="19"/><w:color w:val="202020"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>
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
      <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/></w:rPr>
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
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="708" w:footer="708" w:gutter="0"/>
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
