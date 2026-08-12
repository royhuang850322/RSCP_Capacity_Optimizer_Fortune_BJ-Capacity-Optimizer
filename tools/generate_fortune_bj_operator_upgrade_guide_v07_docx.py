from __future__ import annotations

import html
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT_PATH = ROOT / "docs" / "Fortune_BJ_操作员升级说明与仪表盘阅读指引_CN_V07.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def run(text: str, *, bold: bool = False, italic: bool = False, color: str | None = None) -> str:
    props: list[str] = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def paragraph(
    text: str = "",
    *,
    style: str | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    keep_next: bool = False,
) -> str:
    ppr_parts: list[str] = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if keep_next:
        ppr_parts.append("<w:keepNext/>")
    ppr = f"<w:pPr>{''.join(ppr_parts)}</w:pPr>" if ppr_parts else ""
    content = run(text, bold=bold, italic=italic, color=color) if text else ""
    return f"<w:p>{ppr}{content}</w:p>"


def rich_paragraph(parts: list[tuple[str, bool]], *, style: str | None = None) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return "<w:p>" + ppr + "".join(run(text, bold=bold) for text, bold in parts) + "</w:p>"


def page_break() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def bullets(items: list[str]) -> str:
    return "".join(
        "<w:p><w:pPr><w:pStyle w:val=\"ListBullet\"/>"
        "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
        f"{run(item)}</w:p>"
        for item in items
    )


def numbered(items: list[str]) -> str:
    return "".join(
        "<w:p><w:pPr><w:pStyle w:val=\"ListNumber\"/>"
        "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr>"
        f"{run(item)}</w:p>"
        for item in items
    )


def table(headers: list[str], rows: list[list[object]], widths: list[int]) -> str:
    if len(headers) != len(widths):
        raise ValueError("width count must match headers")
    widths = list(widths)
    widths[-1] += 9360 - sum(widths)

    def cell(value: object, width: int, *, header: bool = False) -> str:
        shade = '<w:shd w:fill="E8EEF5"/>' if header else ""
        return (
            f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
            f'<w:vAlign w:val="center"/>{shade}</w:tcPr>'
            f'{paragraph(str(value), style="TableText", bold=header)}</w:tc>'
        )

    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in widths)
    out = [
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
        '<w:tblW w:w="9360" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/></w:tblPr>',
        f"<w:tblGrid>{grid}</w:tblGrid>",
        "<w:tr>" + "".join(cell(value, widths[index], header=True) for index, value in enumerate(headers)) + "</w:tr>",
    ]
    for row in rows:
        padded = list(row) + [""] * max(len(headers) - len(row), 0)
        out.append("<w:tr>" + "".join(cell(value, widths[index]) for index, value in enumerate(padded[: len(headers)])) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def callout(label: str, text: str) -> str:
    return rich_paragraph([(f"{label}  ", True), (text, False)], style="Callout")


def h1(title: str) -> str:
    return paragraph(title, style="Heading1", keep_next=True)


def h2(title: str) -> str:
    return paragraph(title, style="Heading2", keep_next=True)


def build_body() -> str:
    parts: list[str] = []
    parts.append(paragraph("操作员升级指引", style="Kicker"))
    parts.append(paragraph("Fortune BJ 产能优化工具", style="Title"))
    parts.append(paragraph("2026年8月11日至8月12日升级说明与仪表盘阅读指引", style="Subtitle"))
    parts.append(table(
        ["适用版本", "适用人员", "更新日期"],
        [["v0.1.7", "计划、生产、工艺、设备、外协及管理人员", "2026-08-12"]],
        [1800, 4860, 2700],
    ))
    parts.append(callout(
        "先看这里",
        "本次升级的重点是让预测需求能够按BOM层级进入产能分析、让数据缺口在仪表盘上直接可见、让热力图按客户习惯展示需求工时与设备需求数量，并让工具在下次打开时继续使用上次配置。工具输出用于产能评估和改善讨论，不作为车间执行排产单。",
    ))

    parts.append(h1("1. 两天内完成了哪些升级"))
    parts.append(table(
        ["日期", "升级内容", "操作员能看到的变化"],
        [
            ["8月11日", "工作组热力图升级", "热力图按工作中心表中的“大类”分区，再按资源组展示需求工时、理论工时、负荷和设备需求数量。"],
            ["8月11日", "三地BOM接入", "预测需求先在北京、沈阳、南通BOM中查找顶层物料，并展开需要制造的下层M类物料。"],
            ["8月11日", "预测制造顺序", "预测BOM采用整批齐套方式，下层制造件先完成，再进入上层制造件，最后支持顶层物料。"],
            ["8月12日", "BOM后匹配工艺路线", "先确认预测物料的BOM制造范围，再为顶层和下层制造件分别匹配三地工艺路线。"],
            ["8月12日", "缺失数据不再阻断整次分析", "未找到BOM或工艺路线的预测物料会被记录；可计算的其他订单和预测物料继续运行。"],
            ["8月12日", "仪表盘增加预测完整性数据", "首页直接显示预测顶层物料、未匹配BOM、未匹配工艺路线、下层制造件缺路线和可计算虚拟订单数量。"],
            ["8月12日", "GUI记住上次配置", "关闭工具时保存模式、路径、输出目录、流转方式、主题等设置；下次打开自动恢复。优化开始日期始终更新为当天。"],
            ["8月12日", "订单明细增加BOM追溯字段", "ModeA和ModeB的订单工序分配明细都可按BOM根订单、顶层物料、层级和父项筛选。"],
        ],
        [1200, 2600, 5560],
    ))

    parts.append(h1("2. 升级后的日常操作"))
    parts.append(numbered([
        "把本次需要分析的订单、生产订单工序、工作中心、工作日历和可选工序文件放入“数据导入”文件夹。",
        "需要分析预测时，维护“需求预测_产能分析输入模板.xlsx”，并在首页勾选“需求预测导入”；不需要预测时取消勾选。",
        "确认三份BOM文件和三份工艺路线文件位于“数据导入”文件夹，文件名与sheet名称保持不变。",
        "打开工具，确认分析模式、周/月粒度、工序流转方式和输出目录。上次配置会自动恢复。",
        "确认“优化开始日期”。每次启动默认是当天；本次要从其他日期开始时，再手动修改。",
        "运行ModeA查看原始需求压力；需要评估可选工序和外包改善时，再运行ModeB。",
        "运行完成后点击“马上打开生成的报告”，先看仪表板，再看工作组热力图和订单工序分配明细。",
    ]))
    parts.append(callout("配置提醒", "工具会记住上次输入文件路径和分析选项，但不会记住上次的优化开始日期。这样可以避免隔天运行时继续使用旧日期。"))

    parts.append(h1("3. 本次新增或需要维护的输入"))
    parts.append(table(
        ["文件或字段", "操作员维护内容", "使用说明"],
        [
            ["工作中心_产能分析输入模板.xlsx / E列“大类”", "为每个工作中心填写机加、表处、热处等大类。", "工作组热力图按该字段分区；空白会显示为“未分类”，同一资源组出现多个大类会显示“大类冲突”。"],
            ["北京BOM.xlsx", "sheet保持为Data。", "预测顶层物料优先在北京BOM中查找。"],
            ["沈阳BOM.xlsx", "sheet保持为Data。", "北京未找到时继续在沈阳BOM中查找。"],
            ["南通BOM.xlsx", "sheet保持为Data。", "北京和沈阳均未找到时继续在南通BOM中查找。"],
            ["工艺路线-北京.xlsx", "sheet保持为Data，维护物料、工序编码和标准工时。", "预测物料工艺路线优先使用北京。"],
            ["工艺路线-沈阳.xlsx", "sheet保持为Data。", "北京未匹配时使用沈阳。"],
            ["工艺路线-南通.xlsx", "sheet保持为Data。", "北京和沈阳未匹配时使用南通。"],
        ],
        [2800, 3000, 3560],
    ))
    parts.append(paragraph("预测BOM中的顶层物料始终作为制造件；下层只有物料编码以M开头的项目进入制造产能分析。其他下层物料不生成制造工序。"))

    parts.append(h2("3.1 预测需求的结果怎么理解"))
    parts.append(table(
        ["情况", "工具处理", "操作员下一步"],
        [
            ["顶层物料找到BOM和工艺路线", "顶层物料和可匹配路线的下层M制造件进入计算。", "在订单工序分配明细按预测虚拟订单号或BOM根订单查看完整链条。"],
            ["顶层物料未找到BOM，但找到自身工艺路线", "不展开下层制造件，顶层物料自身继续参与计算。", "在数据质量/调整记录中确认是否需要补BOM。"],
            ["顶层物料未找到工艺路线", "该顶层预测物料不生成工序，不参与本次负荷。", "查看仪表盘“预测顶层物料未匹配工艺路线数”，再到数据质量页取得物料清单。"],
            ["某个BOM下层M制造件未找到工艺路线", "只跳过该制造件；同一BOM内其他已匹配物料继续计算。", "查看“BOM下层制造件未匹配工艺路线数”，安排工艺路线维护。"],
        ],
        [2500, 3500, 3360],
    ))

    parts.append(page_break())
    parts.append(h1("4. 仪表盘怎么读"))
    parts.append(paragraph("仪表盘用于确认本次运行范围、核心数量、重点瓶颈和报告口径。建议每次打开报告先检查仪表盘，确认模式、日期和数据完整性，再进入其他sheet。"))

    parts.append(h2("4.1 顶部运行摘要"))
    parts.append(table(
        ["显示项", "含义", "阅读示例"],
        [
            ["模式", "本次报告是ModeA还是ModeB。", "ModeA表示查看原始需求压力；ModeB表示查看可选路径优化后的产能建议。"],
            ["运行时间", "本次报告实际生成时间。", "用于区分同一天多次运行的报告。"],
            ["优化粒度", "报告按周还是按月汇总。", "选择周时，热力图横向显示周及日期跨度。"],
            ["优化开始日期", "订单链条不得早于的分析起点。", "显示2026-08-12，表示早于该日的未完成链条整体后移到该日开始。"],
            ["工序流转", "本次使用整批、半批、单件流或交期强制。", "整批表示本批完成后再进入下一工序。"],
            ["需求预测", "本次是否把预测需求与真实订单一起计算。", "显示“启用”时，仪表盘会出现预测完整性数量。"],
            ["参考工序数/周期", "用于观察每个周/月的模型规模。", "这是参考提示，不是自动跳过阈值。"],
            ["工作日历", "本次计算产能使用的日历文件。", "运行前确认文件名是当前有效日历。"],
        ],
        [2200, 4100, 3060],
    ))

    parts.append(h2("4.2 核心指标：订单与产能范围"))
    parts.append(table(
        ["指标", "表示什么", "怎么用"],
        [
            ["分析订单数", "本次实际生成工序并进入分析的订单号数量，包含可计算的真实订单和预测虚拟订单。", "与输入订单数不同是正常的；未匹配路线、占位交期或无有效需求的记录不会计入。"],
            ["厂内产能占用工序数", "本次占用厂内工作中心产能的工序数量。", "用于判断本次分析规模；它不是产品数量，也不是报告明细拆分行数。"],
            ["外协模拟工序数", "本次被分配到外包路径、不占厂内工作中心产能的工序数量。", "ModeB出现较大数量时，继续看可选工序分流和外包相关明细。"],
            ["瓶颈工作中心数", "当前模式下至少有一个周期负荷超过产能的工作中心数量。", "数量大时，先看重点瓶颈和工作组热力图。"],
            ["周度/月度产能分析行数", "产能分析sheet中的工作中心与周期组合行数。", "用于确认报告规模，不代表订单或工序数量。"],
        ],
        [2700, 3950, 2710],
    ))

    parts.append(h2("4.3 核心指标：ModeB优化结果"))
    parts.append(table(
        ["指标", "表示什么", "怎么用"],
        [
            ["可选工序分流数", "ModeB中实际形成可选工作中心或外包分流的记录数。", "数量为0时，可能是没有瓶颈、没有匹配可选路径，或优化后仍选择原路径。"],
            ["ModeB整数分配候选工序", "具有两个或以上可选路径、交给OR-Tools比较的工序数量。", "它反映优化选择规模，不等于全部工序数。"],
            ["ModeB整数分配求解状态", "OR-Tools本次求解结果。", "OPTIMAL表示找到最优结果；FEASIBLE表示找到可用结果；SKIPPED表示本次没有需要求解的候选，并非运行错误。"],
            ["ModeB产能缺口改善", "优化前缺口减去优化后缺口的小时数。", "数值越大，表示可选工序和外包缓解的产能压力越多。"],
            ["ModeB剩余产能缺口", "可选路径优化后仍超过本周期理论产能的小时数合计。", "不为0时继续查看“超产能解决建议”。"],
            ["超产能解决建议", "报告中针对剩余缺口生成的外包、加班、排班或设备建议记录数。", "先按缺口小时降序处理，并结合连续多周期情况判断临时措施或投资。"],
            ["ModeB优化周期数", "本次ModeB实际分析的周或月数量。", "与热力图横向周期范围相互核对。"],
            ["ModeB异常周期数", "求解状态为失败的周期数量。", "大于0时先查看ModeB优化周期明细中的失败说明。"],
        ],
        [2750, 4020, 2590],
    ))

    parts.append(h2("4.4 核心指标：维护与数据质量"))
    parts.append(table(
        ["指标", "表示什么", "怎么用"],
        [
            ["未维护工作中心工序数", "生产工序能读到，但工作中心表中没有对应产能维护的工序数量。", "这些工序保留在报告中，但不进入可选路径查找和正常产能率计算。"],
            ["未维护工作中心负荷", "上述未维护工作中心工序的总负荷小时。", "数值大时优先补充工作中心和日历，否则总产能判断会缺一部分。"],
            ["占位交期订单数", "供给年份为2049年及以后、被识别为占位日期并排除计算的订单数量。", "到“占位交期订单”sheet确认这些订单是否应改成真实交期。"],
            ["数据质量/调整记录数", "本次读取过程中产生的提醒、缺失、跳过和日期调整记录总数。", "它不是全部错误数；应打开数据质量页按“类型”筛选，区分提醒与需维护事项。"],
        ],
        [2800, 4000, 2560],
    ))

    parts.append(h2("4.5 核心指标：预测需求完整性"))
    parts.append(table(
        ["指标", "表示什么", "怎么用"],
        [
            ["预测顶层物料数", "本次预测表中有正需求、并进入BOM与工艺路线检查的不同顶层物料数量。", "作为后续三个缺失数量的对照基数。"],
            ["预测顶层物料未匹配BOM数", "三地BOM均未找到的预测顶层物料种数。", "这些物料仍会尝试匹配自身工艺路线；需要BOM完整分析时应补BOM。"],
            ["预测顶层物料未匹配工艺路线数", "三地工艺路线均未找到的预测顶层物料种数。", "这些顶层预测物料未生成工序，因此没有进入负荷。"],
            ["BOM下层制造件未匹配工艺路线数", "按“顶层物料+下层M制造件”去重后的缺路线组合数量。", "同一M物料出现在不同顶层BOM时可分别计数，便于定位受影响的产品链。"],
            ["可参与计算的预测虚拟订单数", "至少成功生成一道工序并进入分析的预测虚拟订单号数量。", "可在订单工序分配明细中按FCST开头的订单号筛选。"],
        ],
        [3000, 3950, 2410],
    ))
    parts.append(callout("回归样例", "当前测试数据中，预测顶层物料数为1,037，未匹配BOM为13，顶层未匹配工艺路线为1,035，BOM下层制造件缺路线组合为479，可参与计算的预测虚拟订单为26。每次运行会根据输入文件重新计算，操作员应以当次报告为准。"))

    parts.append(h2("4.6 重点瓶颈工作中心"))
    parts.append(table(
        ["字段", "含义", "阅读方法"],
        [
            ["工作中心", "需要重点关注的具体工作中心。", "可复制该名称到周度/月度产能分析中筛选。"],
            ["期间", "发生瓶颈的周或月。", "周报告显示W周号，月报告显示月份。"],
            ["日期跨度", "该周/月实际包含的起止日期。", "第一周可能从优化开始日期起算，不一定包含完整7天。"],
            ["负荷率", "负荷小时除以产能小时。", "100%表示刚好满载，150%表示需求是可用产能的1.5倍。"],
            ["负荷小时", "该工作中心在该周期需要承担的厂内小时。", "ModeA是原始需求口径；ModeB是优化后厂内口径。"],
            ["产能小时", "按设备数量和工作日历计算的周期理论产能。", "先确认日历和设备数量维护正确，再使用负荷率判断。"],
        ],
        [1800, 4050, 3510],
    ))
    parts.append(paragraph("仪表盘“运行判断”中的ModeA和ModeB说明是固定的阅读口径提示，不是本次运行自动生成的结论。"))

    parts.append(page_break())
    parts.append(h1("5. 工作组热力图怎么读"))
    parts.append(paragraph("热力图先按工作中心表E列“大类”分区，再按资源组显示。ModeA的需求工时使用无限产能负荷；ModeB的需求工时使用可选路径优化后的厂内负荷。"))
    parts.append(table(
        ["区域", "含义", "例子"],
        [
            ["资源组", "工作中心表中的资源组分类。", "立车KV1000、北京小五轴等。"],
            ["设备台数", "该资源组在工作中心表中维护的设备数量。", "显示8，表示该资源组当前维护8台设备。"],
            ["需求工时", "该资源组在各周期需要承担的总工时。", "9月需求7,008小时。"],
            ["理论工时", "按设备数量和工作日历得到的资源组可用小时。", "8台设备合计理论工时5,376小时。"],
            ["负荷", "需求工时除以理论工时。", "7,008÷5,376≈130%，表示超过理论产能约30%。"],
            ["设备需求数量", "在相同日历和单台能力下，满足需求所需的设备总台数，结果向上取整。", "单台672小时，7,008÷672≈10.43，向上取整为11台；当前8台，说明还差约3台等效能力。"],
        ],
        [2200, 4000, 3160],
    ))
    parts.append(h2("5.1 颜色规则"))
    parts.append(table(
        ["负荷范围", "颜色", "操作提示"],
        [
            ["0%至25%", "深绿色", "资源余量较多，可关注是否具备承接分流的条件。"],
            ["大于25%至75%", "浅绿色", "处于可用区间，继续结合未来周期观察。"],
            ["大于75%至100%", "浅红色", "接近满载，新增需求前应复核。"],
            ["超过100%", "深红色", "需求超过理论产能，需要分流、外包、排班或设备措施。"],
        ],
        [2100, 1800, 5460],
    ))
    parts.append(callout("阅读顺序", "先看深红色周期，再看设备需求数量与当前设备台数的差额，最后回到周度/月度产能分析定位具体工作中心，并在订单工序分配明细中找出贡献负荷的订单。"))

    parts.append(h1("6. 订单工序分配明细中的BOM字段"))
    parts.append(table(
        ["字段", "含义", "使用示例"],
        [
            ["BOM根订单", "同一条预测BOM链共享的顶层预测虚拟订单号。", "筛选一个BOM根订单，可一起查看顶层物料和全部可计算下层制造件。"],
            ["BOM顶层物料", "需求预测表中最上层的物料。", "不论当前行是顶层还是M制造件，该字段都保持顶层物料号。"],
            ["BOM来源", "本次使用的BOM地区来源。", "显示北京、沈阳、南通或未找到BOM。"],
            ["BOM来源文件", "实际匹配到的BOM文件路径。", "用于追溯应维护哪一份BOM。"],
            ["BOM层级", "当前制造物料在BOM中的层级；0表示顶层。", "层级越深的制造件越早准备。"],
            ["BOM父项物料", "当前制造件所关联的上级制造件集合。", "用于理解当前件完成后支持哪些上层制造件。"],
            ["BOM直接父项物料", "BOM表中与当前物料直接相连的父项物料。", "用于检查直接装配或投料关系。"],
            ["BOM精确需求数量", "按BOM用量计算、向上取整前的需求数量。", "精确需求12.4件时，实际制造分析数量按13件进入。"],
            ["BOM齐套需求日期", "下层制造件需要完成、以便上层物料开始的日期。", "用于判断下层制造件是否在父项开始前齐套。"],
        ],
        [2300, 4200, 2860],
    ))
    parts.append(callout("筛选技巧", "预测虚拟订单号以FCST开头。先筛选BOM根订单，再按BOM层级从大到小、活动号从小到大查看，可以更快理解下层制造件到顶层物料的产能链。"))

    parts.append(h1("7. 常见情况与处理"))
    parts.append(table(
        ["看到的情况", "表示什么", "处理建议"],
        [
            ["预测缺失数较大，但工具运行完成", "预测数据缺口采用记录并跳过的方式，没有影响其他可计算数据。", "先在数据质量页按缺失类型筛选，再安排BOM或工艺路线补充。"],
            ["可参与计算的预测虚拟订单数明显偏少", "多数预测物料没有匹配到工艺路线，或预测月份没有正需求。", "优先核对三地工艺路线的物料编码和Data sheet。"],
            ["资源组显示“未分类”", "工作中心表E列大类为空。", "补充大类后重新运行。"],
            ["资源组显示“大类冲突”", "同一资源组下的工作中心维护了不同大类。", "确认资源组归属，统一大类或拆分资源组。"],
            ["ModeB求解状态为SKIPPED", "没有瓶颈、没有可参与优化的工序，或瓶颈工序没有可选候选。", "这不是报错；结合状态说明和可选工序分流数判断是否需要维护可选路径。"],
            ["热力图设备需求数量显示“产能未维护”", "缺少设备数量、日历或可用产能。", "补全工作中心设备数量和日历映射。"],
            ["未维护工作中心负荷不为0", "部分工序的工作中心没有产能主数据。", "在工作中心表补充对应工作中心后重新运行，避免低估总压力。"],
        ],
        [2800, 3600, 2960],
    ))

    parts.append(h1("8. 每次运行前检查清单"))
    parts.append(bullets([
        "确认输入文件已保存并关闭，避免Excel锁定或读取到未保存版本。",
        "确认订单号和物料号保持文本格式，没有科学计数法或末尾精度丢失。",
        "确认优化开始日期是本次希望使用的日期。",
        "需要预测时确认已勾选“需求预测导入”，预测文件路径不为空。",
        "确认三份BOM和三份工艺路线文件名、Data sheet名称未被修改。",
        "确认工作中心E列“大类”、资源组、设备数量和日历名称已维护。",
        "先运行ModeA确认原始压力，再运行ModeB评估可选工序和外包改善。",
        "报告生成后先看仪表盘中的数据质量、预测完整性和异常周期，再阅读热力图。",
    ]))
    parts.append(callout("版本确认", "GUI标题和报告“运行信息”sheet中的工具版本应显示0.1.7。若不是，请确认启动的是本次发布包中的FortuneBJOptimizer.exe。"))
    return "".join(parts)


def styles_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="202020"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Kicker"><w:name w:val="Kicker"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:caps/><w:sz w:val="20"/><w:color w:val="2E74B5"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="90"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="50"/><w:color w:val="1F4D78"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240" w:line="280" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="25"/><w:color w:val="555555"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="32"/><w:color w:val="2E74B5"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:b/><w:sz w:val="26"/><w:color w:val="2E74B5"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:ind w:left="540" w:hanging="270"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/><w:ind w:left="540" w:hanging="270"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:after="160" w:line="300" w:lineRule="auto"/><w:shd w:fill="F4F6F9"/><w:ind w:left="160" w:right="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:color w:val="1F3A5F"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="20" w:line="260" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="20"/><w:color w:val="202020"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="B7C4D0"/><w:left w:val="single" w:sz="4" w:space="0" w:color="B7C4D0"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="B7C4D0"/><w:right w:val="single" w:sz="4" w:space="0" w:color="B7C4D0"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D7DEE5"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D7DEE5"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''


def numbering_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W_NS}">
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>'''


def document_xml(body: str) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>'''


def header_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="right"/><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="18"/><w:color w:val="6B7280"/></w:rPr><w:t>Fortune BJ 产能优化工具 | 操作员指引</w:t></w:r></w:p></w:hdr>'''


def footer_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="18"/><w:color w:val="6B7280"/></w:rPr><w:t xml:space="preserve">v0.1.7  |  第 </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="6B7280"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="6B7280"/></w:rPr><w:t> 页</w:t></w:r></w:p></w:ftr>'''


def write_docx(path: Path) -> None:
    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''',
        "word/_rels/document.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>''',
        "word/document.xml": document_xml(build_body()),
        "word/styles.xml": styles_xml(),
        "word/numbering.xml": numbering_xml(),
        "word/header1.xml": header_xml(),
        "word/footer1.xml": footer_xml(),
        "word/settings.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/><w:updateFields w:val="true"/></w:settings>''',
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def main() -> None:
    write_docx(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
