from __future__ import annotations

import html
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs"
OUT_PATH = OUT_DIR / "Fortune_BJ_产能优化工具客户使用与报告阅读说明_CN_V05.docx"


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def run(text: str, *, bold: bool = False) -> str:
    props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f"<w:r>{props}<w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r>"


def paragraph(text: str = "", *, style: str | None = None, bold: bool = False) -> str:
    ppr = f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>" if style else ""
    return f"<w:p>{ppr}{run(text, bold=bold) if text else ''}</w:p>"


def bullets(items: list[str]) -> str:
    parts: list[str] = []
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
        widths = widths[:]
        if len(widths) != len(headers):
            raise ValueError("width count must match headers")
        widths[-1] += total_width - sum(widths)

    def cell(value: object, width: int, *, header: bool = False) -> str:
        shade = "<w:shd w:fill=\"E8EEF5\"/>" if header else ""
        return (
            f"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/>"
            "<w:vAlign w:val=\"top\"/>"
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


def callout(text: str) -> str:
    return paragraph(text, style="Callout")


def build_body() -> str:
    parts: list[str] = []
    parts.append(paragraph("Fortune BJ 产能优化工具客户使用与报告阅读说明", style="Title"))
    parts.append(paragraph("适用对象：计划、生产、工艺、设备、外协和管理人员", style="Subtitle"))
    parts.append(paragraph("版本：CN V05 | 对应工具版本：v0.1.2 | 更新日期：2026-07-24", style="Subtitle"))
    parts.append(callout("这份说明只讲客户使用方法和报告阅读方法。工具输出用于产能分析、资源分流和改善建议，不作为车间逐小时执行排程单。"))

    parts.append(section("1. 这套工具用来做什么", [
        paragraph("工具把订单需求、生产订单工序、工作中心、工作日历和可选工序放在一起，生成周度或月度产能报告。客户可以用它回答以下问题。"),
        table(["客户问题", "用工具看什么", "例子"], [
            ["哪些工作组或工作中心最忙？", "看工作组热力图、周度/月度产能分析、瓶颈分析。", "北京立加在W30显示红色，并且负荷率高于100%，说明这周压力集中。"],
            ["某个订单的每道工序被放到哪里？", "看订单工序分配明细。", "筛选订单号后，可以看到活动10、20、30分别对应哪个工作中心。"],
            ["ModeB优化后有没有分流？", "看订单工序分配明细和可选工序分流分析。", "同一活动出现多行，2件保留原工作中心，3件转可选工作中心。"],
            ["仍然超产能时怎么办？", "看100%产能优化总览和超产能解决建议。", "某工作中心W30仍有缺口，建议页会给出外包、加班或设备方向。"],
        ], widths=[2300, 3200, 3860]),
    ]))

    parts.append(section("2. 使用前准备", [
        paragraph("运行前只需要维护“数据导入”文件夹。工具不会再从其他文件夹重新抽取或覆盖数据源。"),
        table(["文件", "客户维护什么", "使用例子"], [
            ["订单交期数量_产能分析输入模板.xlsx", "维护订单号、需求数量、供给日期、紧急类型。订单号请保持文本格式。", "订单1000001数量5，供给日期2026-07-20，紧急类型填写T-14。"],
            ["生产订单工序_排产输入模板.csv", "维护每个订单的物料、活动、工序短文本和标准工时。", "订单1000001有活动10、20、30，每行都有工序短文本。"],
            ["工作中心_排产输入模板.xlsx", "维护工作中心、工作组、设备数量、日历名称；热处/表处专用模式下还可维护炉子或流水线参数。", "三轴立加数量2，资源组分类为北京立加，日历名称为7天24小时。"],
            ["工作日历_排产输入模板.xlsx", "维护每种日历每天工作多少小时、每周工作几天。", "7天24小时日历：每天24小时，每周7天。"],
            ["可选工序_排产输入模板.xlsx", "维护物料和活动的可选工作中心，或维护外包路径。", "物料A活动20可选北京小五轴；外包行把可选资源组分类填为外包。"],
        ], widths=[2600, 3900, 2860]),
        callout("建议客户每次运行前先关闭正在编辑的输入文件。Excel打开状态下保存不完整，可能导致工具读取到旧数据或临时文件。"),
    ]))

    parts.append(subsection("2.1 紧急类型怎么填", [
        paragraph("紧急类型在订单交期数量表中维护。空白表示普通订单。可填写的紧急类型只有四种。"),
        table(["紧急类型", "使用场景示例"], [
            ["越库", "需要最快处理的越库订单。"],
            ["T-14", "按T-14规则需要优先关注的订单。"],
            ["RTM", "RTM相关订单。"],
            ["临时加急", "临时被业务或客户要求加急的订单。"],
        ], widths=[1800, 7560]),
        paragraph("示例：订单A填“越库”，订单B填“T-14”，订单C不填。运行时，报告会把这些订单的紧急类型带到订单工序分配明细中，方便筛选和追溯。"),
    ]))

    parts.append(section("3. 怎么运行工具", [
        table(["步骤", "操作", "例子"], [
            ["1", "打开 FortuneBJOptimizer.exe。", "双击打包文件夹中的 FortuneBJOptimizer.exe。"],
            ["2", "确认授权状态正常。", "左下角显示授权有效；如果授权失效，先联系管理员更新license。"],
            ["3", "确认数据路径。", "配置页中的5个输入文件指向“数据导入”文件夹。"],
            ["4", "选择计算模式。", "先跑ModeA看整体压力，再跑ModeB看优化建议。"],
            ["5", "选择优化粒度。", "想看周度压力选“周”；想看月度压力选“月”。"],
            ["6", "选择优化开始周期。", "从2026-07-20开始看未完成订单，就把日期选到2026-07-20。"],
            ["7", "点击运行产能分析。", "运行完成后，弹窗可直接打开最新报告。"],
        ], widths=[800, 3600, 4960]),
    ]))

    parts.append(section("4. ModeA怎么用", [
        paragraph("ModeA用于先看原始订单压力。它适合做第一轮产能盘点，帮助客户快速找到高负荷工作组、工作中心和订单。"),
        table(["报告页", "客户怎么用", "例子"], [
            ["仪表板", "先确认本次订单数量、工序数量、瓶颈数量和数据质量记录。", "仪表板显示瓶颈数量较多，就继续看瓶颈分析和热力图。"],
            ["周度/月度产能分析", "按周期和工作中心查看负荷小时、产能小时、负荷率。", "某工作中心W30负荷率为250%，说明这一周压力很高。"],
            ["瓶颈分析", "按负荷率排序看压力最大的工作中心。", "前10行通常就是需要重点关注的资源。"],
            ["工作组热力图", "横向看多个周期的工作组压力。", "某工作组连续多周颜色较深，说明不是单周尖峰。"],
            ["订单工序分配明细", "按订单号追溯每道活动落在哪个周期和工作中心。", "筛选订单1000001，看到活动10在W30，活动20跨到W31。"],
        ], widths=[2200, 4100, 3060]),
    ]))

    parts.append(section("5. ModeB怎么用", [
        paragraph("ModeB用于看100%产能优化建议。它会结合可选工序和外包路径，给出周期级的产能分流结果。"),
        table(["报告页", "客户怎么用", "例子"], [
            ["100%产能优化总览", "看优化前后负荷变化，以及哪些工作中心仍有缺口。", "原始缺口300小时，优化后缺口80小时，说明分流后缓解了220小时。"],
            ["订单工序分配明细", "看每个订单活动被分到原工作中心、可选工作中心还是外包。", "活动20拆成两行：3件保留原工作中心，2件转外包。"],
            ["可选工序分流分析", "只看发生变化的工序，便于业务确认分流动作。", "一行显示原工作中心释放40小时，建议工作中心增加36小时。"],
            ["超产能解决建议", "查看优化后仍超产能的解决方向。", "W30某工作中心剩余缺口100小时，可评估加班、外包或设备。"],
            ["OR-Tools本次求解规模", "查看本次优化规模和求解状态。", "求解状态为OPTIMAL或FEASIBLE时，说明本次有可用优化结果。"],
        ], widths=[2300, 4300, 2760]),
        callout("ModeB报告中的负荷已经按周或月拆分。例如一条工序从W30做到W31，订单工序分配明细会显示W30一行、W31一行；热力图和产能分析页也使用同一套周期负荷。"),
    ]))

    parts.append(section("6. 工作组热力图怎么读", [
        paragraph("热力图用于快速定位压力集中在哪个工作组和哪个周期。它不替代工作中心明细；看到压力后，需要回到周度/月度产能分析查看具体工作中心。"),
        table(["热力图行", "含义", "阅读例子"], [
            ["优化后负荷率", "ModeB优化后，该工作组负荷占产能的比例。", "1.35表示该工作组在这一周仍超过100%。"],
            ["优化后负荷小时", "ModeB优化后，该工作组本周期需要消化的小时。", "W30北京小五轴显示5156.2小时，表示这周该工作组负荷很高。"],
            ["优化后缺口小时", "超过100%后剩下的小时。", "缺口1796.2小时，表示仅靠本周期产能消化不完。"],
            ["周期产能小时", "该工作组本周期可用产能合计。", "产能3360小时，负荷5156.2小时，所以需要看建议页。"],
        ], widths=[2200, 3300, 3860]),
    ]))

    parts.append(section("7. 订单工序分配明细怎么读", [
        paragraph("订单工序分配明细是最适合追溯单个订单的页面。客户可以按订单号、物料、活动、周期、工作组、优化动作筛选。"),
        table(["字段", "客户怎么理解", "例子"], [
            ["订单/物料/活动/工序短文本", "定位是哪张订单、哪个产品、哪道工序。", "订单1000001，物料A，活动20，工序短文本为北京小五轴。"],
            ["周期/周期日期跨度", "该行负荷归属到哪一周或哪一月。", "2026-W30，2026-07-20至2026-07-26。"],
            ["原工序总产品数量", "该工序本来一共要处理多少件。", "5件。"],
            ["本行分配产品数量", "这一行代表多少件产品走这条路径。", "2件转可选工作中心。"],
            ["同工序分配行", "同一订单同一活动被拆到几条路径。", "1/2、2/2表示该工序有两条路径分配。"],
            ["同分配周期拆分行", "同一条路径的负荷跨了几个周期。", "1/2表示这条路径负荷跨两周，这是第一周。"],
            ["原工作中心/优化后工作中心", "看工序从哪里来，优化后建议去哪里。", "原工作中心WC1，优化后工作中心WC2。"],
            ["优化动作", "看该行是保留原工作中心、转可选工作中心还是转外包。", "转外包表示不占用厂内工作中心产能。"],
            ["本周期负荷小时", "该行在当前周期占用的小时。", "一条工序总负荷100小时，跨两周后可能W30显示60小时，W31显示40小时。"],
            ["优化后厂内负荷小时", "优化后仍占用厂内的小时；外包行为0。", "转外包行显示0。"],
            ["原工作中心释放小时", "从原工作中心释放出来的小时。", "转可选或外包后，原工作中心压力会下降。"],
            ["外包释放本厂工时", "转外包释放的厂内小时。", "外包2件后释放20小时。"],
            ["外包返回日历天", "外包路径当前按7天返回展示。", "外包行显示7。"],
            ["说明", "补充说明拆分和外包等阅读提示。", "提示负荷小时按周期拆分，产品数量是涉及数量。"],
        ], widths=[2300, 4400, 2660]),
    ]))

    parts.append(subsection("7.1 三个常见明细例子", [
        table(["场景", "明细里看到什么", "客户怎么理解"], [
            ["全部留在原工作中心", "同一订单活动只有1行，优化动作为保留原工作中心。", "这道工序没有被转移，所有数量仍占用原工作中心。"],
            ["按整数件分流", "同一订单活动有2行：3件保留原工作中心，2件转可选工作中心。", "这道工序按产品件数拆分，不会按半件产品拆。"],
            ["跨周显示", "同一行路径出现W30和W31两条周期拆分行。", "这是同一条路径的负荷分摊到两个周期，汇总起来才是这条路径的总负荷。"],
        ], widths=[2200, 4200, 2960]),
    ]))

    parts.append(section("8. 超产能解决建议怎么用", [
        paragraph("超产能解决建议只列出优化后仍超过100%的工作中心周期。它用于把问题转成可讨论的业务动作。"),
        table(["字段", "客户怎么用", "例子"], [
            ["剩余缺口小时", "看本周期还差多少小时。", "缺口80小时，表示需要额外补80小时。"],
            ["建议外包小时", "用于评估能不能把缺口交给外包。", "如果外包资源可用，可把80小时作为外包讨论基数。"],
            ["建议加班小时", "用于评估临时排班或加班。", "本周补80小时，可以拆到多台设备或多天。"],
            ["建议新增设备数", "用于设备投资或长期能力评估。", "连续多周期都建议新增设备时，适合进入设备投资讨论。"],
            ["建议排班调整", "用于和生产现场讨论班次。", "可评估延长每日工作小时或增加周末班次。"],
        ], widths=[2200, 3900, 3260]),
    ]))

    parts.append(section("9. 热处/表处页面怎么读", [
        paragraph("如果选择热处/表处专用逻辑，报告会增加热处/表处相关字段和页面。客户使用时重点看周期负荷、容量占用、炉次或流水线吞吐情况。"),
        table(["字段或页面", "客户怎么理解", "例子"], [
            ["产能计算类型", "显示普通工时、批量处理、流水线处理或外包。", "热处理炉显示批量处理。"],
            ["容量占用", "该周期该工序占用的装炉面积或容量。", "一批产品占用2平方米。"],
            ["单炉容量", "该炉子一次可承载的有效容量。", "单炉容量5平方米。"],
            ["折算炉次", "用于看本周期大约需要多少炉。", "显示3，表示这批负荷大约需要3炉。"],
            ["流水线吞吐率", "用于看流水炉单位时间处理能力。", "每小时50件。"],
            ["热处表处产能分析", "按周期汇总热处/表处负荷和产能。", "W30热处理负荷高，需要评估外包或班次。"],
        ], widths=[2300, 4300, 2760]),
        callout("热处/表处页面用于产能分析，不是实际装炉单。实际装炉还需要现场按工艺、质量和设备状态确认。"),
    ]))

    parts.append(section("10. 推荐阅读顺序", [
        table(["目的", "阅读顺序", "例子"], [
            ["第一次看报告", "仪表板 → 工作组热力图 → 周度/月度产能分析 → 瓶颈分析。", "先找压力最大的工作组，再定位到具体工作中心。"],
            ["解释某个订单", "订单工序分配明细 → 按订单号筛选 → 按活动号排序。", "看到活动20分到外包，就继续看外包释放小时。"],
            ["解释ModeB优化效果", "100%产能优化总览 → 可选工序分流分析 → 超产能解决建议。", "先看缺口有没有下降，再看分流到哪里。"],
            ["准备管理汇报", "工作组热力图 → 超产能解决建议 → 订单工序分配明细示例。", "用热力图说明压力，用建议页说明动作，用明细页举例。"],
        ], widths=[2200, 4300, 2860]),
    ]))

    parts.append(section("11. 使用前检查清单", [
        bullets([
            "订单号保持文本格式，不要让Excel把长订单号保存成科学计数法。",
            "订单交期数量表中只保留本次要分析的有效订单。",
            "供给日期维护为业务希望满足的日期；历史未完订单会从优化开始周期进入ModeB分析。",
            "生产订单工序表中，每个有效订单都有完整活动和工序短文本。",
            "工作中心表中的资源组分类要维护清楚，因为热力图按工作组展示。",
            "工作中心表中的日历名称要能在工作日历表中找到。",
            "可选工序表中的物料和活动号要与生产订单工序表一致。",
            "外包行用可选资源组分类=外包维护；外包单位工时可以不填。",
            "运行完成后优先打开最新报告，不要混用旧版本报告解释新逻辑。",
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
