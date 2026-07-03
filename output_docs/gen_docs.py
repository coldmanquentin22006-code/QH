# -*- coding: utf-8 -*-
"""
生成东升幼儿园：
1. 工作联系单（二） —— 补齐联系单一没有覆盖、但确认单里存在的内容
2. 工程量确认单（新）—— 按联系单顺序重排，并对差异内容加注

仅涉及东升幼儿园（万泉镇中心幼儿园东升分园）部分。
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_font(run, size=12, bold=False, name="仿宋_GB2312"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), name)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_font(run, size=20, bold=True)
    return p


def add_para(doc, text, size=12, bold=False, indent_first=False):
    p = doc.add_paragraph()
    if indent_first:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    return p


def set_cell_text(cell, text, size=12, bold=False, align=None, vcenter=True):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    if vcenter:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        vAlign = tcPr.makeelement(qn('w:vAlign'), {qn('w:val'): 'center'})
        tcPr.append(vAlign)
    return cell


def build_lianxidan_2():
    doc = Document()
    for section in doc.sections:
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)

    add_title(doc, "工作联系单")
    doc.add_paragraph()

    # 工程名称 + 编号 行
    p = doc.add_paragraph()
    run = p.add_run("工程名称：东升、东红、千秋、东平、南俸、上埇和机关等7所幼儿园室外配套附属设施建设工程"
                     "（万泉镇中心幼儿园东升分园）")
    set_font(run, size=12)
    p2 = doc.add_paragraph()
    run2 = p2.add_run("编号：2")
    set_font(run2, size=12, bold=True)

    doc.add_paragraph()

    units = [
        "海口市工程监理有限公司  （监理单位）",
        "琼海市万泉镇中心幼儿园  （使用单位）",
        "中昌设计集团有限公司  （设计单位）",
        "琼海市城市投资运营有限公司  （建设单位）",
        "琼海市教育局",
    ]
    p = doc.add_paragraph()
    run = p.add_run("致：  " + units[0])
    set_font(run, size=12)
    for u in units[1:]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.0)
        run = p.add_run(u)
        set_font(run, size=12)

    doc.add_paragraph()
    add_para(doc, "事由：")
    add_para(doc, "根据现场实际情况及使用单位提出的若干意见，对原设计图纸做出部分修改，"
                   "以下内容为对工作联系单（编号1）的补充说明，工作联系单（编号1）与本联系单（编号2）"
                   "合并后与本项目东升幼儿园工程量确认单内容一一对应：")

    items = [
        ("1.",
         "大门口消防车道拆除原混凝土路面后新建，面积103.79㎡（拆除量按施工资料结算）；"
         "大门内消防车道坡道两侧增加挡土墙，总长22.97m；坡道两侧新增镀锌栏杆扶手，总长度15.5m。"),
        ("2.",
         "取消原电动伸缩门1座，替换为成品双扇咖色氟碳漆热镀锌方管大门，高2.5m，总宽6m；"
         "大门外指定场地换填300厚种植土铺设台湾草438㎡，种植面积以最终现场为准。"),
        ("3.",
         "应使用单位要求拆除单层混合结构建筑（面积据拆除影像资料据实结算）；"
         "清理后原场地做硬化并铺设悬浮式地板，面积70平方米，拼接图案厂家提供，"
         "中心点标高高于周边地面0.03m并与之找坡平接；"
         "原单层混合结构建筑门口位置增设铸铁盖板14.4m。"),
        ("4.",
         "应业主要求对新建围墙做法进行变更，其中总高2.9m的围墙长105.05m，"
         "总高2.5m的围墙长29.94m，做法见图。"),
        ("5.",
         "根据现场排水需要，在隔油池（见联系单一第3条）旁边增设两个铸铁篦子雨水口，"
         "通过5.4m DN200波纹管排至雨水井。"),
    ]
    for num, text in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.0)
        run = p.add_run(num + text)
        set_font(run, size=12)

    add_para(doc, "请各参建单位予以确认。")
    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("承包单位：")
    set_font(run, size=12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("项目经理：")
    set_font(run, size=12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("日期：")
    set_font(run, size=12)

    doc.add_paragraph()

    # 意见栏 - 5行 x 2列 表格，模仿原件（监理/使用/设计/建设/教育局）
    rows_info = [
        ("监理单位意见：", ["监理单位（盖章）：", "总监理工程师（签字）：", "日期："]),
        ("使用单位意见：", ["使用单位（盖章）：", "项目负责人（签字）：", "日期："]),
        ("设计单位意见：", ["设计单位（盖章）：", "项目负责人（签字）：", "日期："]),
        ("建设单位意见：", ["建设单位（盖章）：", "项目负责人（签字）：", "日期："]),
        ("教育局意见：", ["琼海市教育局（盖章）：", "项目负责人（签字）：", "日期："]),
    ]

    table = doc.add_table(rows=len(rows_info), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.columns[0].width = Cm(7.5)
    table.columns[1].width = Cm(8.5)

    for i, (left_label, right_lines) in enumerate(rows_info):
        row = table.rows[i]
        row.height = Cm(2.6)
        left_cell = row.cells[0]
        left_cell.text = ""
        p = left_cell.paragraphs[0]
        run = p.add_run(left_label)
        set_font(run, size=12)
        left_cell.add_paragraph()
        left_cell.add_paragraph()

        right_cell = row.cells[1]
        right_cell.text = ""
        for j, line in enumerate(right_lines):
            if j == 0:
                p = right_cell.paragraphs[0]
            else:
                p = right_cell.add_paragraph()
            run = p.add_run(line)
            set_font(run, size=11)

    doc.save("/workspace/output_docs/东升幼儿园_工作联系单（二）.docx")
    print("联系单二 done")


def build_quedingdan_new():
    doc = Document()
    for section in doc.sections:
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)

    add_title(doc, "工程量确认单（东升幼儿园）")
    doc.add_paragraph()

    add_para(doc, "说明：本确认单根据《工作联系单》（编号1，已签字盖章）与《工作联系单》（编号2，补充）"
                   "整理调整，替代原《工程量确认单》编号003（东升幼儿园）。条目顺序已按工作联系单顺序重排，"
                   "与工作联系单不一致或联系单中已取消的内容，均在对应条目下加注说明。")
    doc.add_paragraph()

    # 顶部信息表格：工程名称 / 内容 / 位置 / 计量编号
    info_table = doc.add_table(rows=4, cols=2)
    info_table.style = "Table Grid"
    info_table.columns[0].width = Cm(3.2)
    info_table.columns[1].width = Cm(12.8)

    rows = [
        ("工程名称",
         "东升、东红、千秋、东平、南俸、上埇和机关等7所幼儿园室外配套附属设施建设工程-"
         "万泉镇中心幼儿园东升分园室外配套附属设施建设工程"),
        ("内容", "东升幼儿园增加部分改造工程"),
        ("位置", "东升幼儿园"),
        ("计量编号", "003（修订）"),
    ]
    for i, (label, value) in enumerate(rows):
        set_cell_text(info_table.rows[i].cells[0], label, size=12, bold=True,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(info_table.rows[i].cells[1], value, size=12)

    doc.add_paragraph()
    add_para(doc, "事由：关于东升幼儿园增加部分改造工程的确认事宜")
    add_para(doc, "根据会议纪要，结合使用单位提出的若干意见做出修改，修改内容如下：")

    items = [
        "原设计PVC运动地板735.22㎡，变更为减震悬浮地板，规格型号：30cm×30cm×1.6cm，"
        "铺设面积664㎡；硬化区域现调整硬化为200厚级配碎石+200厚C25混凝土，硬化面积725㎡，"
        "硬化优化工序：拉毛、切缝宽6mm、深5cm、填充沥青砂浆、棉毡养护。"
        "（对应工作联系单一第1条）",

        "原设计靠近新建教学楼位置的砖砌围墙，变更为镀锌管栏杆围墙（小柱：50mm×50mm×2.0mm@80；"
        "大柱：80mm×80mm×2.5mm@3000；横杆上下两道，规格40×40×2.5mm），立柱整体高度2m，长度124米。\n"
        "【备注：工作联系单一第2条原为“铝艺栏杆围墙，立柱整体高度2.5米”，"
        "经设计单位优化调整，现变更为本条内容。】",

        "增设一个容量为2T的成品玻璃钢隔油池，埋置16.79m DN200波纹管接至W8污水井。\n"
        "【备注：工作联系单一第3条中“沙池旁增加两个洗手洗脚水龙头并进行硬化处理”"
        "“戏水池旁鹅卵石小径改为大理石小径”“教学楼入口处花池内增设排水沟”"
        "“大门绿化位置优化缩小”等内容，现均已取消。】",

        "【备注：工作联系单一第4条“新建戏水池顶部加装顶棚（防止树枝、树叶、昆虫掉入水中）”，"
        "此项已取消。】",

        "旁边增设两个铸铁篦子雨水口，通过5.4m DN200波纹管排至雨水井。"
        "（对应工作联系单二第5条，为第3条隔油池排水配套新增内容）",

        "大门口消防车道，面积103.79㎡（拆除原混凝土路面后新建，拆除量按施工资料）；"
        "大门内消防车道坡道两侧增加挡土墙，总长22.97m；坡道两侧新增镀锌栏杆扶手，总长度15.5m。"
        "（对应工作联系单二第1条）",

        "取消原电动伸缩门1座，替换为成品双扇咖色氟碳漆热镀锌方管大门，高2.5m，总宽6m；"
        "大门外指定场地换填300厚种植土铺设台湾草438㎡，种植面积以最终现场为准。"
        "（对应工作联系单二第2条）",

        "应使用单位要求拆除单层混合结构建筑（面积据拆除影像资料据实结算）；"
        "清理后原场地做硬化并铺设悬浮式地板，面积70平方米，拼接图案厂家提供，"
        "中心点标高高于周边地面0.03并与之找坡平接；"
        "原单层混合结构建筑门口位置增设铸铁盖板14.4m。"
        "（对应工作联系单二第3条）",

        "应业主要求对新建围墙做法进行变更，其中总高2.9m的围墙长105.05m，"
        "总高2.5m的围墙长29.94m，做法见图。"
        "（对应工作联系单二第4条）",
    ]

    for idx, text in enumerate(items, start=1):
        p = doc.add_paragraph()
        lines = text.split("\n")
        run = p.add_run(f"{idx}. {lines[0]}")
        set_font(run, size=12)
        for extra in lines[1:]:
            run2 = p.add_run("\n" + extra)
            set_font(run2, size=11)

    add_para(doc, "以上变更工程量详见后附计算表。")
    doc.add_paragraph()

    # 底部四方意见栏
    bottom = doc.add_table(rows=2, cols=4)
    bottom.style = "Table Grid"
    labels = ["建设单位意见：", "代管单位意见：", "监理单位意见：", "施工单位意见："]
    for i, label in enumerate(labels):
        set_cell_text(bottom.rows[0].cells[i], label, size=12)
        bottom.rows[1].cells[i].text = ""
    bottom.rows[1].height = Cm(3.0)

    doc.save("/workspace/output_docs/东升幼儿园_工程量确认单（新）.docx")
    print("新确认单 done")


if __name__ == "__main__":
    build_lianxidan_2()
    build_quedingdan_new()
