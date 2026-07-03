# -*- coding: utf-8 -*-
"""
生成《施工合同基本情况说明》Word 文档。
数据来源：2024-192.东升、东红、千秋、东平、南俸、上埇和机关等7所幼儿园
室外配套附属设施建设工程施工合同.doc
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_east_asian_font(run, font_name="仿宋_GB2312", size=None, bold=False):
    run.font.name = font_name
    run.font.bold = bold
    if size:
        run.font.size = size
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)


def add_para(doc, text, font="仿宋_GB2312", size=Pt(14), bold=False,
             align=None, first_line_indent=True, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if first_line_indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    set_east_asian_font(run, font, size, bold)
    return p


def add_title(doc, text, size=Pt(22)):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run(text)
    set_east_asian_font(run, "黑体", size, bold=True)
    return p


def add_heading2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_east_asian_font(run, "黑体", Pt(15), bold=True)
    return p


def add_kv_row(table, key, value, key_width=Cm(3.5)):
    row = table.add_row().cells
    row[0].text = ""
    row[1].text = ""
    k_run = row[0].paragraphs[0].add_run(key)
    set_east_asian_font(k_run, "黑体", Pt(12), bold=True)
    v_run = row[1].paragraphs[0].add_run(value)
    set_east_asian_font(v_run, "仿宋_GB2312", Pt(12))
    row[0].width = key_width
    return row


doc = Document()

section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.left_margin = Cm(2.8)
section.right_margin = Cm(2.6)
section.top_margin = Cm(2.6)
section.bottom_margin = Cm(2.5)

style = doc.styles["Normal"]
style.font.name = "仿宋_GB2312"
style.font.size = Pt(14)
style_rpr = style.element.get_or_add_rPr()
style_rfonts = style_rpr.find(qn("w:rFonts"))
if style_rfonts is None:
    style_rfonts = style_rpr.makeelement(qn("w:rFonts"), {})
    style_rpr.append(style_rfonts)
style_rfonts.set(qn("w:eastAsia"), "仿宋_GB2312")

add_title(doc, "关于7所幼儿园室外配套附属设施")
add_title(doc, "建设工程施工合同基本情况的说明", size=Pt(20))

add_para(
    doc,
    "为便于领导掌握相关工程施工合同的基本情况，现将合同名称、合同双方、"
    "合同价格等主要信息说明如下：",
    first_line_indent=True,
)

add_heading2(doc, "一、项目基本情况")
info_table1 = doc.add_table(rows=0, cols=2)
info_table1.style = "Table Grid"
info_table1.alignment = WD_TABLE_ALIGNMENT.CENTER
add_kv_row(info_table1, "合同编号", "QHCT(FW)2024-192号")
add_kv_row(
    info_table1,
    "项目（合同）名称",
    "东升、东红、千秋、东平、南俸、上埇和机关等7所幼儿园室外配套附属设施建设工程",
)
add_kv_row(
    info_table1,
    "工程地点",
    "琼海市嘉积镇、万泉镇、塔洋镇、大路镇、会山镇及石壁镇",
)
add_kv_row(info_table1, "资金来源", "市级财政资金")
add_kv_row(info_table1, "计划工期", "180个日历天")
add_kv_row(info_table1, "签订地点", "海南省琼海市")

doc.add_paragraph()

add_heading2(doc, "二、合同双方")
info_table2 = doc.add_table(rows=0, cols=2)
info_table2.style = "Table Grid"
info_table2.alignment = WD_TABLE_ALIGNMENT.CENTER
add_kv_row(info_table2, "发包人（业主）", "琼海市城市投资运营有限公司")
add_kv_row(info_table2, "发包人地址", "琼海市嘉积镇爱华东路19号")
add_kv_row(info_table2, "发包人信用代码", "91469002747790848Q")
add_kv_row(info_table2, "承包人（施工单位）", "陕西宏基源建设有限公司")
add_kv_row(info_table2, "承包人地址", "西安市文艺北路5号敬业大厦1302室—1307室")
add_kv_row(info_table2, "承包人信用代码", "91610000741282199A")
add_kv_row(info_table2, "承包人项目经理", "樊俊飞（注册证号：陕1612022202302782）")

doc.add_paragraph()

add_heading2(doc, "三、合同价格")
info_table3 = doc.add_table(rows=0, cols=2)
info_table3.style = "Table Grid"
info_table3.alignment = WD_TABLE_ALIGNMENT.CENTER
add_kv_row(info_table3, "签约合同价（含税）", "人民币 11,444,630.26 元")
add_kv_row(info_table3, "　　其中：不含税价", "人民币 10,499,660.79 元")
add_kv_row(info_table3, "　　其中：增值税额（税率9%）", "人民币 944,969.47 元")
add_kv_row(info_table3, "合同价格形式", "固定单价合同")
add_kv_row(info_table3, "下浮率", "4.27%")
add_kv_row(info_table3, "安全文明施工费", "人民币 429,538.91 元")
add_kv_row(info_table3, "专业工程暂估价", "人民币 200,537.43 元")

doc.add_paragraph()

add_heading2(doc, "四、备注")
add_para(
    doc,
    "1. 以上信息摘录自《建设工程施工合同》（合同编号：QHCT(FW)2024-192号）"
    "第一部分“合同协议书”，最终结算合同价以工程竣工验收结算审核结果为准。",
    first_line_indent=False,
    size=Pt(12),
)
add_para(
    doc,
    "2. 如需查阅合同全文或相关附件（招标文件、投标文件、图纸等），"
    "请与项目经办人联系。",
    first_line_indent=False,
    size=Pt(12),
)

doc.add_paragraph()
doc.add_paragraph()
tail = doc.add_paragraph()
tail.alignment = WD_ALIGN_PARAGRAPH.RIGHT
tail_run = tail.add_run("汇报人：____________")
set_east_asian_font(tail_run, "仿宋_GB2312", Pt(14))

tail2 = doc.add_paragraph()
tail2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
tail2_run = tail2.add_run("日 期：2026年7月3日")
set_east_asian_font(tail2_run, "仿宋_GB2312", Pt(14))

output_path = "合同基本情况说明.docx"
doc.save(output_path)
print(f"已生成：{output_path}")
