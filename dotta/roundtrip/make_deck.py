"""Builds the python-pptx deck the dotta round-trip test edits."""

import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

COLUMN_WIDTHS = [Inches(3.2), Inches(2.4), Inches(2.4), Inches(3.5)]
ROW_HEIGHTS = [Inches(0.7), Inches(0.6), Inches(0.9), Inches(0.6), Inches(1.2)]
TABLE_SLIDE = 1
ELLIPSE_SLIDE = 2
ELLIPSE_LAYOUT = 1


def set_font(run, name, size=None, bold=None, color=None):
    run.font.name = name
    rpr = run._r.get_or_add_rPr()
    ea = etree.SubElement(rpr, qn("a:ea"))
    ea.set("typeface", name)
    rpr.find(qn("a:latin")).addnext(ea)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def add_run(paragraph, text, **style):
    run = paragraph.add_run()
    run.text = text
    set_font(run, **style)
    return run


def make_layout_body_an_ellipse(prs):
    body = prs.slide_layouts[ELLIPSE_LAYOUT].placeholders[1]._element
    spPr = body.spPr
    for old in spPr.findall(qn("a:prstGeom")) + spPr.findall(qn("a:custGeom")):
        spPr.remove(old)
    geometry = etree.Element(qn("a:prstGeom"), prst="ellipse")
    etree.SubElement(geometry, qn("a:avLst"))
    xfrm = spPr.find(qn("a:xfrm"))
    if xfrm is None:
        spPr.insert(0, geometry)
    else:
        xfrm.addnext(geometry)


def build(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    make_layout_body_an_ellipse(prs)

    title = prs.slides.add_slide(prs.slide_layouts[0])
    add_run(title.shapes.title.text_frame.paragraphs[0], "2026年度 第3四半期 売上報告", name="游ゴシック", bold=True)
    add_run(title.placeholders[1].text_frame.paragraphs[0], "営業企画部", name="メイリオ", size=24, color="404040")

    tabled = prs.slides.add_slide(prs.slide_layouts[5])
    tabled.shapes.title.text = "地域別の売上"
    frame = tabled.shapes.add_table(5, 4, Inches(0.8), Inches(1.6), sum(COLUMN_WIDTHS), sum(ROW_HEIGHTS))
    table = frame.table
    for index, width in enumerate(COLUMN_WIDTHS):
        table.columns[index].width = width
    for index, height in enumerate(ROW_HEIGHTS):
        table.rows[index].height = height
    data = [
        ["地域と区分", "", "前年比", "備考"],
        ["関東", "1,234", "+12.3%", "新規顧客が増加"],
        ["関西", "856", "-2.1%", "大型案件の失注"],
        ["", "432", "+4.0%", "下期に回復見込み"],
        ["合計", "2,522", "+6.8%", "目標比 98.5%"],
    ]
    for r, row in enumerate(data):
        for c, text in enumerate(row):
            if text:
                add_run(table.cell(r, c).text_frame.paragraphs[0], text, name="メイリオ", size=16, bold=r == 0)
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(2, 0).merge(table.cell(3, 0))

    badge = prs.slides.add_slide(prs.slide_layouts[ELLIPSE_LAYOUT])
    add_run(badge.shapes.title.text_frame.paragraphs[0], "番号の丸", name="游ゴシック")
    add_run(badge.placeholders[1].text_frame.paragraphs[0], "1", name="BIZ UDPゴシック", size=40)

    shapes = prs.slides.add_slide(prs.slide_layouts[6])
    rounded = shapes.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(2))
    add_run(rounded.text_frame.paragraphs[0], "手を付けない図形", name="游ゴシック", size=24)
    shapes.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6), Inches(1.5), Inches(3), Inches(1))

    prose = prs.slides.add_slide(prs.slide_layouts[6])
    box = prose.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(6.0), Inches(5.5))
    box.text_frame.word_wrap = True
    add_run(box.text_frame.paragraphs[0], "関東の売上は前年同期比で12.3%増加した。", name="游ゴシック", size=20)

    prs.save(path)


if __name__ == "__main__":
    build(sys.argv[1])
