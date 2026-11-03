"""Writes samples/sample_template.pptx, samples/sample_data.xlsx and app.ico. All content is made up here.

Run: .venv\\Scripts\\python.exe -X utf8 tools\\make_samples.py
"""

import datetime as dt
import os

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(ROOT, "samples")
INK = RGBColor(0x1F, 0x2A, 0x37)
MUTED = RGBColor(0x5F, 0x68, 0x72)
ACCENT = RGBColor(0x1F, 0x5F, 0xBF)
FONT = "Malgun Gothic"


def make_workbook(path):
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "요약"
    head = Font(bold=True)
    rows = [("보고 월", dt.datetime(2026, 10, 1), 'yyyy"년" m"월"', "보고월"),
            ("매출", 1284500000, '#,##0"원"', "매출"),
            ("목표 달성률", 0.927, "0.0%", "달성률"),
            ("전월 대비", 0.042, "+0.0%;-0.0%", "전월대비"),
            ("신규 고객", 1362, '#,##0"명"', "신규고객"),
            ("담당", "경영기획팀", "General", "담당")]
    summary["A1"], summary["B1"] = "지표", "값"
    summary["A1"].font = summary["B1"].font = head
    for i, (label, value, fmt, name) in enumerate(rows, start=2):
        summary.cell(i, 1, label)
        summary.cell(i, 2, value).number_format = fmt
        wb.defined_names[name] = DefinedName(name, attr_text=f"요약!$B${i}")
    summary.column_dimensions["A"].width = 16
    summary.column_dimensions["B"].width = 20

    stores = wb.create_sheet("지점별")
    lines = [("지점", "매출", "목표", "달성률"),
             ("서울", 482000000, 500000000, 0.964), ("부산", 263000000, 280000000, 0.939),
             ("대구", 171500000, 190000000, 0.903), ("광주", 138000000, 150000000, 0.920),
             ("대전", 126000000, 140000000, 0.900), ("제주", 104000000, 125000000, 0.832)]
    for line in lines:
        stores.append(line)
    for row in stores["B2:C7"]:
        for cell in row:
            cell.number_format = "#,##0"
    for row in stores["D2:D7"]:
        row[0].number_format = "0.0%"
    table = Table(displayName="지점표", ref="A1:D7")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    stores.add_table(table)
    for col in "BCD":
        stores.column_dimensions[col].width = 16

    trend = wb.create_sheet("월별 추이")
    for line in (("월", "매출(억원)", "목표(억원)"), ("5월", 10.8, 11.0), ("6월", 11.2, 11.5), ("7월", 11.9, 12.0),
                 ("8월", 11.6, 12.5), ("9월", 12.3, 13.0), ("10월", 12.8, 13.9)):
        trend.append(line)
    for row in trend["B2:C7"]:
        for cell in row:
            cell.number_format = "0.0"
    for cell in trend[1] + stores[1]:
        cell.font = Font(bold=True)
    wb.defined_names["추이"] = DefinedName("추이", attr_text="'월별 추이'!$A$1:$C$7")
    wb.save(path)


def text(slide, x, y, w, h, value, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = value
    run.font.size, run.font.bold, run.font.name = Pt(size), bold, FONT
    run.font.color.rgb = color
    return box


def rule(slide, y):
    line = slide.shapes.add_shape(1, Inches(0.6), Inches(y), Inches(12.1), Emu(19050))
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()


def card(slide, x, label, value):
    panel = slide.shapes.add_shape(1, Inches(x), Inches(2.2), Inches(3.85), Inches(2.5))
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xFB)
    panel.line.fill.background()
    panel.shadow.inherit = False
    text(slide, x + 0.25, 2.45, 3.35, 0.5, label, size=16, color=MUTED)
    text(slide, x + 0.25, 3.05, 3.35, 1.1, value, size=26, bold=True, color=ACCENT)


def set_alt(frame, value):
    frame._element.xpath(".//p:cNvPr")[0].set("descr", value)


def make_template(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    first = prs.slides.add_slide(blank)
    text(first, 0.6, 0.45, 12, 0.9, "{{보고월}} 실적 보고", size=32, bold=True)
    text(first, 0.6, 1.2, 12, 0.5, "작성: {{담당}}", size=16, color=MUTED)
    rule(first, 1.8)
    card(first, 0.6, "매출", "{{매출}}")
    card(first, 4.74, "목표 달성률", "{{달성률}}")
    card(first, 8.88, "신규 고객", "{{신규고객}}")
    text(first, 0.6, 5.1, 12, 0.6, "전월 대비 매출 증감은 {{전월대비}}입니다. 서울 지점 달성률은 {{지점별!D2}}입니다.", size=18)
    first.notes_slide.notes_text_frame.text = "발표 메모: 이번 달 매출은 {{매출}}입니다."

    second = prs.slides.add_slide(blank)
    text(second, 0.6, 0.45, 12, 0.9, "지점별 실적", size=28, bold=True)
    rule(second, 1.3)
    frame = second.shapes.add_table(3, 4, Inches(0.6), Inches(1.6), Inches(12.1), Inches(1.5))
    set_alt(frame, "{{table:지점표}}")
    for r in range(3):
        for c in range(4):
            cell = frame.table.cell(r, c)
            cell.text = " "
            font = cell.text_frame.paragraphs[0].runs[0].font
            font.size, font.name = Pt(16), FONT
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT

    third = prs.slides.add_slide(blank)
    text(third, 0.6, 0.45, 12, 0.9, "월별 매출 추이", size=28, bold=True)
    rule(third, 1.3)
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("계열", [1, 2])
    chart_frame = third.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.5),
                                         Inches(12.1), Inches(5.4), data)
    set_alt(chart_frame, "{{chart:추이}}")
    chart = chart_frame.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "{{보고월}} 기준"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.font.size, chart.font.name = Pt(14), FONT
    prs.save(path)


def make_icon(path):
    size = 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((16, 44, 240, 212), radius=20, fill=(31, 95, 191, 255))
    draw.rounded_rectangle((36, 64, 220, 192), radius=8, fill=(255, 255, 255, 255))
    for i, height in enumerate((44, 72, 58, 96)):
        x = 62 + i * 36
        draw.rectangle((x, 176 - height, x + 22, 176), fill=(31, 95, 191, 255))
    image.save(path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    os.makedirs(SAMPLES, exist_ok=True)
    make_workbook(os.path.join(SAMPLES, "sample_data.xlsx"))
    make_template(os.path.join(SAMPLES, "sample_template.pptx"))
    make_icon(os.path.join(ROOT, "app.ico"))
    print("Wrote samples and app.ico")
