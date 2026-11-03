"""Unit tests for numfmt.py and core.py. Run: python -X utf8 test_core.py"""

import datetime as dt
import hashlib
import json
import os
import shutil
import tempfile
import unittest
import urllib.error
import zipfile
from unittest import mock

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

import core
import facts
import insight
import numfmt
import suggest
from numfmt import format_value

KOREAN = numfmt.Locale("-", "yyyy-mm-dd", True)
US = numfmt.Locale("/", "m/d/yyyy", False)
numfmt.LOCALE = KOREAN  # tests must not depend on the PC running them


def shown(value, fmt="General"):
    text, exact = format_value(value, fmt)
    assert exact, f"format not understood: {fmt}"
    return text


class Formats(unittest.TestCase):
    def test_general(self):
        self.assertEqual(shown(1234567), "1234567")
        self.assertEqual(shown(1234567.0), "1234567")
        self.assertEqual(shown(0.1 + 0.2), "0.3")
        self.assertEqual(shown(1 / 3), "0.333333333")
        self.assertEqual(shown(1234567.891), "1234567.891")
        self.assertEqual(shown(123456789012), "123456789012")
        self.assertEqual(shown(-2.5), "-2.5")
        self.assertEqual(shown("글자"), "글자")
        self.assertEqual(shown(True), "TRUE")

    def test_numbers(self):
        self.assertEqual(shown(1234567, "#,##0"), "1,234,567")
        self.assertEqual(shown(1234.5, "#,##0.00"), "1,234.50")
        self.assertEqual(shown(-1234.5, "#,##0"), "-1,235")
        self.assertEqual(shown(2.675, "0.00"), "2.68")  # Excel rounds half up on the displayed digits
        self.assertEqual(shown(0.5, "0"), "1")
        self.assertEqual(shown(0, "#,##0"), "0")
        self.assertEqual(shown(1.5, "0.0##"), "1.5")
        self.assertEqual(shown(1.23456, "0.0##"), "1.235")
        self.assertEqual(shown(0.5, "#.00"), ".50")

    def test_leftover_placeholders_match_real_excel(self):
        # Digits fill placeholders outward from the point. A leftover 0 prints a zero, a leftover # prints nothing.
        self.assertEqual(shown(4.13, "0.##0"), "4.130")
        self.assertEqual(shown(4, "0.##0"), "4.0")
        self.assertEqual(shown(4.5, "0.#0#"), "4.50")
        self.assertEqual(shown(4.123, "0.#0#"), "4.123")
        self.assertEqual(shown(0.5, "#.##0"), ".50")
        self.assertEqual(shown(4.13, r"\+###,###,###,##0.##0"), "+4.130")
        self.assertEqual(shown(5, "0#"), "05")
        self.assertEqual(shown(5, "0##"), "05")
        self.assertEqual(shown(5, "#0#"), "05")
        self.assertEqual(shown(5, "0,000"), "0,005")

    def test_number_in_a_text_format_is_not_blank(self):
        self.assertEqual(shown(1, "@"), "1")
        self.assertEqual(shown(1.5, "@"), "1.5")
        self.assertEqual(shown(-2, "@"), "-2")
        self.assertEqual(shown(0, "@"), "0")
        self.assertEqual(shown(1234567, '@" 님"'), "1234567")  # Excel drops the literal for a number
        self.assertEqual(shown("글자", '@" 님"'), "글자 님")

    def test_percent(self):
        self.assertEqual(shown(0.853, "0.0%"), "85.3%")
        self.assertEqual(shown(0.853, "0%"), "85%")
        self.assertEqual(shown(1.2, "0%"), "120%")

    def test_scaling(self):
        self.assertEqual(shown(1234567, "#,##0,"), "1,235")
        self.assertEqual(shown(1234567, '#,##0,,"백만"'), "1백만")
        self.assertEqual(shown(1234567, "0.0,"), "1234.6")

    def test_literals_and_currency(self):
        self.assertEqual(shown(1500, '#,##0"원"'), "1,500원")
        self.assertEqual(shown(1500, "[$₩-412]#,##0"), "₩1,500")
        self.assertEqual(shown(1500, "$#,##0.00"), "$1,500.00")
        self.assertEqual(shown(-1500, "$#,##0"), "-$1,500")
        self.assertEqual(shown(5, r'0\개'), "5개")

    def test_sections(self):
        self.assertEqual(shown(-5, "#,##0;(#,##0)"), "(5)")
        self.assertEqual(shown(-5, "#,##0;[Red]-#,##0"), "-5")
        self.assertEqual(shown(0, '#,##0;-#,##0;"-"'), "-")
        accounting = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
        self.assertEqual(shown(1234, accounting), "1,234")
        self.assertEqual(shown(0, accounting), "-")
        self.assertEqual(shown("없음", accounting), "없음")

    def test_scientific(self):
        self.assertEqual(shown(12345, "0.00E+00"), "1.23E+04")
        self.assertEqual(shown(0.00012, "0.0E+00"), "1.2E-04")
        self.assertEqual(shown(9.99, "0.0E+00"), "1.0E+01")

    def test_dates(self):
        day = dt.datetime(2026, 11, 3, 14, 5, 9)
        self.assertEqual(shown(day, "yyyy-mm-dd"), "2026-11-03")
        self.assertEqual(shown(day, 'yyyy"년" m"월" d"일"'), "2026년 11월 3일")
        self.assertEqual(shown(day, "mm-dd-yy"), "2026-11-03")  # built-in 14 follows the PC's locale
        self.assertEqual(shown(day, "aaa"), "화")
        self.assertEqual(shown(day, "aaaa"), "화요일")
        self.assertEqual(shown(day, "ddd, mmm d"), "Tue, Nov 3")
        self.assertEqual(shown(day, "[$-412]dddd"), "화요일")
        self.assertEqual(shown(day, "yy.m.d"), "26.11.3")

    def test_times_and_minutes(self):
        moment = dt.datetime(2026, 11, 3, 14, 5, 9)
        self.assertEqual(shown(moment, "h:mm"), "14:05")
        self.assertEqual(shown(moment, "h:mm AM/PM"), "2:05 PM")
        self.assertEqual(shown(moment, "[$-412]AM/PM h:mm"), "오후 2:05")
        self.assertEqual(shown(moment, "mm:ss"), "05:09")
        self.assertEqual(shown(moment, "yyyy-mm-dd hh:mm:ss"), "2026-11-03 14:05:09")

    def test_serial_number_with_date_format(self):
        self.assertEqual(shown(46329, "yyyy-mm-dd"), "2026-11-03")

    def test_elapsed(self):
        self.assertEqual(shown(dt.timedelta(hours=30, minutes=5), "[h]:mm"), "30:05")
        self.assertEqual(shown(1.25, "[h]:mm"), "30:00")

    def test_slash_in_a_date_is_the_system_separator(self):
        self.assertEqual(shown(46329, "yyyy/mm/dd"), "2026-11-03")
        self.assertEqual(shown(46329, r"yyyy\/mm\/dd"), "2026/11/03")
        self.assertEqual(shown(46329, 'yyyy"/"mm'), "2026/11")

    def test_locale_dependent_builtins_on_korean_windows(self):
        self.assertEqual(shown(-5, "#,##0_);[Red](#,##0)"), "-5")
        self.assertEqual(shown(-1234.5, "#,##0.00_);(#,##0.00)"), "-1,234.50")
        self.assertEqual(shown(1234, '"$"#,##0_);("$"#,##0)'), "₩1,234")
        self.assertEqual(shown(-1234, '"$"#,##0_);("$"#,##0)'), "-₩1,234")
        self.assertEqual(shown(0, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'), "-")
        self.assertEqual(shown(46329.5, "m/d/yy h:mm"), "2026-11-03 12:00")

    def test_same_builtins_on_us_windows(self):
        numfmt.LOCALE = US
        self.addCleanup(setattr, numfmt, "LOCALE", KOREAN)
        self.assertEqual(shown(-5, "#,##0_);[Red](#,##0)"), "(5)")
        self.assertEqual(shown(1234, '"$"#,##0_);("$"#,##0)'), "$1,234")
        self.assertEqual(shown(46329, "mm-dd-yy"), "11/3/2026")
        self.assertEqual(shown(46329, "yyyy/mm/dd"), "2026/11/03")

    def test_unknown_format_is_reported_not_guessed(self):
        for fmt in ("# ?/?", "[DBNum1]General", "[>=100]0;0", "000-0000", "[$-F800]dddd"):
            text, exact = format_value(46329, fmt)
            self.assertFalse(exact, fmt)
            self.assertEqual(text, "46329")


# ---------------------------------------------------------------- fixtures

def make_workbook(path, fill=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    ws["A1"], ws["B1"] = "항목", "값"
    ws["A2"], ws["B2"] = "매출", 1234567
    ws["B2"].number_format = "#,##0"
    ws["A3"], ws["B3"] = "달성률", 0.853
    ws["B3"].number_format = "0.0%"
    ws["B4"] = dt.datetime(2026, 11, 3)
    ws["B4"].number_format = 'yyyy"년" m"월" d"일"'
    ws["B5"] = "한 줄\n두 줄"
    wb.defined_names["매출"] = DefinedName("매출", attr_text="요약!$B$2")
    wb.defined_names["달성률"] = DefinedName("달성률", attr_text="요약!$B$3")
    wb.defined_names["매출범위"] = DefinedName("매출범위", attr_text="'데이터 표'!$A$1:$C$4")
    wb.defined_names["동적"] = DefinedName("동적", attr_text="OFFSET(요약!$A$1,0,0,2,2)")

    data = wb.create_sheet("데이터 표")
    rows = [["지점", "1월", "2월"], ["서울", 100, 110.5], ["부산", 80, 90], ["대구", 60, 70]]
    for row in rows:
        data.append(row)
    for row in data["B2:C4"]:
        for cell in row:
            cell.number_format = "#,##0.0"
    tbl = wb.create_sheet("표시트")
    for row in rows:
        tbl.append(row)
    table = Table(displayName="매출표", ref="A1:C4")
    tbl.add_table(table)
    if fill:
        fill(wb)
    wb.save(path)
    return path


def textbox_runs(slide, parts, size=None, color=None):
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    paragraph = box.text_frame.paragraphs[0]
    for i, text in enumerate(parts):
        run = paragraph.add_run()
        run.text = text
        if i == 0 and size:
            run.font.size = Pt(size)
            run.font.bold = True
            run.font.color.rgb = color
    return box


def add_table(slide, rows, cols, alt):
    frame = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(6), Inches(0.4 * rows))
    frame._element.xpath(".//p:cNvPr")[0].set("descr", alt)
    return frame


def add_chart(slide, alt, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
    data = CategoryChartData()
    data.categories = ["옛날"]
    data.add_series("예전", [1])
    frame = slide.shapes.add_chart(chart_type, Inches(1), Inches(1), Inches(6), Inches(4), data)
    frame._element.xpath(".//p:cNvPr")[0].set("descr", alt)
    return frame


class Base(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.xlsx = make_workbook(os.path.join(self.dir, "data.xlsx"))
        self.pptx = os.path.join(self.dir, "template.pptx")
        self.out = os.path.join(self.dir, "out.pptx")
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def go(self, xlsx=None):
        self.prs.save(self.pptx)
        return core.run(self.pptx, xlsx or self.xlsx, self.out), Presentation(self.out)

    @staticmethod
    def item(report, key):
        return next(i for i in report.items if i.key == key)


class TextFill(Base):
    def test_defined_name(self):
        textbox_runs(self.slide, ["매출은 {{매출}}입니다"])
        report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "매출은 1,234,567입니다")
        item = self.item(report, "매출")
        self.assertEqual((item.value, item.source, item.status), ("1,234,567", "요약!B2", "ok"))

    def test_cell_reference(self):
        textbox_runs(self.slide, ["{{요약!B3}} / {{'데이터 표'!B2}} / {{데이터 표!C2}} / {{요약!$B$2}}"])
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "85.3% / 100.0 / 110.5 / 1,234,567")

    def test_split_runs_keep_first_run_format(self):
        red = RGBColor(0xC0, 0, 0)
        textbox_runs(self.slide, ["{{", "매", "출}}", " 원"], size=30, color=red)
        _report, out = self.go()
        paragraph = out.slides[0].shapes[0].text_frame.paragraphs[0]
        self.assertEqual(paragraph.text, "1,234,567 원")
        run = paragraph.runs[0]
        self.assertEqual(run.text, "1,234,567")
        self.assertEqual((run.font.size, run.font.bold, run.font.color.rgb), (Pt(30), True, red))

    def test_several_in_one_paragraph_across_runs(self):
        textbox_runs(self.slide, ["A{{매", "출}}B{{달성", "률}}C"])
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "A1,234,567B85.3%C")

    def test_spaces_and_smart_quotes_in_key(self):
        textbox_runs(self.slide, ["{{ 매출 }} {{‘데이터 표’!B2}}"])
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "1,234,567 100.0")

    def test_multiline_value_becomes_line_break(self):
        textbox_runs(self.slide, ["{{요약!B5}}"])
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "한 줄\v두 줄")

    def test_date_value(self):
        textbox_runs(self.slide, ["{{요약!B4}}"])
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "2026년 11월 3일")

    def test_inside_table_group_notes_and_layout(self):
        frame = self.slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
        frame.table.cell(0, 0).text = "{{매출}}"
        group = self.slide.shapes.add_group_shape()
        group.shapes.add_textbox(Inches(1), Inches(3), Inches(2), Inches(1)).text_frame.text = "{{달성률}}"
        self.slide.notes_slide.notes_text_frame.text = "메모 {{매출}}"
        layout_shape = next(s for s in self.prs.slide_layouts[6].shapes if s.has_text_frame)
        layout_shape.text_frame.text = "바닥글 {{달성률}}"
        report, out = self.go()
        shapes = list(out.slides[0].shapes)
        self.assertEqual(shapes[0].table.cell(0, 0).text, "1,234,567")
        self.assertEqual(shapes[1].shapes[0].text_frame.text, "85.3%")
        self.assertEqual(out.slides[0].notes_slide.notes_text_frame.text, "메모 1,234,567")
        layout_text = [s.text_frame.text for s in out.slide_layouts[6].shapes if s.has_text_frame]
        self.assertIn("바닥글 85.3%", layout_text)
        self.assertIn(("notes", 1), self.item(report, "매출").wheres)

    def test_chart_title(self):
        frame = add_chart(self.slide, "")
        frame.chart.has_title = True
        frame.chart.chart_title.text_frame.text = "기준일 {{요약!B4}}"
        _report, out = self.go()
        title = out.slides[0].shapes[0].chart.chart_title.text_frame.text
        self.assertEqual(title, "기준일 2026년 11월 3일")


class TableFill(Base):
    def cell_font(self, frame, row, col):
        return frame.table.cell(row, col).text_frame.paragraphs[0].runs[0].font

    def style_cell(self, frame, row, col, text):
        cell = frame.table.cell(row, col)
        cell.text = text
        font = cell.text_frame.paragraphs[0].runs[0].font
        font.size = Pt(19)
        font.bold = True

    def test_same_size(self):
        add_table(self.slide, 4, 3, "{{table:'데이터 표'!A1:C4}}")
        report, out = self.go()
        table = out.slides[0].shapes[0].table
        self.assertEqual(table.cell(0, 0).text, "지점")
        self.assertEqual(table.cell(1, 2).text, "110.5")
        self.assertEqual(table.cell(3, 1).text, "60.0")
        self.assertEqual(self.item(report, "'데이터 표'!A1:C4").status, "ok")

    def test_more_rows_than_template_copies_format(self):
        frame = add_table(self.slide, 2, 3, "{{table:매출범위}}")
        self.style_cell(frame, 1, 0, "자리")
        _report, out = self.go()
        new = out.slides[0].shapes[0]
        self.assertEqual(len(new.table.rows), 4)
        self.assertEqual(new.table.cell(3, 0).text, "대구")
        font = self.cell_font(new, 3, 0)
        self.assertEqual((font.size, font.bold), (Pt(19), True))
        self.assertEqual(new.height, sum(r.height for r in new.table.rows))

    def test_fewer_rows(self):
        add_table(self.slide, 8, 3, "{{table:매출범위}}")
        _report, out = self.go()
        new = out.slides[0].shapes[0]
        self.assertEqual(len(new.table.rows), 4)
        self.assertEqual(new.height, sum(r.height for r in new.table.rows))

    def test_excel_table_name(self):
        add_table(self.slide, 4, 3, "{{표:매출표}}")
        _report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].table.cell(2, 0).text, "부산")

    def test_total_row_keeps_its_place(self):
        frame = add_table(self.slide, 3, 3, "{{table:매출범위}}")
        frame.table.last_row = True
        self.style_cell(frame, 2, 0, "합계")
        frame.table.cell(1, 0).text = "본문"
        _report, out = self.go()
        new = out.slides[0].shapes[0]
        self.assertEqual(len(new.table.rows), 4)
        self.assertEqual(self.cell_font(new, 3, 0).size, Pt(19))  # last row keeps the total format
        self.assertIsNone(self.cell_font(new, 2, 0).size)  # the added row looks like body rows

    def test_new_rows_get_unique_ids(self):
        frame = add_table(self.slide, 2, 3, "{{table:매출범위}}")
        tr = frame.table._tbl.tr_lst[0]
        ext = ('<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="x">'
               '<a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="7"/>'
               '</a:ext></a:extLst>')
        from lxml import etree
        for i, row in enumerate(frame.table._tbl.tr_lst):
            node = etree.fromstring(ext)
            node[0][0].set("val", str(7 + i))
            row.append(node)
        self.assertIsNotNone(tr)
        _report, out = self.go()
        ids = [n.get("val") for n in out.slides[0].shapes[0]._element.iter()
               if etree.QName(n).localname == "rowId"]
        self.assertEqual(len(ids), 4)
        self.assertEqual(len(set(ids)), 4)

    def test_column_mismatch_leaves_table_untouched(self):
        frame = add_table(self.slide, 2, 2, "{{table:매출범위}}")
        frame.table.cell(0, 0).text = "그대로"
        report, out = self.go()
        self.assertEqual(self.item(report, "매출범위").status, "error")
        self.assertEqual(out.slides[0].shapes[0].table.cell(0, 0).text, "그대로")

    def test_marker_removed_from_alt_text_on_success_only(self):
        add_table(self.slide, 4, 3, "{{table:매출범위}}")
        add_table(self.slide, 4, 3, "{{table:없는이름}}")
        _report, out = self.go()
        first, second = out.slides[0].shapes[0], out.slides[0].shapes[1]
        self.assertIsNone(first._element.xpath(".//p:cNvPr")[0].get("descr"))
        self.assertEqual(second._element.xpath(".//p:cNvPr")[0].get("descr"), "{{table:없는이름}}")

    def test_cell_data_that_looks_like_placeholder_is_not_expanded(self):
        def fill(wb):
            wb["데이터 표"]["A2"] = "{{매출}}"
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        add_table(self.slide, 4, 3, "{{table:매출범위}}")
        _report, out = self.go(xlsx)
        self.assertEqual(out.slides[0].shapes[0].table.cell(1, 0).text, "{{매출}}")

    def test_overflow_warning(self):
        def fill(wb):
            sheet = wb.create_sheet("긴표")
            for i in range(60):
                sheet.append([i, i])
            wb.defined_names["긴표"] = DefinedName("긴표", attr_text="긴표!$A$1:$B$60")
        xlsx = make_workbook(os.path.join(self.dir, "d3.xlsx"), fill)
        add_table(self.slide, 2, 2, "{{table:긴표}}")
        report, _out = self.go(xlsx)
        codes = [i.code for i in self.item(report, "긴표").issues]
        self.assertIn("table_overflow", codes)
        self.assertEqual(self.item(report, "긴표").status, "warn")

    def wrapped_table(self, text, cols=6, rows=8):
        def fill(wb):
            sheet = wb.create_sheet("줄바꿈")
            for _ in range(rows):
                sheet.append([text] * cols)
        xlsx = make_workbook(os.path.join(self.dir, "d4.xlsx"), fill)
        key = f"줄바꿈!A1:{chr(64 + cols)}{rows}"
        add_table(self.slide, rows, cols, "{{table:" + key + "}}")
        report, _out = self.go(xlsx)
        return self.item(report, key)

    def test_overflow_warning_counts_wrapped_lines(self):
        item = self.wrapped_table("유량 및 수질관리 인천광역시 강화군 온수 공공하수처리시설")
        self.assertIn("table_overflow", [i.code for i in item.issues])

    def test_no_overflow_warning_for_short_text(self):
        self.assertEqual(self.wrapped_table("짧은 글").status, "ok")

    def test_overflow_warning_counts_line_breaks(self):
        item = self.wrapped_table("가\n나\n다\n라\n마\n바", cols=2)
        self.assertIn("table_overflow", [i.code for i in item.issues])


class ChartFill(Base):
    def test_replaces_data_and_embedded_workbook(self):
        add_chart(self.slide, "{{chart:'데이터 표'!A1:C4}}")
        report, out = self.go()
        chart = out.slides[0].shapes[0].chart
        plot = chart.plots[0]
        self.assertEqual(list(plot.categories), ["서울", "부산", "대구"])
        self.assertEqual([s.name for s in plot.series], ["1월", "2월"])
        self.assertEqual(list(plot.series[1].values), [110.5, 90.0, 70.0])
        self.assertEqual(self.item(report, "'데이터 표'!A1:C4").status, "ok")
        # PowerPoint's "Edit Data" opens this embedded workbook, so it must hold the new numbers too.
        import io
        blob = chart.part.chart_workbook.xlsx_part.blob
        embedded = openpyxl.load_workbook(io.BytesIO(blob)).active
        self.assertEqual(embedded["A2"].value, "서울")
        self.assertEqual(embedded["C2"].value, 110.5)

    def test_quoted_number_format(self):
        def fill(wb):
            for row in wb["데이터 표"]["B2:C4"]:
                for cell in row:
                    cell.number_format = '#,##0"원"'
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        add_chart(self.slide, "{{차트:매출범위}}")
        _report, out = self.go(xlsx)
        self.assertEqual(out.slides[0].shapes[0].chart.plots[0].series[0].values[0], 100)

    def test_scatter(self):
        def fill(wb):
            sheet = wb.create_sheet("산점")
            for row in (["x", "y"], [1, 2], [3, 4]):
                sheet.append(row)
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        from pptx.chart.data import XyChartData
        data = XyChartData()
        data.add_series("s").add_data_point(9, 9)
        frame = self.slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, 0, 0, Inches(4), Inches(3), data)
        frame._element.xpath(".//p:cNvPr")[0].set("descr", "{{chart:산점!A1:B3}}")
        _report, out = self.go(xlsx)
        self.assertEqual(list(out.slides[0].shapes[0].chart.plots[0].series[0].values), [2, 4])

    def test_text_in_values_warns_and_leaves_gap(self):
        def fill(wb):
            wb["데이터 표"]["B3"] = "미집계"
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        add_chart(self.slide, "{{chart:매출범위}}")
        report, out = self.go(xlsx)
        self.assertEqual(self.item(report, "매출범위").status, "warn")
        self.assertIsNone(out.slides[0].shapes[0].chart.plots[0].series[0].values[1])

    def chart_with_rows(self, rows):
        def fill(wb):
            sheet = wb.create_sheet("차트원본")
            for row in rows:
                sheet.append(row)
        xlsx = make_workbook(os.path.join(self.dir, "d5.xlsx"), fill)
        key = f"차트원본!A1:{chr(64 + len(rows[0]))}{len(rows)}"
        add_chart(self.slide, "{{chart:" + key + "}}")
        report, _out = self.go(xlsx)
        return self.item(report, key)

    def test_range_without_header_row_warns_about_the_lost_row(self):
        item = self.chart_with_rows([["12/25", 484, 602], ["12/26", 634, 618], ["12/27", 592, 629]])
        issue = item.issues[0]
        self.assertEqual(issue.code, "chart_header_number")
        self.assertEqual(issue.params["value"], "484")
        self.assertEqual(item.status, "warn")

    def test_year_columns_are_a_real_header(self):
        item = self.chart_with_rows([["지점", 2024, 2025], ["서울", 100, 110], ["부산", 80, 90]])
        self.assertEqual(item.status, "ok")

    def test_text_header_does_not_warn(self):
        item = self.chart_with_rows([["일자", "수송"], ["12/25", 100], ["12/26", 80]])
        self.assertEqual(item.status, "ok")

    def test_marker_on_wrong_shape(self):
        box = textbox_runs(self.slide, ["글자"])
        box._element.xpath(".//p:cNvPr")[0].set("descr", "{{chart:매출범위}}")
        report, _out = self.go()
        self.assertEqual(self.item(report, "매출범위").issues[0].code, "marker_wrong_shape")


class Latest(Base):
    """{{최신:시트!날짜열,값열}}: the row with the newest date in the first column listed."""

    DAILY = [["일자", "매출", "방문"], [dt.datetime(2026, 11, 1), 100, 10], [dt.datetime(2026, 11, 3), 300, 30],
             [dt.datetime(2026, 11, 2), 200, 20]]

    def book(self, rows=None, fmt=None):
        def fill(wb):
            sheet = wb.create_sheet("일별")
            for row in rows or self.DAILY:
                sheet.append(row)
            for row in sheet.iter_rows():
                if isinstance(row[0].value, dt.datetime):
                    row[0].number_format = "yyyy-mm-dd"
                if fmt:
                    row[1].number_format = fmt
        return make_workbook(os.path.join(self.dir, "daily.xlsx"), fill)

    def text(self, key, rows=None, fmt=None):
        textbox_runs(self.slide, ["{{" + key + "}}"])
        report, out = self.go(self.book(rows, fmt))
        return self.item(report, key), out.slides[0].shapes[0].text_frame.text

    def table(self, key, rows=None, template_rows=2, cols=3, header=None):
        frame = add_table(self.slide, template_rows, cols, "{{table:" + key + "}}")
        for c, value in enumerate(header or []):
            frame.table.cell(0, c).text = value
        report, out = self.go(self.book(rows))
        table = out.slides[0].shapes[0].table
        return self.item(report, key), [[c.text for c in row.cells] for row in table.rows]

    def chart(self, key, rows=None):
        add_chart(self.slide, "{{chart:" + key + "}}")
        report, out = self.go(self.book(rows))
        return self.item(report, key), out.slides[0].shapes[0].chart.plots[0]

    def test_value_on_the_newest_date_even_when_not_the_last_row(self):
        item, text = self.text("최신:일별!A,B", fmt="#,##0원")
        self.assertEqual(text, "300원")
        self.assertEqual(item.status, "ok")
        self.assertEqual(item.source, "2026-11-03 · 일별!B3")

    def test_the_date_itself(self):
        rows = [["일자", "매출"], ["2026-11-02", 1], ["2026-11-03", 2]]
        self.assertEqual(self.text("최신:일별!A", rows)[1], "2026-11-03")

    def test_english_name_and_spaces(self):
        self.assertEqual(self.text("LATEST : 일별!A, B")[1], "300")

    def test_quoted_sheet_with_a_space(self):
        def fill(wb):
            sheet = wb.create_sheet("일별 자료")
            for row in self.DAILY:
                sheet.append(row)
        textbox_runs(self.slide, ["{{최신:'일별 자료'!A,B}}"])
        _report, out = self.go(make_workbook(os.path.join(self.dir, "q.xlsx"), fill))
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "300")

    def test_text_dates_skip_subtotal_rows(self):
        rows = [["일자", "수송"], [None, "합계"], ["2025/12/30", 5], ["소계", 5], ["2025/12/31", 7], ["소계", 7], ["총계", 12]]
        item, text = self.text("최신:일별!A,B", rows)
        self.assertEqual(text, "7")
        self.assertIn("일별!B5", item.source)

    def test_date_spellings_people_use(self):
        for spelling in ("2026.11.03", "2026. 11. 3.", "2026년 11월 3일", "2026-11-03(화)", "2026/11/3", "20261103", 20261103):
            with self.subTest(spelling=spelling):
                self.setUp()
                rows = [["일자", "값"], ["2026-11-02", "옛값"], [spelling, "새값"]]
                self.assertEqual(self.text("최신:일별!A,B", rows)[1], "새값")

    def test_months_count_as_dates(self):
        rows = [["월", "값"], ["2026-10", "시월"], ["2026년 11월", "십일월"], ["2026-09", "구월"]]
        self.assertEqual(self.text("최신:일별!A,B", rows)[1], "십일월")

    def test_things_that_are_not_dates_are_ignored(self):
        rows = [["일자", "값"], ["2026-11-02", "맞음"], ["2026-13-45", "틀림"], [99999999, "틀림"], [12345, "틀림"], ["메모", "틀림"]]
        self.assertEqual(self.text("최신:일별!A,B", rows)[1], "맞음")

    def test_future_rows_with_no_value_yet_are_skipped(self):
        rows = [["일자", "매출"], [dt.datetime(2026, 11, 2), 200], [dt.datetime(2026, 11, 3), None], [dt.datetime(2026, 11, 4), ""]]
        self.assertEqual(self.text("최신:일별!A,B", rows)[1], "200")

    def test_the_date_alone_also_skips_rows_with_nothing_in_them_yet(self):
        rows = [["일자", "매출"], ["2026-11-02", 200], ["2026-11-03", None]]
        self.assertEqual(self.text("최신:일별!A", rows)[1], "2026-11-02")

    def test_a_row_still_being_filled_in_warns_that_the_value_is_older(self):
        rows = [["일자", "매출", "방문"], ["2026-11-02", 200, 20], ["2026-11-03", 300, None]]
        item, text = self.text("최신:일별!A,C", rows)
        self.assertEqual(text, "20")
        self.assertEqual(item.status, "warn")
        self.assertEqual(item.issues[0].code, "latest_older")
        self.assertEqual(item.issues[0].params, {"used": "2026-11-02", "newest": "2026-11-03"})

    def test_a_table_warns_about_a_row_still_being_filled_in_too(self):
        rows = [["일자", "매출", "방문"], ["2026-11-02", 200, 20], ["2026-11-03", None, 30]]
        item, cells = self.table("최신1:일별!A,B", rows, cols=2)
        self.assertEqual(cells[1], ["2026-11-02", "200"])
        self.assertEqual(item.issues[0].code, "latest_older")

    def test_future_dates_with_nothing_beside_them_do_not_warn(self):
        rows = [["일자", "매출", "방문"], ["2026-11-02", 200, 20], ["2026-11-03", None, None]]
        self.assertEqual(self.text("최신:일별!A,C", rows)[0].status, "ok")

    def test_same_date_with_different_values_is_not_guessed(self):
        rows = [["일자", "값"], ["2026-11-03", 1], ["2026-11-03", 2]]
        item, text = self.text("최신:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "latest_ambiguous")
        self.assertEqual(text, "{{최신:일별!A,B}}")

    def test_same_date_with_the_same_value_is_fine(self):
        rows = [["일자", "값"], ["2026-11-03", 1], ["2026-11-03", 1]]
        self.assertEqual(self.text("최신:일별!A,B", rows)[1], "1")

    def test_no_dates_stays_visible(self):
        item, text = self.text("최신:일별!B,C")
        self.assertEqual(item.issues[0].code, "latest_no_dates")
        self.assertEqual(text, "{{최신:일별!B,C}}")

    def test_bad_requests_stay_visible(self):
        for key, code in (("최신:일별", "latest_bad"), ("최신:일별!A1,2", "latest_bad"), ("최신0:일별!A,B", "latest_bad"),
                          ("최신:없는시트!A,B", "missing_sheet")):
            with self.subTest(key=key):
                self.setUp()
                item, text = self.text(key)
                self.assertEqual(item.issues[0].code, code)
                self.assertEqual(text, "{{" + key + "}}")

    def test_text_needs_one_value(self):
        for key in ("최신:일별!A,B,C", "최신7:일별!A,B"):
            with self.subTest(key=key):
                self.setUp()
                self.assertEqual(self.text(key)[0].issues[0].code, "range_for_text")

    def test_formula_without_saved_result_is_not_skipped_for_an_older_row(self):
        rows = [["일자", "값"], ["2026-11-02", 1], ["2026-11-03", "=1+1"]]
        item, text = self.text("최신:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "formula_no_cache")
        self.assertEqual(text, "{{최신:일별!A,B}}")

    def test_table_holds_the_newest_rows_oldest_first(self):
        item, cells = self.table("최신2:일별!A,B,C", header=["날짜", "매출액", "방문자"])
        self.assertEqual(item.status, "ok")
        self.assertEqual(cells, [["날짜", "매출액", "방문자"], ["2026-11-02", "200", "20"], ["2026-11-03", "300", "30"]])
        self.assertEqual(item.source, "2026-11-02 ~ 2026-11-03 · 일별!A,B,C")

    def test_blank_table_header_comes_from_the_sheet(self):
        item, cells = self.table("최신1:일별!A,B,C")
        self.assertEqual(cells[0], ["일자", "매출", "방문"])
        self.assertEqual(cells[1][1], "300")
        self.assertEqual(item.source, "2026-11-03 · 일별!A,B,C")

    def test_table_grows_to_fit_and_columns_need_not_touch(self):
        _item, cells = self.table("최신3:일별!A,C", template_rows=2, cols=2)
        self.assertEqual(cells, [["일자", "방문"], ["2026-11-01", "10"], ["2026-11-02", "20"], ["2026-11-03", "30"]])

    def test_column_range_shorthand(self):
        _item, cells = self.table("최신1:일별!A:C")
        self.assertEqual(cells[1], ["2026-11-03", "300", "30"])

    def test_fewer_rows_than_asked_warns(self):
        item, cells = self.table("최신7:일별!A,B,C")
        self.assertEqual(item.issues[0].code, "latest_fewer")
        self.assertEqual(item.issues[0].params, {"wanted": 7, "found": 3})
        self.assertEqual(len(cells), 4)

    def test_placeholder_in_a_kept_header_is_filled(self):
        _item, cells = self.table("최신1:일별!A,B,C", header=["날짜", "{{매출}}", "방문"])
        self.assertEqual(cells[0][1], "1,234,567")

    def test_marker_leaves_the_report(self):
        add_table(self.slide, 2, 3, "{{table:최신1:일별!A,B,C}}")
        _report, out = self.go(self.book())
        self.assertIsNone(out.slides[0].shapes[0]._element.xpath(".//p:cNvPr")[0].get("descr"))

    def test_chart_with_names_from_the_nearest_header_above(self):
        rows = [["합계", "수송인원\n(명)", "승차"], [None, None, "합계(A)"], ["2025/12/30", 5, 3], ["소계", 5, 3],
                ["2025/12/31", 7, 4], ["소계", 7, 4]]
        item, plot = self.chart("최신2:일별!A,B,C", rows)
        self.assertEqual(item.status, "ok")
        self.assertEqual(list(plot.categories), ["2025/12/30", "2025/12/31"])
        self.assertEqual([s.name for s in plot.series], ["수송인원 (명)", "합계(A)"])
        self.assertEqual(list(plot.series[0].values), [5, 7])

    def test_chart_with_no_header_uses_the_column_letter(self):
        _item, plot = self.chart("최신2:일별!A,B", [["2026-11-02", 1], ["2026-11-03", 2]])
        self.assertEqual([s.name for s in plot.series], ["B"])
        self.assertEqual(list(plot.series[0].values), [1, 2])

    def test_a_new_day_moves_the_value_with_no_template_change(self):
        textbox_runs(self.slide, ["{{최신:일별!A,B}}"])
        _report, out = self.go(self.book())
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "300")
        _report, out = self.go(self.book(self.DAILY + [[dt.datetime(2026, 11, 4), 400, 40]]))
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "400")


def daily(*values, start=dt.datetime(2026, 1, 5)):
    """A series with one point a day. 2026-01-05 is a Monday."""
    return [facts.Point(start + dt.timedelta(days=i), value) for i, value in enumerate(values)]


class Figures(unittest.TestCase):
    def test_change_uses_the_size_of_the_base(self):
        self.assertAlmostEqual(facts.ratio(110, 100), 0.1)
        self.assertAlmostEqual(facts.ratio(-50, -100), 0.5)  # a smaller loss is an improvement
        self.assertIsNone(facts.ratio(5, 0))

    def test_percent_text(self):
        self.assertEqual(facts.percent(-0.0427), "-4.3%")
        self.assertEqual(facts.percent(0.151), "+15.1%")
        self.assertEqual(facts.percent(0), "0.0%")

    def test_week_ago_is_exactly_seven_days_or_nothing(self):
        points = daily(*range(1, 9))
        self.assertEqual(facts.week_ago(points).value, 1)
        self.assertIsNone(facts.week_ago(points[1:]))
        gap = [p for p in points if p.value != 1]
        self.assertIsNone(facts.week_ago(gap))

    def test_mean_leaves_the_newest_out(self):
        self.assertEqual(facts.mean_before(daily(10, 20, 30, 1000), 3), 20)
        self.assertIsNone(facts.mean_before(daily(10, 20), 3))

    def test_extremes_look_only_at_recent_points_and_prefer_the_latest_tie(self):
        points = daily(999, 5, 7, 7, 1)
        self.assertEqual(facts.extreme(points, 4).date, points[3].date)
        self.assertEqual(facts.extreme(points, 4, highest=False).value, 1)

    def test_streak(self):
        self.assertEqual(facts.streak(daily(1, 2, 3, 4)), 3)
        self.assertEqual(facts.streak(daily(4, 3, 5, 4, 2)), -2)
        self.assertEqual(facts.streak(daily(1, 1)), 0)
        self.assertEqual(facts.streak(daily(1)), 0)

    def test_same_weekday_needs_three_earlier_weeks(self):
        points = daily(*range(1, 23))  # the newest is a Monday, with Mondays on 1, 8 and 15 before it
        self.assertEqual(facts.same_weekday(points)[0], 8)
        self.assertEqual(facts.same_weekday(points)[2], 3)
        self.assertIsNone(facts.same_weekday(points[7:]))

    def test_rank(self):
        self.assertEqual(facts.rank(daily(5, 9, 7)), (2, 3))
        self.assertEqual(facts.rank(daily(9, 9)), (1, 2))

    def test_week_total_change(self):
        self.assertAlmostEqual(facts.week_total_change(daily(*[10] * 7, *[11] * 7)), 0.1)
        self.assertIsNone(facts.week_total_change(daily(*[10] * 13)))

    def test_sentence_in_both_languages(self):
        points = daily(10, 10, 10, 10, 10, 10, 10, 20)  # the newest is Monday 2026-01-12
        show = lambda v: f"{v:,.0f}"  # noqa: E731
        self.assertEqual(facts.sentence(points, "수송", show),
                         "1월 12일(월) 수송 20. 전주 같은 요일보다 100.0% 증가, 전일보다 100.0% 증가, 직전 7개 평균보다 100.0% 증가.")
        self.assertEqual(facts.sentence(points, "Sales", show, "en"),
                         "Jan 12 (Mon) Sales 20. up 100.0% from the same day last week, up 100.0% from the previous day, "
                         "up 100.0% from the average of the previous 7.")

    def test_fact_pack_cites_only_calculated_figures(self):
        points = daily(*[100, 100, 100, 100, 100, 50, 50] * 3, 130)  # the newest is a Monday
        pack = {fact.id: fact for fact in facts.fact_pack(points, "수송", lambda v: f"{v:,.0f}")}
        self.assertEqual(pack["week"].text, "전주 같은 요일 1월 19일(월) 100 대비 +30.0%")
        self.assertEqual(pack["weekday"].text, "직전 3번 월요일의 중앙값 100 대비 +30.0%, 평소 변동폭의 +30.0배")
        self.assertEqual(pack["rank"].text, "최근 22개 값 중 1번째로 높음")
        self.assertIn("주말이 평일 대비 -50.9%", pack["weekend"].text)
        self.assertEqual(pack["streak"].text if "streak" in pack else "", "")
        self.assertIn("100", pack["week"].numbers)

    def test_fact_pack_leaves_out_what_it_cannot_work_out(self):
        pack = {fact.id for fact in facts.fact_pack(daily(5, 6), "값", str)}
        self.assertEqual(pack, {"value", "prev", "rank", "high", "low", "recent"})

    def test_monthly_data_gets_no_weekly_figures(self):
        points = [facts.Point(dt.datetime(2025 + (m // 12), m % 12 + 1, 1), 100 + m) for m in range(14)]
        pack = {fact.id: fact for fact in facts.fact_pack(points, "매출", str)}
        self.assertFalse({"week", "weekday", "weekend", "week_total", "signal_conflict", "signal_same_kind"} & set(pack))
        self.assertEqual(pack["value"].text, "2026년 2월 매출: 113")
        self.assertEqual(pack["year"].text, "전년 같은 달 2025년 2월 101 대비 +11.9%")

    def test_monthly_sentence_names_months(self):
        points = [facts.Point(dt.datetime(2026, m, 1), m * 10) for m in range(1, 5)]
        self.assertEqual(facts.sentence(points, "매출", str), "2026년 4월 매출 40. 직전(2026년 3월)보다 33.3% 증가.")
        self.assertEqual(facts.sentence(points, "Sales", str, "en"), "Apr 2026 Sales 40. up 33.3% from Mar 2026.")

    def test_signal_when_comparisons_point_different_ways(self):
        # The weekday level is 100 and the weekend level is 20, so a mean with weekends in it reads low.
        week = [100, 100, 100, 100, 100, 20, 20]
        pack = {fact.id: fact for fact in facts.fact_pack(daily(*week * 3, 90), "값", str)}
        self.assertEqual(pack["signal_conflict"].text,
                         "비교 기준의 방향이 엇갈림: 전주 같은 요일 기준 -10.0%, 직전 7개 평균 기준 +16.7%. 직전 7개에는 주말 2개가 섞여 있음")
        self.assertIn("평일 5개만의 평균 100 대비 -10.0%", pack["signal_same_kind"].text)

    def test_signal_for_a_day_out_of_the_usual_range(self):
        week = [100, 102, 98, 100, 102, 20, 20]
        pack = {fact.id for fact in facts.fact_pack(daily(*week * 3, 100.5, 102, 98, 100, 60, 20, 20), "값", str)}
        self.assertIn("signal_odd_day", pack)
        odd = facts.odd_day(daily(*week * 3, 100.5, 102, 98, 100, 60, 20, 20))
        self.assertEqual(odd[0].value, 60)
        self.assertAlmostEqual(odd[2], (60 - 102) / 102)

    def test_a_noisy_series_needs_a_bigger_swing_to_count_as_odd(self):
        mondays = [100, 130, 70, 110, 90]  # these swing widely, so 80 is an ordinary Monday even though it is 20% below the median
        values = [v for monday in mondays for v in (monday, 100, 100, 100, 100, 50, 50)]
        self.assertIsNone(facts.stands_out(daily(*values, 80)))

    def test_no_odd_day_when_every_day_is_usual(self):
        self.assertIsNone(facts.odd_day(daily(*[100, 100, 100, 100, 100, 20, 20] * 4)))

    def test_fact_pack_in_english(self):
        pack = {fact.id: fact for fact in facts.fact_pack(daily(*range(1, 16)), "Sales", str, "en")}
        self.assertEqual(pack["streak"].text, "14 rises in a row")
        self.assertIn("against the same weekday last week", pack["week"].text)

    def test_sentence_leaves_out_what_cannot_be_compared(self):
        self.assertEqual(facts.sentence(daily(5), "수송", str), "1월 5일(월) 수송 5.")

    def test_sentence_names_the_date_when_the_previous_point_is_not_yesterday(self):
        points = [facts.Point(dt.datetime(2026, 1, 5), 10), facts.Point(dt.datetime(2026, 1, 9), 5)]
        self.assertIn("직전(1월 5일(월))보다 50.0% 감소", facts.sentence(points, "값", str))


class Calculations(Base):
    """{{전일대비:시트!날짜열,값열}} and friends: figures worked out by SlideFill itself."""

    def key(self, key, rows):
        def fill(wb):
            sheet = wb.create_sheet("일별")
            for row in rows:
                sheet.append(row)
        textbox_runs(self.slide, ["{{" + key + "}}"])
        report, out = self.go(make_workbook(os.path.join(self.dir, "calc.xlsx"), fill))
        return self.item(report, key), out.slides[0].shapes[0].text_frame.text

    ROWS = [["일자", "수송"]] + [[f"2026-01-{day:02d}", 100 + day] for day in range(5, 12)] + [["2026-01-12", 200], ["소계", 9999]]

    def test_previous_day(self):
        item, text = self.key("전일대비:일별!A,B", self.ROWS)
        self.assertEqual(text, "+80.2%")  # 111 -> 200
        self.assertEqual(item.status, "ok")
        self.assertEqual(item.source, "2026-01-12 ↔ 2026-01-11 · 일별!B9")

    def test_same_day_last_week(self):
        item, text = self.key("전주대비:일별!A,B", self.ROWS)
        self.assertEqual(text, "+90.5%")  # 105 -> 200
        self.assertIn("2026-01-05", item.source)

    def test_average_of_the_rows_before(self):
        item, text = self.key("평균대비3:일별!A,B", self.ROWS)
        self.assertEqual(text, "+81.8%")  # mean of 109, 110, 111 -> 200
        self.assertIn("2026-01-09 ~ 2026-01-11", item.source)

    def test_highest_and_lowest_day(self):
        self.assertEqual(self.key("최고일:일별!A,B", self.ROWS)[1], "2026-01-12")
        self.setUp()
        self.assertEqual(self.key("최저일:일별!A,B", self.ROWS)[1], "2026-01-05")

    def test_sentence(self):
        _item, text = self.key("해설:일별!A,B", self.ROWS)
        self.assertEqual(text, "1월 12일(월) 수송 200. 전주 같은 요일보다 90.5% 증가, 전일보다 80.2% 증가, 직전 7개 평균보다 85.2% 증가.")

    def test_how_to_read_it(self):
        rows = [["일자", "수송"]] + [[f"2026-01-{day:02d}", value] for day, value in zip(range(5, 27), [100, 101, 99, 100, 102, 20, 21] * 3 + [90])]
        item, text = self.key("읽는법:일별!A,B", rows)
        self.assertIn("같은 요일 기준 -10.0%로 읽으세요.", text)
        self.assertEqual(item.status, "ok")
        self.setUp()
        self.assertIn("Read it by the same weekday", self.key("reading:일별!A,B", rows)[1])

    def test_too_little_data_to_read_stays_visible(self):
        rows = [["일자", "수송"], ["2026-01-11", 1], ["2026-01-12", 2]]
        item, text = self.key("읽는법:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "calc_no_base")
        self.assertEqual(text, "{{읽는법:일별!A,B}}")

    def test_english_names(self):
        self.assertEqual(self.key("week:일별!A,B", self.ROWS)[1], "+90.5%")
        self.setUp()
        self.assertTrue(self.key("note:일별!A,B", self.ROWS)[1].startswith("Jan 12 (Mon) 수송 200."))

    def test_no_day_to_compare_with_stays_visible(self):
        rows = [["일자", "수송"], ["2026-01-11", 1], ["2026-01-12", 2]]
        item, text = self.key("전주대비:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "calc_no_base")
        self.assertEqual(text, "{{전주대비:일별!A,B}}")

    def test_zero_base_stays_visible(self):
        rows = [["일자", "수송"], ["2026-01-11", 0], ["2026-01-12", 2]]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "calc_no_base")
        self.assertEqual(text, "{{전일대비:일별!A,B}}")

    def test_needs_exactly_a_date_and_a_value_column(self):
        for key in ("전일대비:일별!A", "전일대비:일별!A,B,C"):
            with self.subTest(key=key):
                self.setUp()
                self.assertEqual(self.key(key, self.ROWS)[0].issues[0].code, "calc_columns")

    def test_text_in_the_value_column_is_not_a_number(self):
        rows = [["일자", "수송"], ["2026-01-11", "많음"], ["2026-01-12", "적음"]]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "latest_no_dates")
        self.assertEqual(text, "{{전일대비:일별!A,B}}")

    def test_same_date_twice_is_not_guessed(self):
        rows = [["일자", "수송"], ["2026-01-11", 1], ["2026-01-12", 2], ["2026-01-12", 3]]
        self.assertEqual(self.key("전일대비:일별!A,B", rows)[0].issues[0].code, "latest_ambiguous")

    def source(self, rows, name="source.xlsx"):
        def fill(wb):
            sheet = wb.create_sheet("일별")
            for row in rows:
                sheet.append(row)
        return core.ExcelSource(make_workbook(os.path.join(self.dir, name), fill))

    # A calendar printed ahead: each coming month already has its queue names, but no figures.
    CALENDAR = [["월", "큐", "건수"], [dt.datetime(2026, 8, 1), "민원", 10], [dt.datetime(2026, 9, 1), "민원", 12],
                [dt.datetime(2026, 10, 1), "민원", None], [dt.datetime(2026, 11, 1), "민원", None]]

    def test_a_row_with_only_a_label_is_not_the_latest(self):
        item, text = self.key("최신:일별!A", self.CALENDAR)
        self.assertIn("2026-09-01", item.source)
        self.assertEqual(item.issues, [])

    def test_a_latest_table_skips_rows_with_only_a_label(self):
        rows, source, issues = self.source(self.CALENDAR).range_cells("최신2:일별!A,B,C")
        self.assertEqual([line[2][1] for line in rows[1:]], ["10", "12"])
        self.assertIn("2026-08-01 ~ 2026-09-01", source)
        self.assertEqual(issues, [])

    def test_a_column_of_notes_still_has_a_latest_row(self):
        rows = [["일자", "메모"], ["2026-01-11", "맑음"], ["2026-01-12", "흐림"]]
        self.assertEqual(self.key("최신:일별!A,B", rows)[1], "흐림")

    def test_the_same_date_anywhere_in_the_series_is_not_compared(self):
        # Two tables on one sheet: without this, a revision of 0.1 is compared with an index of 108.7 as if they were neighbours.
        rows = [["월", "지수"]] + [[f"2026-0{m}-01", 100 + m] for m in range(1, 7)] + [["개정", None]] + [[f"2026-0{m}-01", 0.1] for m in range(1, 4)]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "latest_duplicates")
        self.assertEqual(item.issues[0].params, {"date": "2026-03-01", "count": 3})
        self.assertEqual(text, "{{전일대비:일별!A,B}}")
        for key in ("읽는법:일별!A,B", "해설:일별!A,B", "통찰:일별!A,B"):
            with self.subTest(key=key):
                self.setUp()
                self.assertEqual(self.key(key, rows)[0].issues[0].code, "latest_duplicates")

    def test_dates_only_on_the_first_row_of_each_group_warn(self):
        rows = [["월", "큐", "건수"], [dt.datetime(2026, 8, 1), "민원", 10], [None, "상담", 5], [dt.datetime(2026, 9, 1), "민원", 12], [None, "상담", 6]]
        item, text = self.key("전일대비:일별!A,C", rows)
        self.assertEqual(text, "+20.0%")
        self.assertEqual([i.code for i in item.issues], ["latest_grouped"])
        self.assertEqual(item.issues[0].params["count"], 2)

    def test_one_total_under_the_table_is_not_a_group(self):
        rows = self.ROWS[:-1] + [[None, 9999]]
        item, _text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(item.issues, [])

    YEARS = [["구분", "외래"]] + [[year, 800000 + year] for year in range(2018, 2022)] + [["2022.10.", 702544]]

    def test_a_column_of_years_is_a_date_column(self):
        rows = [["구분", "외래"]] + [[2019, 100], ["2020년", 110], ["2021", 121.5]]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(text, "+10.5%")
        self.assertIn("2021-01-01 ↔ 2020-01-01", item.source)

    def test_a_year_among_real_dates_is_not_a_date(self):
        rows = [["일자", "수송"], ["2026-01-10", 1], [2021, 50], ["2026-01-11", 2], ["2026-01-12", 4]]
        self.assertEqual(self.key("전일대비:일별!A,B", rows)[1], "+100.0%")

    def test_part_of_a_year_warns_and_reads_per_month(self):
        item, text = self.key("읽는법:일별!A,B", self.YEARS)
        self.assertIn("10개월치로 보입니다", text)
        self.assertIn("월평균으로 맞추면", text)
        self.assertEqual(item.issues[0].code, "latest_partial_year")
        self.assertEqual(item.issues[0].params, {"date": "2022-10", "months": 10})
        self.setUp()
        item, text = self.key("직전대비:일별!A,B", self.YEARS)
        self.assertEqual(item.status, "warn")  # the figure is correct, but the reader must be told what it compares

    def test_a_month_not_over_when_the_file_was_saved_warns(self):
        now = dt.date.today()
        months = [dt.datetime(now.year - 1, m, 1) for m in range(1, 13)] + [dt.datetime(now.year, now.month, 1)]
        months = [day for day in months if day.date() <= now]
        rows = [["월", "건수"]] + [[day, 100] for day in months]
        last_day = (dt.datetime(now.year + now.month // 12, now.month % 12 + 1, 1) - dt.timedelta(days=1)).date()
        item, _text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual([i.code for i in item.issues], ["latest_unfinished"] if last_day > now else [])

    def test_a_cell_with_no_format_reads_with_sensible_digits_in_a_sentence(self):
        rows = [["일자", "수송"], ["2026-01-10", 4930000.5], ["2026-01-11", 4930900.25], ["2026-01-12", 4930915.845]]
        self.assertIn("4,930,916", self.key("해설:일별!A,B", rows)[1])
        self.setUp()
        self.assertEqual(self.key("최신:일별!A,B", rows)[1], "4930915.845")  # the value itself still shows as Excel shows it

    def test_a_row_still_being_filled_in_warns(self):
        rows = [["일자", "수송", "승차"], ["2026-01-10", 1, 1], ["2026-01-11", 2, 1], ["2026-01-12", None, 1]]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(text, "+100.0%")
        self.assertEqual(item.issues[0].code, "latest_older")

    def test_formula_without_saved_result_is_reported(self):
        rows = [["일자", "수송"], ["2026-01-11", 1], ["2026-01-12", "=1+1"]]
        item, text = self.key("전일대비:일별!A,B", rows)
        self.assertEqual(item.issues[0].code, "formula_no_cache")
        self.assertEqual(text, "{{전일대비:일별!A,B}}")


class DataShapes(Base):
    """Workbooks as people really keep them: several rows per date, periods across a row, several sheets, no dates at all."""

    def sheets(self, **tabs):
        def fill(wb):
            for name, rows in tabs.items():
                sheet = wb.create_sheet(name)
                for row in rows:
                    sheet.append(row)
        return make_workbook(os.path.join(self.dir, "shapes.xlsx"), fill)

    def value(self, key, **tabs):
        source = core.ExcelSource(self.sheets(**tabs))
        return source.text_value(key)

    def code(self, key, **tabs):
        with self.assertRaises(core.IssueError) as caught:
            self.value(key, **tabs)
        return caught.exception.issue.code

    LONG = [["월", "방향", "승차", "하차"], [dt.datetime(2026, 1, 1), "상행", 10, 1], [dt.datetime(2026, 1, 1), "하행", 20, 2],
            [dt.datetime(2026, 2, 1), "상행", 30, 3], [dt.datetime(2026, 2, 1), "하행", 50, 5], ["합계", None, 110, 11]]

    def test_dates_written_in_other_ways(self):
        for raw, wanted in (("2026 Aug", (2026, 8, 1)), ("Aug 2026", (2026, 8, 1)), ("Aug. 2026", (2026, 8, 1)), ("August 2026", (2026, 8, 1)), ("Jul. 2026(p)", (2026, 7, 1)),
                            ("Aug-26", (2026, 8, 1)), ("26년 8월", (2026, 8, 1)), ("2026 Q3", (2026, 7, 1)), ("Q3 2026", (2026, 7, 1)), ("3Q26", (2026, 7, 1)), ("2026년 3분기", (2026, 7, 1))):
            with self.subTest(raw=raw):
                self.assertEqual(core.as_date(raw), dt.datetime(*wanted))
        for raw in ("Aug 12", "메모", "Total", "Revision to index numbers", "2026 Foo", "26년 13월", "Q5 2026"):
            with self.subTest(raw=raw):
                self.assertIsNone(core.as_date(raw))

    def test_numbers_in_text_cells(self):
        for raw, wanted in (("1,234", 1234), ("362(37%)", 362), ("38,300 (78.5%)", 38300), ("15건", 15), ("12.5", 12.5), (7, 7)):
            with self.subTest(raw=raw):
                self.assertEqual(core.as_number(raw), wanted)
        for raw in ("-", "78.5%", "해당사항없음", "2025년", True, None, "1 2", "(S)"):
            with self.subTest(raw=raw):
                self.assertIsNone(core.as_number(raw))

    def test_a_mixed_cell_gives_its_number_and_says_so(self):
        rows = [["일자", "건수"], ["2026-01-11", "100(10%)"], ["2026-01-12", "150(15%)"]]
        text, _source, issues = self.value("전일대비:일별!A,B", 일별=rows)
        self.assertEqual(text, "+50.0%")
        self.assertEqual(issues[0].code, "number_from_text")
        self.assertEqual(issues[0].params["count"], 2)

    def test_rows_of_a_date_are_only_added_when_asked(self):
        self.assertEqual(self.code("전일대비:월별!A,C", 월별=self.LONG), "latest_ambiguous")
        text, source, _issues = self.value("전일대비:월별!A,합계(C)", 월별=self.LONG)
        self.assertEqual(text, "+166.7%")  # 30 -> 80; the 합계 row is not a date and drops out
        self.assertIn("2026-02-01", source)

    def test_adding_columns_and_a_condition(self):
        self.assertEqual(self.value("전일대비:월별!A,합계(C:D)", 월별=self.LONG)[0], "+166.7%")  # 33 -> 88
        self.assertEqual(self.value("전일대비:월별!A,합계(C, B=하행)", 월별=self.LONG)[0], "+150.0%")  # 20 -> 50
        self.assertEqual(self.value("prev:월별!A,sum(C, B=상행)", 월별=self.LONG)[0], "+200.0%")  # 10 -> 30
        self.assertEqual(self.code("전일대비:월별!A,합계(C, B=없는값)", 월별=self.LONG), "latest_no_dates")

    def test_counting_rows_and_a_product(self):
        rows = [["일자", "번호", "수량", "단가"], ["2026-01-11", "a", 2, 10], ["2026-01-11", "b", 1, 5], ["2026-01-12", "c", 3, 10]]
        self.assertEqual(self.value("전일대비:일별!A,건수(B)", 일별=rows)[0], "-50.0%")
        self.assertEqual(self.value("최신:일별!A,합계(C*D)", 일별=rows)[0], "30")
        self.assertEqual(self.value("전일대비:일별!A,합계(C*D)", 일별=rows)[0], "+20.0%")  # 25 -> 30

    def test_a_date_on_the_first_row_of_a_group_covers_the_rows_below_when_summing(self):
        rows = [["월", "큐", "건수"], [dt.datetime(2026, 8, 1), "민원", 10], [None, "상담", 5], [dt.datetime(2026, 9, 1), "민원", 12], [None, "상담", 6], ["Total", None, 33], [None, "민원", 999]]
        self.assertEqual(self.value("최신:일별!A,합계(C)", 일별=rows)[0], "18")
        self.assertEqual(self.value("최신:일별!A,합계(C, B=상담)", 일별=rows)[0], "6")

    def test_a_cut_short_group_warns(self):
        rows = [["월", "큐", "건수"]]
        for month in range(1, 8):
            rows += [[dt.datetime(2026, month, 1), "민원", 10], [dt.datetime(2026, month, 1), "상담", 10]]
        rows.append([dt.datetime(2026, 8, 1), "민원", 10])
        issues = self.value("최신:일별!A,합계(C)", 일별=rows)[2]
        self.assertEqual([i.code for i in issues], ["latest_group_size"])
        self.assertEqual((issues[0].params["found"], issues[0].params["usual"]), (1, 2))

    def test_a_day_that_stops_early_warns(self):
        rows = [["일시", "건수"]]
        for day in range(1, 9):
            rows += [[dt.datetime(2026, 1, day, 9), 1], [dt.datetime(2026, 1, day, 18), 1]]
        rows.append([dt.datetime(2026, 1, 9, 12, 50), 1])
        issues = self.value("최신:일별!A,합계(B)", 일별=rows)[2]
        self.assertIn("latest_partial_day", [i.code for i in issues])
        self.assertEqual(next(i for i in issues if i.code == "latest_partial_day").params["last"], "12:50")

    def test_a_table_of_totals_by_date(self):
        rows, source, _issues = core.ExcelSource(self.sheets(월별=self.LONG)).range_cells("최신2:월별!A,합계(C)")
        self.assertEqual([line[1][1] for line in rows[1:]], ["30", "80"])
        self.assertIn("2026-01-01 ~ 2026-02-01", source)

    WIDE = [["항목", "Jan. 2026", "Feb. 2026", "Mar. 2026", "합계"], ["매출", 100, 110, 121, 331], ["비용", 50, 60, 70, 180], ["메모", None, None, None, None]]

    def test_periods_across_a_row(self):
        self.assertEqual(self.value("전일대비:표!1,2", 표=self.WIDE)[0], "+10.0%")
        self.assertEqual(self.value("최신:표!1,3", 표=self.WIDE)[0], "70")
        text, source, _issues = self.value("해설:표!1,2", 표=self.WIDE)
        self.assertEqual(text, "2026년 3월 매출 121. 직전(2026년 2월)보다 10.0% 증가.")
        self.assertIn("표!D2", source)
        self.assertEqual(self.value("전일대비:표!1,합계(2:3)", 표=self.WIDE)[0], "+12.4%")  # 170 -> 191

    def test_a_table_across_a_row(self):
        rows, _source, _issues = core.ExcelSource(self.sheets(표=self.WIDE)).range_cells("최신2:표!1,2")
        self.assertEqual([[cell[1] for cell in line] for line in rows], [["항목", "매출"], ["Feb. 2026", "110"], ["Mar. 2026", "121"]])

    def test_a_month_takes_its_year_from_the_heading_above(self):
        rows = [["항", "2024년", None, None], [None, "01월", "02월", "합계"], ["가", 10, 20, 30], ["나", 5, 5, 10]]
        self.assertEqual(self.value("해설:표!2,3", 표=rows)[0], "2024년 2월 가 20. 직전(2024년 1월)보다 100.0% 증가.")
        self.assertEqual(self.value("최신:표!2,합계(3:4)", 표=rows)[0], "25")
        self.assertEqual(self.code("전일대비:표!A,B", 표=[["월", "값"], ["01월", 1], ["02월", 2]]), "latest_no_dates")  # no year anywhere is not guessed

    def test_a_range_keeps_two_tables_apart(self):
        rows = [["월", "지수"]] + [[f"2026-0{m}-01", 100 + m] for m in range(1, 5)] + [[None, None], ["개정", None]] + [[f"2026-0{m}-01", 0.1] for m in range(1, 5)]
        self.assertEqual(self.code("전일대비:표!A,B", 표=rows), "latest_ambiguous")
        self.assertEqual(self.value("전일대비:표!A2:A5,B2:B5", 표=rows)[0], "+1.0%")
        self.assertEqual(self.value("전일대비:표!A2:A5,B", 표=rows)[0], "+1.0%")
        self.assertEqual(self.code("전일대비:표!A2:A5,B2:B4", 표=rows), "latest_bad")

    def test_a_series_across_sheets(self):
        old = [["월", "건수"]] + [[dt.datetime(2025, m, 1), 80] for m in range(8, 13)]
        new = [["월", "건수"]] + [[dt.datetime(2026, m, 1), 90] for m in range(1, 8)] + [[dt.datetime(2026, 8, 1), 100]]
        text = self.value("읽는법:실적*!A,B", 실적2025=old, 실적2026=new)[0]
        self.assertIn("전년 같은 달 2025년 8월 80 대비 +25.0%", text)

    def test_a_line_called_something_else_on_another_sheet_is_refused(self):
        old = [["월", "반품"], [dt.datetime(2025, 8, 1), 80]]
        new = [["월", "건수"], [dt.datetime(2026, 7, 1), 90], [dt.datetime(2026, 8, 1), 100]]
        self.assertEqual(self.code("전일대비:실적*!A,B", 실적2025=old, 실적2026=new), "latest_sheets_differ")

    def test_sheets_that_overlap_are_not_counted_twice(self):
        one = [["일", "값"], [dt.datetime(2026, 1, 1), 1], [dt.datetime(2026, 1, 2), 2]]
        self.assertEqual(self.code("최신:년*!A,합계(B)", 년1=one, 년2=one), "latest_sheets_overlap")
        self.assertEqual(self.code("최신:없는*!A,B", 년1=one), "missing_sheet")

    SHARES = [["구분", "수입", "단위"], ["KTX", 700, "원"], ["새마을", 200, "원"], ["무궁화", 60, "원"], ["통근", 30, "원"], ["기타", 10, "원"], ["합 계", 1000, "원"]]

    def test_a_table_without_dates_reads_shares(self):
        text, source, issues = self.value("읽는법:표!A,B", 표=self.SHARES)
        self.assertEqual(text, "가장 큰 항목은 KTX(700)이고, 전체 1,000의 70.0%입니다. 2위 새마을(200, 20.0%)의 3.5배입니다. 항목 5개 중 상위 3개가 전체의 96.0%입니다.")
        self.assertEqual(issues, [])
        self.assertIn("표!A,B", source)
        self.assertEqual(self.value("해설:표!A,B", 표=self.SHARES)[0], "수입 전체 1,000. KTX 70.0%, 새마을 20.0%, 무궁화 6.0%.")
        self.assertIn("The largest is KTX", self.value("reading:표!A,B", 표=self.SHARES)[0])

    def test_a_total_that_does_not_add_up_gives_no_share(self):
        rows = self.SHARES[:-1] + [["합계", 5000, "원"]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertNotIn("%", text)
        self.assertEqual(issues[0].code, "cross_total_differs")

    def test_a_rate_is_not_a_share_of_anything(self):
        rows = [["구분", "유수율"], ["가", 90.5], ["나", 80.0], ["다", 70.5]]
        text = self.value("읽는법:표!A,B", 표=rows)[0]
        self.assertEqual(text, "항목 3개 중 가장 높은 항목은 가(90.5), 가장 낮은 항목은 다(70.5)입니다.")
        self.assertEqual(self.code("구성비1:표!A,B", 표=rows), "cross_not_additive")

    def test_two_periods_read_what_moved(self):
        rows = [["구분", 2023, "2024"], ["생산", 100, 110], ["누수", 10, 15], ["판매", 90, 81], ["먼지", 0.1, 0.9], ["합계", 200.1, 206.9]]
        text = self.value("읽는법:표!A,B,C", 표=rows)[0]
        self.assertEqual(text, "전체는 200.1 → 206.9, +3.4%입니다. 가장 크게 늘어난 항목은 누수(10 → 15, +50.0%)입니다. 증가량이 가장 큰 항목은 생산(100 → 110, +10.0%)입니다. 가장 크게 줄어든 항목은 판매(90 → 81, -10.0%)입니다.")

    def test_two_rates_move_in_points(self):
        rows = [["구분", "비율", "비율"], ["유수율", 92.8, 91.0], ["누수율", 4.2, 5.0]]
        self.assertEqual(self.value("읽는법:표!A,B,C", 표=rows)[0], "가장 크게 늘어난 항목은 누수율(4.2 → 5, +0.8포인트)입니다. 가장 크게 줄어든 항목은 유수율(92.8 → 91, -1.8포인트)입니다.")

    def test_names_across_several_columns_and_a_total_left_of_them(self):
        rows = [["구분", None, "값"], ["유\n효", "소계", 30], [None, "요금", 20], [None, "분  수량", 10], ["합     계", None, 30]]
        text = self.value("읽는법:표!A:B,C", 표=rows)[0]
        self.assertIn("요금(20)", text)
        self.assertIn("전체 30의 66.7%", text)

    def test_grouping_by_name_with_a_condition(self):
        rows = [["연도", "부문", "강원", "경기"], [2023, "산업", 1, 2], [2023, "건물", 3, 4], [2024, "산업", 5, 6], [2024, "건물", 7, 8]]
        self.assertIn("건물(22)", self.value("읽는법:표!B,합계(C:D)", 표=rows)[0])
        self.assertEqual(self.value("읽는법:표!B,합계(C:D, A=2024)", 표=rows)[0], "가장 큰 항목은 건물(15)이고, 항목 합계 26의 57.7%입니다. 2위 산업(11, 42.3%)의 1.4배입니다.")
        self.assertEqual(self.value("전일대비:표!A,합계(C:D, B=산업)", 표=rows)[0], "+266.7%")  # 3 -> 11, by year

    def test_mixed_cells_in_a_comparison(self):
        rows = [["은행", "인력"], ["가", "362(37%)"], ["나", "105(27%)"], ["다", "-"]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertIn("가(362)", text)
        self.assertEqual(issues[0].code, "number_from_text")

    def test_a_total_under_a_name_it_does_not_know_is_not_an_item(self):
        rows = [["구분", "값"], ["가", 60], ["나", 30], ["다", 10], ["모두 더한 값", 100]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertIn("가장 큰 항목은 가(60)이고, 전체 100의 60.0%", text)
        self.assertEqual(issues[0].code, "cross_total_found")
        self.assertEqual(self.value("순위1:표!A,B", 표=[["구분", "값"], ["가", 60], ["나", 30], ["총합계", 90]])[0], "가")

    def test_headings_as_a_reader_means_them(self):
        rows = [["번호", "은행", "IT인력", None], [None, None, "직원", "소계"], [1, "가", 3, 5], [2, "나", 1, 2], [3, "다", 1, 3]]
        self.assertTrue(self.value("해설:표!B,D", 표=rows)[0].startswith("IT인력 소계 항목 합계 10."))  # 5 = 2 + 3 among only 3 items is a coincidence, not a total
        coded = [["항목", "매출"], ["코드", "J468"], ["2026-01", 10], ["2026-02", 12], ["2026-03", 13]]
        self.assertTrue(self.value("해설:표!A,B", 표=coded)[0].endswith("13. 직전(2026년 2월)보다 8.3% 증가.") and "매출" in self.value("해설:표!A,B", 표=coded)[0])

    def test_a_sum_the_sheet_does_not_state_is_not_called_the_whole(self):
        rows = [["구분", "값"], ["가", 60], ["나", 30], ["다", 20]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertIn("항목 합계 110의 54.5%", text)
        self.assertNotIn("전체", text)
        self.assertEqual(issues[0].code, "cross_no_total")
        self.assertEqual(self.value("해설:표!A,B", 표=rows)[0], "값 항목 합계 110. 가 54.5%, 나 27.3%, 다 18.2%.")

    def test_an_index_on_a_sheet_named_for_it_is_not_a_share(self):
        rows = [["업종", "9월"], ["가", 177.2], ["나", 158.2], ["다", 120.5]]
        self.assertEqual(self.value("읽는법:생산지수!A,B", 생산지수=rows)[0], "항목 3개 중 가장 높은 항목은 가(177.2), 가장 낮은 항목은 다(120.5)입니다.")

    def test_a_rate_keeps_the_digits_that_explain_its_change(self):
        def fill(wb):
            sheet = wb.create_sheet("표")
            sheet.append(["구분", "비율", "비율"])
            for label, a, b in (("누수율", 4.2, 6.0), ("유수율", 92.8, 91.0)):
                sheet.append([label, a, b])
                sheet.cell(sheet.max_row, 2).number_format = sheet.cell(sheet.max_row, 3).number_format = "0"
        source = core.ExcelSource(make_workbook(os.path.join(self.dir, "rates.xlsx"), fill))
        self.assertIn("누수율(4.2 → 6, +1.8포인트)", source.text_value("읽는법:표!A,B,C")[0])

    def test_zero_filled_ahead_warns_and_is_not_a_newer_value(self):
        rows = [["월", "값", "비고값"], ["2026-01", 10, 1], ["2026-02", 12, 2], ["2026-03", 0, 0]]
        issues = core.ExcelSource(self.sheets(표=rows)).insight_request("통찰:표!A,B")[2]
        self.assertIn("latest_zero", [i.code for i in issues])
        rows = [["월", "값", "비고값"], ["2026-01", 10, 1], ["2026-02", 12, 2], ["2026-03", None, 0]]
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows)[2], [])

    def test_another_table_below_does_not_count_as_a_newer_value(self):
        rows = [["구분", "1월", "2월", "3월"], ["매출", 10, 20, None], ["비용", 1, 2, None], [None, None, None, None], ["전월비", 1.1, 1.2, 1.3]]
        rows[0] = ["구분", "2026-01", "2026-02", "2026-03"]
        self.assertEqual(self.value("전일대비:표!1,2", 표=rows)[2], [])

    def test_a_catch_all_is_not_ranked_with_the_named_items(self):
        rows = [["구분", "값"], ["가", 30], ["기타", 50], ["나", 20], ["합계", 100]]
        self.assertEqual(self.value("순위1:표!A,B", 표=rows)[0], "가")
        self.assertEqual(self.value("순위3:표!A,B", 표=rows)[0], "기타")
        self.assertIn("가장 큰 항목은 가(30)", self.value("읽는법:표!A,B", 표=rows)[0])

    def test_a_table_with_no_heading_over_the_names_leaves_it_blank(self):
        rows = [[None, "값"], ["가", 30], ["나", 20], ["합계", 50]]
        table = core.ExcelSource(self.sheets(표=rows)).range_cells("순위2:표!A,B")[0]
        self.assertEqual(table[0][0][1], "")

    def test_a_merged_name_covers_the_rows_it_spans(self):
        def fill(wb):
            sheet = wb.create_sheet("표")
            for row in (["국가", "구분", "값"], ["중국", "월계", 10], [None, "누계", 100], ["일본", "월계", 5], [None, "누계", 50], ["대만", "월계", 1], [None, "누계", 10]):
                sheet.append(row)
            for top in (2, 4, 6):
                sheet.merge_cells(start_row=top, end_row=top + 1, start_column=1, end_column=1)
        source = core.ExcelSource(make_workbook(os.path.join(self.dir, "merged.xlsx"), fill))
        self.assertIn("중국(100)", source.text_value("읽는법:표!A,합계(C, B=누계)")[0])

    def test_a_multiple_too_large_to_mean_anything_is_left_out(self):
        rows = [["구분", "값"], ["가", 3000000], ["나", 16], ["합계", 3000016]]
        text = self.value("읽는법:표!A,B", 표=rows)[0]
        self.assertNotIn("배입니다", text)

    def test_numbers_are_not_names(self):
        rows = [["번호", "값"], [71.2, 71.8], [68.1, 70.0], [-2.8, -0.9]]
        self.assertEqual(self.code("읽는법:표!A,B", 표=rows), "cross_no_items")

    def test_a_series_of_rates_moves_in_points(self):
        rows = [["Percentage change, year on year"], ["월", "All retailing"], ["2026-01", 3.0], ["2026-02", 4.0], ["2026-03", 2.0]]
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows)[0], "-2.0포인트")
        self.assertEqual(self.value("prev:표!A,B", 표=rows)[0], "-2.0 points")
        self.assertEqual(self.value("전일대비:가동률!A,B", 가동률=rows[1:])[0], "-2.0포인트")
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows[1:])[0], "-50.0%")
        self.assertEqual(self.value("읽는법:표!A,B", 표=rows)[0], "직전 2026년 2월 4 대비 -2.0포인트입니다.")
        self.assertEqual(self.value("해설:표!A,B", 표=rows)[0], "2026년 3월 All retailing 2. 직전 2026년 2월 4 대비 -2.0포인트입니다.")

    def test_the_fallback_sentence_for_a_rate_does_not_repeat_itself(self):
        rows = [["월", "통화성공률"]] + [[f"2025-{m:02d}", 90.0] for m in range(1, 13)] + [["2026-01", 88.0], ["2026-02", 85.5]]
        request = core.ExcelSource(self.sheets(표=rows)).insight_request("통찰:표!A,B")[0]
        self.assertEqual(request.fallback.count("직전 2026년 1월"), 1)

    def test_a_score_moves_in_points(self):
        rows = [["구분", "종합점수"], ["2021년", 87.4], ["2022년", 87.6], ["2023년", "(결측)"], ["2024년", "(결측)"]]
        text, source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertEqual(text, "그 뒤 2개 기간은 '(결측)'이라 2022년 값을 썼습니다. 직전 2021년 87.4 대비 +0.2포인트입니다.")
        self.assertEqual(issues[0].code, "latest_text_newer")
        self.assertEqual(issues[0].params, {"used": "2022-01-01", "count": 2, "text": "(결측)"})
        self.assertEqual(self.value("최신:표!A,B", 표=rows)[2][0].code, "latest_text_newer")

    def test_a_note_in_a_label_column_is_not_a_newer_value(self):
        rows = [["월", "큐", "건수"], [dt.datetime(2026, 1, 1), "민원", 10], [dt.datetime(2026, 2, 1), "민원", 20], [dt.datetime(2026, 3, 1), "민원", None]]
        self.assertEqual(self.value("전일대비:표!A,C", 표=rows)[2], [])

    def test_counting_rows_leaves_the_heading_out(self):
        rows = [["부서", "나이", "이탈"]] + [["영업", 30, "Yes"]] * 2 + [["영업", 31, "No"]] * 2 + [["개발", 40, "No"]] * 5 + [["개발", 41, "Yes"]]
        text = self.value("읽는법:표!A,건수(B)", 표=rows)[0]
        self.assertIn("개발(6)", text)
        self.assertIn("항목 합계 10의", text)
        self.assertIn("건수 항목 합계 10.", self.value("해설:표!A,건수(B)", 표=rows)[0])

    def test_a_count_with_a_condition_keeps_its_first_row(self):
        rows = [["부서", "나이", "이탈"], ["영업", 30, "Yes"], ["영업", 31, "No"], ["개발", 40, "Yes"], ["개발", 41, "Yes"]]
        self.assertIn("개발(2)", self.value("읽는법:표!A,건수(B, C=Yes)", 표=rows)[0])
        self.assertIn("항목 합계 3의", self.value("읽는법:표!A,건수(B, C=Yes)", 표=rows)[0])

    def test_a_rate_by_group(self):
        rows = [["부서", "이탈"]] + [["영업", "Yes"]] * 2 + [["영업", "No"]] * 8 + [["개발", "Yes"]] + [["개발", "No"]] * 9 + [["인사", "Yes"]] * 3 + [["인사", "No"]] * 7
        text = self.value("읽는법:표!A,비율(B=Yes)", 표=rows)[0]
        self.assertEqual(text, "항목 3개 중 가장 높은 항목은 인사(30.0%), 가장 낮은 항목은 개발(10.0%)입니다.")
        self.assertEqual(self.value("순위1:표!A,rate(B=Yes)", 표=rows)[0], "인사")
        self.assertEqual(self.code("구성비1:표!A,비율(B=Yes)", 표=rows), "cross_not_additive")

    def test_a_mean_by_group(self):
        rows = [["부서", "급여"], ["영업", 100], ["영업", 200], ["개발", 300], ["개발", 500]]
        self.assertEqual(self.value("읽는법:표!A,평균(B)", 표=rows)[0], "항목 2개 중 가장 높은 항목은 개발(400), 가장 낮은 항목은 영업(150)입니다.")

    def test_a_rate_and_a_mean_by_date(self):
        rows = [["일자", "이탈", "값"], ["2026-01-01", "Yes", 10], ["2026-01-01", "No", 20], ["2026-02-01", "Yes", 30], ["2026-02-01", "Yes", 50]]
        self.assertEqual(self.value("최신:표!A,비율(B=Yes)", 표=rows)[0], "100.0%")
        self.assertEqual(self.value("전일대비:표!A,비율(B=Yes)", 표=rows)[0], "+50.0포인트")
        self.assertEqual(self.value("최신:표!A,평균(C)", 표=rows)[0], "40")

    def test_an_exact_name_does_not_take_in_longer_names(self):
        rows = [["월", "품목", "가격"], ["2026-01", "쌀", 100], ["2026-01", "찹쌀", 900], ["2026-02", "쌀", 110], ["2026-02", "찹쌀", 900]]
        self.assertEqual(self.value("최신:표!A,합계(C, B=쌀)", 표=rows)[0], "110")
        self.assertEqual(self.value("최신:표!A,평균(C, B=찹쌀)", 표=rows)[0], "900")
        rows = [["월", "큐", "건"], ["2026-01", "nidirect Contact Centre", 5], ["2026-01", "Back Office", 7]]
        self.assertEqual(self.value("최신:표!A,합계(C, B=nidirect)", 표=rows)[0], "5")  # no cell is exactly this, so a part is enough

    def test_a_column_of_ups_and_downs_changes_by_an_amount(self):
        rows = [["월", "무선국 증감"], ["2026-01", 1000], ["2026-02", 271576], ["2026-03", 512938]]
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows)[0], "+241,362")
        self.assertNotIn("포인트", self.value("읽는법:표!A,B", 표=rows)[0])
        self.assertEqual(self.value("전일대비:표!A,B", 표=[["월", "증감률"], ["2026-01", 1.5], ["2026-02", 2.0]])[0], "+0.5포인트")

    def test_a_base_at_its_peak_pulls_the_change_down(self):
        points = [facts.Point(dt.datetime(2025 + (m // 12), m % 12 + 1, 1), v) for m, v in enumerate([50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 90, 80])]
        text = next(f.text for f in facts.readings(points, str, limit=9) if f.id == "read_prev_extreme")
        self.assertIn("가장 높아", text)
        self.assertIn("-11.1%는 그만큼 낮게 나옵니다", text)

    def test_a_loss_is_not_points_and_is_not_hidden(self):
        rows = [["과목", "2024", "2025"], ["매출액", 100, 55], ["매출원가", 80, 40], ["영업이익", -9, -1], ["순이익", 5, 1]]
        text = self.value("읽는법:표!A,B,C", 표=rows)[0]
        self.assertNotIn("포인트", text)
        self.assertIn("가장 크게 줄어든 항목은 순이익(5 → 1, -80.0%)", text)
        self.assertIn("기준값이 음수라 비율을 낼 수 없는 항목: 영업이익(-9 → -1).", text)

    def test_only_the_top_level_of_a_numbered_table_is_compared(self):
        rows = [["과목", "2025"], ["Ⅰ. 자산", 100], ["1. 유동자산", 60], ["(1) 당좌자산", 40], ["(2) 재고자산", 20], ["2. 비유동자산", 40], ["Ⅱ. 부채", 70], ["1. 유동부채", 70], ["Ⅲ. 자본", 30]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertEqual(text, "항목 3개 중 가장 높은 항목은 Ⅰ. 자산(100), 가장 낮은 항목은 Ⅲ. 자본(30)입니다.")
        self.assertEqual(issues[0].code, "cross_outline")
        self.assertEqual(self.code("구성비1:표!A,B", 표=rows), "cross_not_additive")

    def test_several_totals_use_the_one_that_adds_up(self):
        rows = [["구분", "값"], ["세금", 90], ["기타", 10], ["Total revenue", 100], ["임금", 40], ["Total expenses", 40]]
        text, _source, issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertTrue(any(i.code == "cross_total_differs" for i in issues), text)
        rows = [["구분", "값"], ["세금", 90], ["임금", 10], ["Total", 100], ["합계 (참고)", None]]
        self.assertIn("전체 100의 90.0%", self.value("읽는법:표!A,B", 표=rows)[0])

    def test_a_period_with_fewer_rows_says_so_in_the_sentence(self):
        rows = [["연도", "월", "값"]]
        for year in range(2020, 2026):
            rows += [[year, month, 100] for month in range(1, 13)]
        rows += [[2026, month, 100] for month in range(1, 10)]
        text, _source, issues = self.value("읽는법:표!A,합계(C)", 표=rows)
        self.assertEqual(text, "2026년 값은 행이 9개로 평소 12개보다 적어 덜 쌓인 값으로 보입니다. 2025년 1,200 대비 -25.0%는 그대로 읽을 수 없고, 행당 평균으로 맞추면 0.0%입니다. 빠진 행의 계절 차이는 반영되지 않습니다.")
        self.assertIn("latest_group_size", [i.code for i in issues])
        self.assertIn("덜 쌓인 값", self.value("해설:표!A,합계(C)", 표=rows)[0])

    def test_a_date_in_two_columns(self):
        rows = [["연도", "월", "구분", "값"]]
        for year in (2025, 2026):
            for month in range(1, 10):
                rows += [[year, f"{month:02d}", "가", year - 2000 + month], [year, f"{month:02d}", "나", 1]]
        rows += [[2025, "10", "가", 1], [2025, "10", "나", 1]]
        text = self.value("읽는법:표!A+B,합계(D, C=가)", 표=rows)[0]
        self.assertIn("2025년 9월", text)
        self.assertEqual(self.value("최신:표!A+B", 표=rows)[0], "2026-09")
        table = core.ExcelSource(self.sheets(표=rows)).range_cells("최신2:표!A+B,합계(D, C=가)")[0]
        self.assertEqual([[cell[1] for cell in line] for line in table], [["연도 월", "값 (가)"], ["2026-08", "34"], ["2026-09", "35"]])
        self.assertEqual(self.code("최신:표!A+B+C+D,C", 표=rows), "latest_bad")

    def test_months_under_a_year_heading(self):
        rows = [["기간", "값"], ["Annual", None], [2024, 1000], [2025, 1100], ["Year 2025", None], ["January", 90], ["February", 80], ["Sept", 70], ["Year 2026", None], ["January", 99], ["February", 88], ["Rolling", None], [2025, 500], [2026, 600]]
        self.assertEqual(self.value("최신:표!A,B", 표=rows)[0], "88")
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows)[0], "-11.1%")

    def test_a_unit_row_is_not_the_name(self):
        rows = [["월", "Net Generation"], [None, "(Billion Kilowatthours)"], ["2026-01", 10], ["2026-02", 12], ["2026-03", 13]]
        self.assertTrue(self.value("해설:표!A,B", 표=rows)[0].startswith("2026년 3월 Net Generation (Billion Kilowatthours) 13."))

    def test_a_school_year_is_not_a_calendar_year(self):
        next_year = dt.date.today().year
        rows = [["학년도", "응시"], [next_year - 2, 100], [next_year - 1, 110], [next_year, 120]]
        self.assertEqual(self.value("전일대비:표!A,B", 표=rows)[2], [])
        rows[0][0] = "연도"
        issues = self.value("읽는법:표!A,B", 표=rows)
        self.assertEqual(issues[2][0].code, "latest_unfinished")
        self.assertIn("아직 끝나지 않은 기간", issues[0])

    def test_a_close_second_is_not_a_multiple(self):
        rows = [["구분", "값"], ["가", 104], ["나", 100], ["합계", 204]]
        self.assertIn("2위 나(100, 49.0%)와 차이가 크지 않습니다.", self.value("읽는법:표!A,B", 표=rows)[0])

    def test_scientific_text_is_a_number(self):
        self.assertEqual(core.as_number("-2.98E+11"), -298000000000)

    def test_text_figures_in_a_sum_are_reported(self):
        rows = [["월", "건수"], ["2026-01", "100(10%)"], ["2026-02", "150(15%)"]]
        self.assertEqual(self.value("최신:표!A,합계(B)", 표=rows)[2][0].code, "number_from_text")

    def test_years_as_text_across_a_row_with_part_of_a_year(self):
        rows = [["구분", "2019", "2020", "2021", "2022.10."], ["의사", 400, 410, 480, 490]]
        text, _source, issues = self.value("읽는법:표!1,2", 표=rows)
        self.assertIn("10개월치로 보입니다", text)
        self.assertEqual(issues[0].code, "latest_partial_year")

    def test_a_word_only_containing_rate_is_not_a_rate(self):
        self.assertTrue(facts.is_additive("Corporate operations revenue", "General", [1, 2]))
        self.assertFalse(facts.is_additive("Growth rate", "General", [1, 2]))
        self.assertFalse(facts.is_additive("생산지수", "General", [1, 2]))
        self.assertFalse(facts.is_additive("매출", "0.0%", [1, 2]))
        self.assertFalse(facts.is_additive("매출", "General", [1, -2]))

    def test_rows_with_figures_but_no_name_are_reported(self):
        rows = [["구분", "값"], ["가", 3], ["나", 1], [None, 7]]
        self.assertEqual(self.value("읽는법:표!A,B", 표=rows)[2][0].code, "cross_unlabeled")

    def test_too_few_items_stay_visible(self):
        self.assertEqual(self.code("읽는법:표!A,B", 표=[["구분", "값"], ["가", 3]]), "cross_no_items")

    def test_rank_words(self):
        self.assertEqual(self.value("순위1:표!A,B", 표=self.SHARES)[0], "KTX")
        self.assertEqual(self.value("순위2:표!A,B", 표=self.SHARES)[0], "새마을")
        self.assertEqual(self.value("순위값2:표!A,B", 표=self.SHARES)[0], "200")
        self.assertEqual(self.value("구성비1:표!A,B", 표=self.SHARES)[0], "70.0%")
        self.assertEqual(self.value("top1:표!A,B", 표=self.SHARES)[0], "KTX")
        self.assertEqual(self.code("순위9:표!A,B", 표=self.SHARES), "latest_fewer")

    def test_rank_tables(self):
        source = core.ExcelSource(self.sheets(표=self.SHARES))
        rows = source.range_cells("순위2:표!A,B")[0]
        self.assertEqual([[cell[1] for cell in line] for line in rows], [["구분", "수입"], ["KTX", "700"], ["새마을", "200"]])
        rows = source.range_cells("구성비2:표!A,B")[0]
        self.assertEqual([[cell[1] for cell in line] for line in rows], [["구분", "수입", "%"], ["KTX", "700", "70.0%"], ["새마을", "200", "20.0%"]])

    def test_a_rank_table_fills_a_slide_and_a_chart(self):
        frame = self.slide.shapes.add_table(2, 3, 0, 0, 914400, 914400)
        frame._element.xpath(".//p:cNvPr")[0].set("descr", "{{표:구성비3:표!A,B}}")
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("S", [1, 2])
        chart = self.slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 914400, 914400, 914400, data)
        chart._element.xpath(".//p:cNvPr")[0].set("descr", "{{차트:순위3:표!A,B}}")
        report, out = self.go(self.sheets(표=self.SHARES))
        self.assertEqual({item.status for item in report.items}, {"ok"})
        table = next(shape for shape in out.slides[0].shapes if shape.has_table).table
        self.assertEqual(table.cell(1, 0).text, "KTX")
        self.assertEqual(table.cell(3, 2).text, "6.0%")
        plot = next(shape for shape in out.slides[0].shapes if shape.has_chart).chart.plots[0]
        self.assertEqual(list(plot.categories), ["KTX", "새마을", "무궁화"])
        self.assertEqual(list(plot.series[0].values), [700, 200, 60])

    def test_an_insight_on_a_table_without_dates(self):
        request, _source, _issues = core.ExcelSource(self.sheets(표=self.SHARES)).insight_request("통찰:표!A,B")
        ids = [fact.id for fact in request.facts]
        self.assertIn("read_top", ids)
        self.assertIn("item_1", ids)
        self.assertEqual(request.lang, "ko")
        self.assertTrue(request.fallback.startswith("수입 전체 1,000."))

    def test_a_reference_with_a_date_line_is_never_read_as_a_comparison(self):
        rows = [["일자", "수송"], ["2026-01-10", 1], ["2026-01-11", 2], ["2026-01-12", 3]]
        self.assertIn("폭이 큽니다", self.value("읽는법:표!A,B", 표=rows)[0])

    def test_bad_references(self):
        for key in ("최신:표!A,5", "최신:표!합계(C),B", "전일대비:표!A,합계(C:)", "전일대비:표!A,합계(C, B)", "전일대비:표!A,합계(C,2=가)", "전일대비:표!B:A,C"):
            with self.subTest(key=key):
                self.assertEqual(self.code(key, 표=self.WIDE), "latest_bad")


class Suggestions(Base):
    """Nobody should have to learn the notation: from a workbook alone, a working reference and its real result."""

    def book(self, name, **tabs):
        """A workbook holding only these sheets, since the first sheet that reads well is what gets offered."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows in tabs.items():
            sheet = wb.create_sheet(title)
            for row in rows:
                sheet.append(row)
        path = os.path.join(self.dir, name)
        wb.save(path)
        return path

    def first(self, **tabs):
        source = core.ExcelSource(self.book("helper.xlsx", **tabs))
        shape, base, found = suggest.first_choice(source)
        return source, shape, base, {item.code: item for item in found}

    MONTHS = [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), 100 + m, 50] for m in range(14)]

    def test_a_tidy_series(self):
        _s, _shape, base, found = self.first(월별=[["월", "매출", "합계"]] + self.MONTHS)
        self.assertEqual(found["latest"].key, "최신:월별!A,C")  # the column named as a total comes first
        self.assertEqual(found["latest"].result, "50")
        self.assertEqual(base.kind, "date")
        self.assertIn("표:최신6:월별!A,C", found["table"].key)
        self.assertEqual(found["table"].place, "alt")
        self.assertEqual(found["latest"].typed, "{{최신:월별!A,C}}")

    def test_several_rows_per_date_add_up_and_offer_by_item(self):
        rows = [["월", "방향", "승차"]]
        for m in range(1, 7):
            rows += [[dt.datetime(2026, m, 1), "상행", 10 * m], [dt.datetime(2026, m, 1), "하행", 20 * m]]
        _s, _shape, base, found = self.first(표=rows)
        self.assertTrue(base.many_rows)
        self.assertEqual(found["latest"].key, "최신:표!A,합계(C)")
        self.assertEqual(found["latest"].result, "180")
        self.assertEqual(found["by_item"].key, "읽는법:표!B,합계(C, A=2026-06)")
        self.assertIn("하행(120)", found["by_item"].result)

    def test_a_log_of_events_is_counted(self):
        rows = [["번호", "등록일", "등록시간"]] + [[i, f"2026-01-{1 + i // 3:02d}", 1000 + i] for i in range(30)]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertEqual(found["latest"].key, "최신:표!B,건수(B)")
        self.assertEqual(found["latest"].result, "3")

    def test_year_and_month_in_two_columns(self):
        rows = [["연도", "월", "값"]] + [[2025 + m // 12, f"{m % 12 + 1:02d}", 100 + m] for m in range(15)]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.spec, "A+B")
        self.assertEqual(found["latest"].key, "최신:표!A+B,C")
        self.assertIn("2026년 3월", found["note"].result)

    def test_periods_across_a_row(self):
        rows = [["구분", "Jan. 2026", "Feb. 2026", "Mar. 2026", "Apr. 2026"], ["매출", 10, 20, 30, 40], ["비용", 1, 2, 3, 4]]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.spec, "1")
        self.assertEqual(found["latest"].key, "최신:표!1,2")
        self.assertEqual(found["latest"].result, "40")

    def test_two_tables_on_one_sheet_use_the_first(self):
        rows = [["월", "지수"]] + [[f"2026-0{m}-01", 100 + m] for m in range(1, 7)] + [[None, None], [None, None], ["개정", None], ["월", "개정치"]] + [[f"2026-0{m}-01", 0.1] for m in range(1, 7)]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.spec, "A2:A7")
        self.assertEqual(found["latest"].result, "106")
        self.assertEqual(base.notes[0][0], "note_first_table")

    def test_a_stray_repeat_reads_what_follows_and_says_so(self):
        rows = [["월", "값"]] + [[f"2026-{m:02d}-01", m] for m in range(1, 5)] + [["2026-01-01", 9]] + [[f"2026-{m:02d}-01", m] for m in range(5, 10)]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.spec, "A7:A11")
        self.assertEqual(base.notes[0][0], "note_repeat")
        self.assertEqual(found["latest"].result, "9")

    def test_sheets_that_continue_one_another(self):
        old = [["월", "값"]] + [[dt.datetime(2025, m, 1), m] for m in range(1, 13)]
        new = [["월", "값"]] + [[dt.datetime(2026, m, 1), m + 1] for m in range(1, 7)]
        _s, shape, _base, found = self.first(실적2025=old, 실적2026=new)
        self.assertEqual(shape.sheet, "실적2026")  # the newest year in a name comes first
        self.assertEqual(shape.glob, "실적202*")
        self.assertIn("전년 같은 달", found["reading"].result)

    def test_a_table_with_no_dates(self):
        rows = [["구분", "수입"], ["KTX", 700], ["새마을", 200], ["무궁화", 100], ["합계", 1000]]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.kind, "name")
        self.assertEqual(found["top"].result, "KTX")
        self.assertEqual(found["share"].result, "70.0%")
        self.assertIn("표:구성비5:표!A,B", found["table_share"].key)

    def test_two_headings_are_not_a_series(self):
        rows = [[None, None, "2025년", "2024년", "증감률"], ["일본", "월계", 10, 5, 100], ["일본", "누계", 100, 50, 100], ["중국", "월계", 30, 40, -25], ["중국", "누계", 300, 400, -25], ["대만", "월계", 20, 10, 100], ["대만", "누계", 200, 100, 100]]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.kind, "name")
        self.assertIn("합계(C, B=월계)", found["reading"].key)
        self.assertIn("중국(30)", found["reading"].result)

    def test_rows_of_people_are_counted_rather_than_their_ages_added(self):
        rows = [["부서", "나이"]] + [["영업", 30 + i] for i in range(6)] + [["개발", 40 + i] for i in range(4)]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertEqual(found["reading"].key, "읽는법:표!A,건수(B)")
        self.assertIn("영업(6)", found["reading"].result)

    def test_prices_of_many_things_take_one_thing(self):
        rows = [["월", "품목", "가격"]]
        for m in range(1, 7):
            rows += [[f"2026-{m:02d}", "쌀", 100 + m], [f"2026-{m:02d}", "콩", 50]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertEqual(found["latest"].key, "최신:표!A,평균(C, B=쌀)")
        self.assertEqual(found["latest"].result, "106")

    def test_two_periods_side_by_side_are_offered_first(self):
        rows = [["과목", "2025년(원)", "2024년(원)"], ["매출", 110, 100], ["비용", 80, 90], ["이익", 30, 10]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertEqual(found["change"].key, "읽는법:표!A,C,B")
        self.assertIn("이익(10 → 30, +200.0%)", found["change"].result)

    def test_financial_years_are_not_months(self):
        rows = [["FISCAL_YEAR", "수"]] + [[f"{y}/{(y + 1) % 100:02d}", 100 + y * y % 37] for y in range(2009, 2026)]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.count, 17)  # 2012/13 has no month 13, and it is still a year
        self.assertEqual(base.last, "2026-01-01")
        self.assertIn("2026년", found["note"].result)
        monthly = [["월", "수"]] + [[f"2026-{m:02d}", m * m] for m in range(1, 13)]
        _s, _shape, base, _found = self.first(표=monthly)
        self.assertEqual(base.first, "2026-01-01")  # real months stay months

    def test_a_row_for_everything_is_not_added_to_its_parts(self):
        rows = [["월", "지역", "병원", "대기"]]
        for m in range(1, 7):
            rows += [[f"2026-{m:02d}", "All Regions", "All Hospitals", 30], [f"2026-{m:02d}", "동부", "가", 10], [f"2026-{m:02d}", "서부", "나", 20]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertEqual(found["latest"].key, "최신:표!A,합계(D, B=All Regions, C=All Hospitals)")
        self.assertEqual(found["latest"].result, "30")

    def test_a_heading_is_not_a_value_to_choose(self):
        rows = [["연도", "공항", "노선", "여객"]] + [[2025, a, b, 10 + i * i] for i, (a, b) in enumerate([("인천", "국내선"), ("인천", "국제선"), ("김포", "국내선"), ("김포", "국제선"), ("김해", "국내선"), ("김해", "국제선")])]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(found["reading"].key, "읽는법:표!B,합계(D, C=국내선)")
        self.assertNotIn("노선", [v for g in base.groups for v in g.values])

    def test_figures_far_to_the_right_are_found(self):
        rows = [["년도", "월"] + [f"글자{i}" for i in range(10)] + ["운항"]]
        for m in range(1, 7):
            rows.append(["2026년", f"{m:02d}월"] + ["가"] * 10 + [m * 3])
            rows.append(["2026년", f"{m:02d}월"] + ["나"] * 10 + [m])
        _s, _shape, base, found = self.first(data=rows)
        self.assertEqual(base.spec, "A+B")
        self.assertEqual(found["latest"].result, "24")

    def test_years_then_quarters_read_the_quarters(self):
        rows = [["기간", "합계"]] + [[f"{y}–{y - 1999}", 700 + y] for y in range(2010, 2021)] + [["분기별", None]]
        rows += [[f"{m} quarter {y}", 100 + i * i] for i, (m, y) in enumerate([("September", 2021), ("December", 2021), ("March", 2022), ("June", 2022), ("September", 2022)])]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.notes[0][0], "note_last_form")
        self.assertEqual(found["latest"].result, "116")
        self.assertIn("2022년 3분기", found["note"].result)

    def test_dates_written_out(self):
        for raw, wanted in (("30 June 2017", (2017, 6, 30)), ("June 30, 2017", (2017, 6, 30)), ("March quarter 2026", (2026, 3, 1))):
            with self.subTest(raw=raw):
                self.assertEqual(core.as_date(raw), dt.datetime(*wanted))
        self.assertEqual(core.as_year("2007–08"), dt.datetime(2008, 1, 1))
        self.assertEqual(core.as_year("1999–00"), dt.datetime(2000, 1, 1))
        self.assertIsNone(core.as_date("31 June 2017"))

    def test_a_row_of_year_ends_is_a_series(self):
        rows = [["Year to", "Funds"]] + [[f"30 June {y}", 100 + (y - 2012) * 7] for y in range(2012, 2018)]
        _s, _shape, base, found = self.first(Table=rows)
        self.assertEqual(base.kind, "date")
        self.assertIn("2017", found["note"].result)

    def test_columns_of_years_start_from_the_newest(self):
        rows = [["사업", "2023년 결산", "2024년 결산", "2025년 결산"], ["도로", 10, 20, 30], ["교량", 5, 5, 6], ["터널", 1, 2, 3]]
        _s, _shape, base, found = self.first(표=rows)
        self.assertEqual(base.values[0].spec, "D")
        self.assertEqual(found["change"].key, "읽는법:표!A,C,D")

    def test_rows_that_are_parts_of_another_are_not_its_rivals(self):
        items = [facts.Item(*pair) for pair in (("본부", 100), ("가", 60), ("나", 40), ("다른 본부", 80), ("- 세부", 10))]
        self.assertEqual([item.label for item in facts.without_parts(items)], ["본부", "다른 본부"])
        self.assertIsNone(facts.without_parts([facts.Item(*pair) for pair in (("가", 700), ("나", 200), ("다", 100))]))
        rows = [["구분", "건수"], ["본부", 100], ["가", 60], ["나", 40], ["지원", 70], ["재무", 30]]
        text, _source, issues = self.value_of("읽는법:표!A,B", rows)
        self.assertIn("가장 큰 항목은 본부(100)", text)
        self.assertNotIn("가(60)", text)
        self.assertEqual(issues[0].code, "cross_nested")

    def value_of(self, key, rows):
        return core.ExcelSource(self.book("value.xlsx", 표=rows)).text_value(key)

    def test_what_fails_is_not_offered(self):
        rows = [["월", "값"]] + [[dt.datetime(2026, m, 1), m] for m in range(1, 5)]  # monthly: last week and the highest day of 30 can work, but nothing here is 7 days apart
        _s, _shape, _base, found = self.first(표=rows)
        self.assertNotIn("week", found)

    def test_a_sheet_of_notes_is_passed_over(self):
        notes = [["안내"], ["이 파일은 시험용입니다"], ["문의 1"]]
        _s, shape, _base, _found = self.first(표지=notes, 월별=[["월", "값"]] + [[dt.datetime(2026, m, 1), m] for m in range(1, 5)])
        self.assertEqual(shape.sheet, "월별")

    def test_nothing_readable_says_so(self):
        _s, shape, base, found = self.first(표=[["메모"], ["그냥 글자"]])
        self.assertIsNone(shape)
        self.assertEqual(found, {})

    def test_a_starter_template_fills_with_no_problems(self):
        rows = [["월", "매출"]] + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), 100 + m] for m in range(14)]
        excel = self.book("한글 파일.xlsx", 월별=rows)
        source = core.ExcelSource(excel)
        _shape, _base, found = suggest.first_choice(source)
        path = suggest.starter_path(self.dir, excel)
        self.assertTrue(path.endswith("한글 파일_template.pptx"))
        suggest.make_starter(path, "한글 파일", found)
        self.assertTrue(suggest.starter_path(self.dir, excel).endswith("한글 파일_template_2.pptx"))  # nothing is overwritten
        report = core.run(path, excel, os.path.join(self.dir, "out.pptx"))
        self.assertEqual(report.count("error"), 0, [(i.key, [x.code for x in i.issues]) for i in report.items])
        self.assertEqual({item.kind for item in report.items}, {"text", "table", "chart"})
        slides = Presentation(os.path.join(self.dir, "out.pptx")).slides
        words = " ".join(shape.text_frame.text for shape in slides[0].shapes if shape.has_text_frame)
        self.assertNotIn("{{", words)
        self.assertIn("2026년 2월", words)
        self.assertTrue(any(shape.has_chart for shape in slides[1].shapes))
        self.assertTrue(any(shape.has_table for shape in slides[1].shapes))

    def test_english_words_follow_the_language(self):
        source = core.ExcelSource(self.book("en.xlsx", Daily=[["Date", "Sales"]] + [[dt.datetime(2026, 1, d), d] for d in range(1, 15)]))
        _shape, _base, found = suggest.first_choice(source, lang="en")
        self.assertEqual({item.code: item.key for item in found}["latest"], "latest:Daily!A,B")
        self.assertTrue(any(item.code == "week" and item.key == "week:Daily!A,B" for item in found))


class Messages(unittest.TestCase):
    def test_every_issue_message_can_be_shown_with_a_key_of_its_own(self):
        import i18n
        for lang in ("ko", "en"):
            self.assertIn("없는표기", i18n.text(lang, "issue_latest_bad", key="없는표기"))

    def test_every_issue_a_reference_can_raise_has_a_message_in_both_languages(self):
        import i18n
        for code in core.ERRORS:
            for lang in ("ko", "en"):
                self.assertIn(f"issue_{code}", i18n.STRINGS[lang], code)


class NegativeZero(unittest.TestCase):
    """As real Excel shows them: checked against Excel by hand. An automatic minus goes when the figure rounds to zero; one written in the format stays."""

    def test_as_excel_shows_them(self):
        for value, fmt, wanted in ((-0.00038, "0.0%", "0.0%"), (-0.4, "0", "0"), (-0.04, "0.0", "0.0"), (-0.4, "#,##0", "0"), (-0.4, "#,##0.0", "-0.4"),
                                   (-0.04, "#,##0.0", "0.0"), (-0.4, "0;(0)", "(0)"), (-0.4, "0;-0", "-0"), (-0.004, "0.00%", "-0.40%")):
            with self.subTest(value=value, fmt=fmt):
                self.assertEqual(format_value(value, fmt)[0], wanted)

    def test_a_percent_in_a_sentence_has_no_negative_zero(self):
        self.assertEqual(facts.percent(-0.0003), "0.0%")
        self.assertEqual(facts.percent(0.0003), "0.0%")
        self.assertEqual(facts.percent(-0.1), "-10.0%")


class Readings(unittest.TestCase):
    """Fixed rules for how to read a figure. A weekday level of 100 and a weekend level of 20 run through these."""

    WEEK = [100, 101, 99, 100, 102, 20, 21]

    def ids(self, values, **options):
        return [fact.id for fact in facts.readings(daily(*values), lambda v: f"{v:,.0f}", **options)]

    def text(self, values, wanted):
        return next(fact.text for fact in facts.readings(daily(*values), lambda v: f"{v:,.0f}", limit=9) if fact.id == wanted)

    def test_a_weekend_in_the_mean_is_called_out_and_the_right_comparison_named(self):
        values = self.WEEK * 3 + [90]  # the newest is a Monday
        self.assertEqual(self.text(values, "read_conflict"),
                         "직전 7개 평균 대비 +16.0%는 주말 2개가 섞인 기준이라 높게 나올 수 있으니, 같은 요일 기준 -10.0%로 읽으세요.")

    def test_no_conflict_when_both_comparisons_agree(self):
        self.assertNotIn("read_conflict", self.ids(self.WEEK * 3 + [120], limit=9))

    def test_no_conflict_when_the_weekend_is_not_different(self):
        flat = [100, 100, 100, 100, 100, 95, 95]
        self.assertNotIn("read_conflict", self.ids(flat * 3 + [97], limit=9))

    def test_the_day_after_a_weekend_warns_about_the_previous_value(self):
        values = self.WEEK * 3 + [100]
        self.assertEqual(self.text(values, "read_prev_weekend"),
                         "직전 값 대비 +376.2%는 평일과 주말을 비교한 수치라 요일 효과가 큽니다. 같은 요일 기준 0.0%를 보세요.")

    def test_two_weekdays_in_a_row_get_no_such_warning(self):
        self.assertNotIn("read_prev_weekend", self.ids(self.WEEK * 3 + [100, 101], limit=9))

    def test_a_value_outside_its_weekday_range(self):
        steady = [100, 102, 98, 100, 102, 20, 21]
        self.assertEqual(self.text(steady * 3 + [60], "read_unusual"), "이번 값은 직전 3번 월요일의 중앙값 100보다 40.0% 낮아, 평소 범위를 벗어난 수준입니다. 마지막 기간의 집계가 끝났는지 확인하세요.")
        self.assertEqual(self.text(steady * 3 + [140], "read_unusual"), "이번 값은 직전 3번 월요일의 중앙값 100보다 40.0% 높아, 평소 범위를 벗어난 수준입니다.")

    def test_a_perfectly_steady_series_can_still_show_a_real_change(self):
        self.assertIn("read_unusual", self.ids([100, 100, 100, 100, 100, 20, 20] * 3 + [60], limit=9))

    def test_one_holiday_does_not_move_the_usual_level(self):
        monday = [100, 100, 10, 100, 100]  # a holiday falls on the third Monday
        values = [v for m in monday for v in [m, 101, 99, 100, 102, 20, 21]]
        self.assertEqual(facts.same_weekday(daily(*values, 100))[0], 100)
        self.assertNotIn("read_unusual", self.ids(values + [100], limit=9))

    def test_last_week_being_odd_changes_which_comparison_to_trust(self):
        week = [100, 101, 99, 100, 102, 20, 21]
        values = week * 3 + [30, 101, 99, 100, 102, 20, 21, 100]  # last Monday was far below usual, this Monday is ordinary
        text = self.text(values, "read_week_odd")
        self.assertEqual(text, "전주 같은 요일 1월 26일(월) 30가 그 요일 평소보다 70.0% 낮았으니, 전주 대비 +233.3% 대신 평소 월요일 중앙값 대비 0.0%로 읽으세요.")

    def test_the_same_advice_is_not_repeated(self):
        found = self.ids([100, 101, 99, 100, 102, 20, 21] * 3 + [30, 101, 99, 100, 102, 20, 21, 100], limit=9)
        self.assertIn("read_week_odd", found)
        self.assertNotIn("read_conflict", found)
        self.assertEqual(self.ids(self.WEEK * 3 + [90], limit=9).count("read_prev_weekend"), 0)  # read_conflict said it already

    def test_a_middling_day_is_not_called_usual(self):
        values = self.WEEK * 4 + [100, 107]  # a Tuesday 6% above its usual level: neither usual nor outside the range
        self.assertEqual(self.ids(values)[0], "read_nothing")
        self.assertTrue(self.text(values, "read_nothing").startswith("크게 짚을 점은 없습니다."))

    def test_enough_data_always_gets_some_reading_so_no_placeholder_is_left(self):
        for seed in range(30):
            values = [100 + (i * (seed + 3)) % 17 for i in range(40)]
            self.assertTrue(self.ids(values), f"no reading for seed {seed}")
        monthly = [facts.Point(dt.datetime(2025 + (m // 12), m % 12 + 1, 1), 100 + (m * 7) % 5) for m in range(16)]
        self.assertTrue(facts.readings(monthly, str))

    def test_an_odd_day_that_barely_moves_the_mean_is_not_named(self):
        values = self.WEEK * 3 + [100, 101, 99, 78, 102, 20, 21, 80]  # 22% below usual, but close to the mean, so leaving it out changes little
        self.assertNotIn("read_odd_day", self.ids(values, limit=9))

    def test_a_tiny_move_in_a_very_steady_series_is_not_unusual(self):
        steady = [1000, 1000, 1000, 1000, 1000, 500, 500]
        values = [v + (i % 3 == 0) for i, v in enumerate(steady * 3)] + [1004]
        self.assertNotIn("read_unusual", self.ids(values, limit=9))

    def test_an_odd_day_in_the_mean_says_how_much_it_moved_the_comparison(self):
        values = self.WEEK * 3 + [100, 101, 99, 40, 102, 20, 21, 100]  # the Thursday before is far below usual
        text = self.text(values, "read_odd_day")
        self.assertTrue(text.startswith("1월 29일(목) 40는 그 요일 평소 100보다 60.0% 낮아 직전 평균을 끌어내렸을 수 있습니다."), text)
        self.assertIn("그날을 빼면 직전 평균 대비는", text)
        self.assertIn("그날 사정을 확인하세요.", text)

    def test_the_newest_day_is_not_also_reported_as_an_odd_day(self):
        self.assertNotIn("read_odd_day", self.ids(self.WEEK * 3 + [40], limit=9))

    def test_weeks_in_a_row(self):
        values = [10] * 7 + [20] * 7 + [30] * 7 + [40] * 7
        self.assertEqual(facts.weeks_in_a_row(daily(*values)), 3)
        self.assertEqual(self.text(values, "read_weeks"), "최근 7개 합계가 3주 연속 증가했습니다.")
        self.assertEqual(facts.weeks_in_a_row(daily(*[40] * 7, *[30] * 7, *[35] * 7)), 1)

    def test_highest_and_lowest_of_the_recent_values(self):
        self.assertIn("read_extreme", self.ids(self.WEEK * 3 + [500], limit=9))
        self.assertIn("가장 낮은", self.text(self.WEEK * 3 + [1], "read_extreme"))

    def test_an_ordinary_day_says_it_is_ordinary(self):
        values = self.WEEK * 4 + [100, 100]
        self.assertEqual(self.ids(values), ["read_usual"])
        self.assertTrue(self.text(values, "read_usual").startswith("같은 요일 평소 수준입니다."))

    def test_monthly_data_warns_when_the_month_before_was_extreme(self):
        points = [facts.Point(dt.datetime(2025 + (m // 12), m % 12 + 1, 1), v) for m, v in enumerate(
            [50, 60, 55, 58, 62, 59, 61, 57, 63, 60, 58, 56, 54, 90, 20, 70])]
        found = {fact.id: fact.text for fact in facts.readings(points, str, limit=9)}
        self.assertEqual(found["read_prev_extreme"], "직전 달(2026년 3월) 값이 최근 12개 중 가장 낮아, 직전 달 대비 +250.0%는 그만큼 높게 나옵니다.")
        self.assertEqual(found["read_year"], "전년 같은 달 2025년 4월 58 대비 +20.7%입니다.")

    def monthly(self, values, start=2024):
        return [facts.Point(dt.datetime(start + (m // 12), m % 12 + 1, 1), v) for m, v in enumerate(values)]

    def test_monthly_data_reads_by_the_same_month_last_year_first(self):
        fact = facts.readings(self.monthly([100] * 12 + [105, 107, 106]), str)[0]
        self.assertEqual(fact.id, "read_year")
        self.assertEqual(fact.text, "전년 같은 달 2024년 3월 100 대비 +6.0%입니다.")

    def test_month_on_month_up_and_year_on_year_down_says_which_to_read(self):
        values = [150] * 20 + [60, 82]  # up from the month before, but far below the year before
        fact = facts.readings(self.monthly(values), str)[0]
        self.assertEqual(fact.id, "read_year_conflict")
        self.assertIn("직전 달 대비 +36.7%와 전년 같은 달 2024년 10월 150 대비 -45.3%의 방향이 다릅니다.", fact.text)
        self.assertIn("전년 같은 달 기준으로 읽으세요.", fact.text)
        self.assertNotIn("집계가 끝났는지", fact.text)  # it rose from the month before, so it does not look like a month still being counted

    def test_a_month_that_fell_sharply_from_the_last_may_still_be_counting(self):
        fact = facts.readings(self.monthly([100] * 12 + [100, 100, 100, 100, 100, 50]), str)[0]
        self.assertEqual(fact.id, "read_year")
        self.assertIn("집계가 끝났는지", fact.text)

    def test_the_sentence_leaves_out_the_mean_once_the_year_before_is_known(self):
        text = facts.sentence(self.monthly([100] * 12 + [110, 120]), "매출", str)
        self.assertEqual(text, "2025년 2월 매출 120. 전년 같은 달보다 20.0% 증가, 직전(2025년 1월)보다 9.1% 증가.")

    def test_a_small_rise_and_fall_is_not_called_a_conflict(self):
        ids = {f.id for f in facts.readings(self.monthly([100] * 12 + [101, 102, 101, 100.5, 99.5]), str, limit=9)}
        self.assertNotIn("read_year_conflict", ids)

    def test_months_in_a_row_above_the_year_before(self):
        found = {f.id: f.text for f in facts.readings(self.monthly([100] * 12 + [101, 102, 103, 104]), str, limit=9)}
        self.assertEqual(found["read_year_run"], "전년 같은 달보다 4번 연속 높습니다.")

    def test_the_same_month_over_many_years(self):
        august = [1536, 1300, 1200, 530]
        values = [v for a in august for v in [500] * 7 + [a] + [500] * 4] + [510] * 7 + [563]
        found = {f.id: f.text for f in facts.readings(self.monthly(values, 2019), lambda v: f"{v:,.0f}", limit=9)}
        self.assertEqual(found["read_same_month"], "같은 달 기준 최고는 2019년 8월 1,536이고, 이번 값은 그 36.7% 수준입니다.")
        lowest = values[:-1] + [100]
        self.assertIn("5개 해 중 가장 낮습니다", {f.id: f.text for f in facts.readings(self.monthly(lowest, 2019), str, limit=9)}["read_same_month"])

    def test_months_without_a_year_before_say_so_rather_than_nothing(self):
        fact = facts.readings(self.monthly([100, 120, 90, 80, 95, 110, 105, 100, 92]), str)[0]
        self.assertEqual(fact.id, "read_no_year")
        self.assertEqual(fact.text, "전년 같은 달 값이 없어 계절 영향을 가릴 수 없습니다. 직전 달 대비 -8.0%입니다. 9개 중 7번째로 높습니다.")

    def test_quarters_read_like_months(self):
        points = [facts.Point(dt.datetime(2024 + (q // 4), q % 4 * 3 + 1, 1), 100 + q) for q in range(7)]
        self.assertEqual(facts.point_text(points[-1].date, points, "ko"), "2025년 3분기")
        self.assertEqual(facts.point_text(points[-1].date, points, "en"), "Q3 2025")
        self.assertEqual(facts.readings(points, str)[0].text, "전년 같은 분기 2024년 3분기 102 대비 +3.9%입니다.")

    def yearly(self, values, start=2010):
        return [facts.Point(dt.datetime(start + i, 1, 1), v) for i, v in enumerate(values)]

    def test_yearly_data_is_about_years(self):
        points = self.yearly([100, 110, 121, 140, 170])
        self.assertEqual(facts.sentence(points, "외래", str), "2014년 외래 170. 직전(2013년)보다 21.4% 증가.")
        found = {f.id: f.text for f in facts.readings(points, str, limit=9)}
        self.assertEqual(found["read_prev_year"], "직전 해 2013년 140 대비 +21.4%입니다.")
        self.assertEqual(found["read_year_run"], "4년 연속 증가했습니다.")
        self.assertEqual(found["read_extreme"], "최근 5개 중 가장 높은 값입니다.")
        self.assertFalse(any("요일" in fact.text or "1월 1일" in fact.text for fact in facts.fact_pack(points, "외래", str)))

    def test_a_part_of_a_year_is_not_compared_with_whole_years(self):
        points = self.yearly([815153, 879753], 2020) + [facts.Point(dt.datetime(2022, 10, 1), 702544)]
        self.assertEqual(facts.partial_year(points), 10)
        fact = facts.readings(points, facts.plain_number, limit=9)
        self.assertEqual([f.id for f in fact], ["read_partial_year"])
        self.assertEqual(fact[0].text, "2022년 10월까지 값은 10개월치로 보입니다. 12개월치인 2021년 879,753 대비 -20.1%는 기간이 달라 그대로 읽을 수 없고, 월평균으로 맞추면 -4.2%입니다.")
        self.assertEqual(facts.sentence(points, "외래", facts.plain_number), "2022년 10월까지 외래 702,544. 직전(2021년)보다 20.1% 감소, 월평균 기준 4.2% 감소.")
        pack = {f.id for f in facts.fact_pack(points, "외래", facts.plain_number)}
        self.assertIn("per_month", pack)
        self.assertFalse({"rank", "low", "streak", "avg7"} & pack)

    def test_a_running_total_that_restarts_each_january(self):
        one_year = [10 * m for m in range(1, 13)]
        points = self.monthly(one_year + [m * 12 for m in range(1, 13)])
        self.assertTrue(facts.year_to_date(points))
        found = facts.readings(points, str, limit=9)
        self.assertEqual([f.id for f in found], ["read_year_to_date"])
        self.assertEqual(found[0].text, "이 값은 매년 1월에 다시 시작하는 누계로 보여, 직전 달 대비로는 읽을 수 없습니다. 전년 같은 달 2024년 12월 120 대비 +20.0%입니다.")
        self.assertEqual(facts.sentence(points, "누계", str), "2025년 12월 누계 144. 전년 같은 달보다 20.0% 증가.")
        self.assertFalse({"prev", "avg7", "rank", "streak"} & {f.id for f in facts.fact_pack(points, "누계", str)})

    def test_a_seasonal_series_is_not_a_running_total(self):
        rising = [50 + m for m in range(12)]  # rises through each year and dips in January, but not to under half of December
        self.assertFalse(facts.year_to_date(self.monthly(rising + rising)))
        self.assertFalse(facts.year_to_date(self.monthly([100] * 14)))
        self.assertFalse(facts.year_to_date(self.monthly([10 * m for m in range(1, 13)])))  # one year alone shows no restart

    def test_a_year_ending_on_its_last_day_is_not_partial(self):
        points = [facts.Point(dt.datetime(2020 + i, 12, 31), 10 + i) for i in range(3)]
        self.assertIsNone(facts.partial_year(points))
        self.assertIsNone(facts.partial_year(self.yearly([1, 2, 3])))

    def test_a_large_move_is_never_nothing_to_note(self):
        mondays = [100, 130, 70, 110, 90]  # a Monday this noisy cannot be called outside its range, but -30% is still not nothing
        values = [v for monday in mondays for v in (monday, 100, 100, 100, 100, 50, 50)] + [70]
        found = {f.id: f.text for f in facts.readings(daily(*values), lambda v: f"{v:,.0f}", limit=9)}
        self.assertNotIn("read_nothing", found)
        self.assertEqual(found["read_large"], "직전 5번 월요일의 중앙값 100 대비 -30.0%입니다. 변동이 큰 요일이라 평소 범위 안이지만 폭은 큽니다.")
        values[-1] = 65  # lower than every Monday before it
        self.assertIn("집계가 끝났는지", next(f.text for f in facts.readings(daily(*values), str, limit=9) if f.id == "read_large"))

    def test_nothing_to_note_always_means_a_small_move(self):
        for seed in range(60):
            values = [100 + (i * (seed + 3)) % (seed + 5) * 4 for i in range(45)]
            points = daily(*values)
            usual = facts.same_weekday(points)
            if any(f.id == "read_nothing" for f in facts.readings(points, str, limit=9)):
                self.assertLess(abs(facts.ratio(values[-1], usual[0])), facts.LARGE_CHANGE, f"seed {seed}")

    def test_a_few_values_still_get_a_sentence(self):
        self.assertEqual(self.ids([100, 102, 101]), ["read_short"])
        self.assertEqual(self.ids([100, 102, 60]), ["read_large"])

    def test_figures_in_sentences_have_sensible_digits(self):
        for value, wanted in ((4930915.845, "4,930,916"), (115.884985, "115.9"), (3.14159, "3.14"), (0.123456, "0.123"), (253186, "253,186"), (7, "7"), (2.5, "2.5"), (-1234.5, "-1,235")):
            with self.subTest(value=value):
                self.assertEqual(facts.plain_number(value), wanted)

    def test_a_period_still_running_when_the_file_was_saved(self):
        months = self.monthly([100] * 6, 2026)  # the newest is June 2026
        self.assertTrue(facts.unfinished(months, dt.date(2026, 6, 15)))
        self.assertFalse(facts.unfinished(months, dt.date(2026, 6, 30)))
        self.assertFalse(facts.unfinished(months, dt.date(2026, 7, 2)))
        self.assertTrue(facts.unfinished(self.yearly([1, 2, 3], 2024), dt.date(2026, 6, 1)))
        self.assertEqual(facts.period_end(self.monthly([1, 2, 3], 2026)), dt.date(2026, 3, 31))
        self.assertEqual(facts.period_end(self.monthly([1] * 12, 2026)), dt.date(2026, 12, 31))

    def test_a_day_entered_on_its_own_day_only_warns_when_far_below_usual(self):
        ordinary = daily(*self.WEEK * 3 + [100])
        today = ordinary[-1].date.date()
        self.assertFalse(facts.unfinished(ordinary, today))
        self.assertTrue(facts.unfinished(daily(*self.WEEK * 3 + [60]), today))
        self.assertFalse(facts.unfinished(daily(*self.WEEK * 3 + [60]), today + dt.timedelta(days=1)))
        self.assertTrue(facts.unfinished(ordinary, today - dt.timedelta(days=1)))

    def test_monthly_data_gets_no_weekday_rules(self):
        points = [facts.Point(dt.datetime(2025 + (m // 12), m % 12 + 1, 1), 100 + m) for m in range(14)]
        self.assertFalse({"read_conflict", "read_unusual", "read_prev_weekend", "read_odd_day", "read_weeks"}
                         & {f.id for f in facts.readings(points, str, limit=9)})

    def test_too_little_data_gives_nothing(self):
        self.assertEqual(self.ids([5]), [])
        self.assertEqual(self.ids([5, 6]), [])

    def test_limit(self):
        values = self.WEEK * 3 + [100, 101, 99, 40, 102, 20, 21, 500]
        self.assertLessEqual(len(self.ids(values)), 3)

    def test_english(self):
        values = self.WEEK * 3 + [90]
        fact = facts.readings(daily(*values), lambda v: f"{v:,.0f}", "en", limit=9)
        self.assertIn("Read it by the same weekday: -10.0%.", next(f.text for f in fact if f.id == "read_conflict"))

    def test_rules_never_state_a_cause(self):
        values = self.WEEK * 3 + [100, 101, 99, 40, 102, 20, 21, 100]
        for fact in facts.readings(daily(*values), str, limit=9):
            self.assertNotRegex(fact.text, r"때문|원인|로 인해")

    def test_the_same_data_always_reads_the_same(self):
        values = self.WEEK * 3 + [90]
        self.assertEqual(facts.readings(daily(*values), str), facts.readings(daily(*values), str))


def pack_of(*pairs):
    return [facts.Fact(name, text, tuple(facts.NUMBER.findall(text))) for name, text in pairs]


PACK = pack_of(("value", "12월 31일(수) 수송 47,759"), ("week", "전주 같은 요일 12월 24일(수) 52,279 대비 -8.6%"),
               ("avg7", "직전 7개 평균 42,137 대비 +13.3%"), ("rank", "최근 30개 값 중 20번째로 높음"),
               ("signal_unusual", "같은 요일 기준으로 평소 범위를 벗어남: 표준편차의 -2.9배"))


def statement(text, *ids, hypothesis=""):
    return {"statement": text, "fact_ids": list(ids), "hypothesis": hypothesis}


class FakeModel:
    """Stands in for the local model server, so these tests need no model and no network."""

    def __init__(self, *replies, error=None):
        self.replies, self.error, self.calls = list(replies), error, []

    def __call__(self, path, body, timeout):
        self.calls.append(body)
        if self.error:
            raise self.error
        reply = self.replies[(len(self.calls) - 1) % len(self.replies)] if self.replies else []
        return {"response": reply if isinstance(reply, str) else json.dumps({"insights": reply}, ensure_ascii=False)}


class Router(unittest.TestCase):
    def test_counts_letters(self):
        self.assertEqual(facts.detect_language(["구리구간 수송인원", "일자"]), "ko")
        self.assertEqual(facts.detect_language(["Monthly sales", "Segment", "Date"]), "en")
        self.assertEqual(facts.detect_language(["수송인원 일별 보고", "A+B"]), "ko")
        self.assertEqual(facts.detect_language(["Sales report for the board", "매출"]), "en")

    def test_no_letters_falls_back_to_korean(self):
        self.assertEqual(facts.detect_language(["12345"]), "ko")

    def test_fingerprint_changes_with_the_facts(self):
        self.assertEqual(facts.fingerprint(PACK, "ko"), facts.fingerprint(PACK, "ko"))
        self.assertNotEqual(facts.fingerprint(PACK, "ko"), facts.fingerprint(PACK[:-1], "ko"))
        self.assertNotEqual(facts.fingerprint(PACK, "ko"), facts.fingerprint(PACK, "en"))


class SentenceChecks(unittest.TestCase):
    def problems(self, item, lang="ko"):
        return insight.check(item, PACK, lang)

    def test_a_sentence_using_its_cited_numbers_passes(self):
        self.assertEqual(self.problems(statement("전주 같은 요일 대비 8.6% 감소했고 직전 7개 평균 대비 13.3% 증가했습니다.", "week", "avg7")), [])

    def test_a_number_from_a_fact_it_did_not_cite_fails(self):
        self.assertEqual(self.problems(statement("직전 7개 평균 대비 8.6% 감소했습니다.", "avg7")), ["number"])

    def test_a_decimal_before_a_unit_is_still_checked(self):
        self.assertEqual(self.problems(statement("표준편차의 -2.9배입니다.", "signal_unusual")), [])
        self.assertEqual(self.problems(statement("표준편차의 3.5배입니다.", "signal_unusual")), ["number"])

    def test_dates_and_windows_are_not_figures(self):
        self.assertEqual(self.problems(statement("12월 31일은 최근 30개 값 중 20번째로 높습니다.", "rank")), [])

    def test_same_number_written_differently_passes(self):
        pack = pack_of(("week", "대비 +11.0%"))
        self.assertEqual(insight.check(statement("11% 증가했습니다.", "week"), pack, "ko"), [])

    def test_no_valid_citation_fails(self):
        self.assertIn("citation", self.problems(statement("감소했습니다.", "없는id")))

    def test_causes_belong_in_the_hypothesis(self):
        self.assertEqual(self.problems(statement("연말 연휴로 인해 감소했습니다.", "week")), ["cause"])
        self.assertEqual(self.problems(statement("감소했습니다.", "week", hypothesis="확인 필요: 연말 연휴")), [])
        self.assertEqual(self.problems(statement("Sales fell because of holidays.", "week"), "en"), ["cause"])

    def test_known_misreadings_fail(self):
        self.assertEqual(self.problems(statement("평년 대비 증가했습니다.", "avg7")), ["meaning"])
        self.assertEqual(self.problems(statement("최근 30개 중 20번째로 낮습니다.", "rank")), ["meaning"])

    def test_a_rank_that_says_low_is_fine_when_the_fact_says_low(self):
        pack = pack_of(("rank", "최근 30개 값 중 20번째로 낮음"))
        self.assertEqual(insight.check(statement("20번째로 낮습니다.", "rank"), pack, "ko"), [])

    def test_wrong_language_and_leaked_names_fail(self):
        self.assertEqual(self.problems(statement("下降了。", "week")), ["language"])
        self.assertEqual(self.problems(statement("The value fell sharply today.", "week")), ["language"])
        self.assertEqual(self.problems(statement("Sales 감소", "week"), "en"), ["language"])
        self.assertIn("leak", self.problems(statement("신호_same_kind에서 감소했습니다.", "week")))
        self.assertIn("leak", self.problems(statement("signal_conflict가 나타납니다.", "week")))


class Candidates(unittest.TestCase):
    def make(self, model, **options):
        with mock.patch.object(insight, "post", model):
            return insight.make_candidates(PACK, "ko", **options)

    def test_router_picks_the_model_and_only_checked_sentences_survive(self):
        good = statement("전주 같은 요일 대비 8.6% 감소했습니다.", "week", hypothesis="확인 필요: 연말 연휴")
        bad = statement("평년 대비 99% 증가했습니다.", "avg7")
        model = FakeModel([good, bad])
        result = self.make(model, runs=1)
        self.assertEqual(model.calls[0]["model"], insight.MODELS["ko"])
        self.assertEqual([c.statement for c in result.candidates], [good["statement"]])
        self.assertEqual(result.rejected, 1)
        self.assertEqual(result.hypotheses, ["확인 필요: 연말 연휴"])

    def test_english_goes_to_the_english_model(self):
        model = FakeModel([])
        pack = pack_of(("value", "Sales 10"))
        with mock.patch.object(insight, "post", model):
            insight.make_candidates(pack, "en", runs=1)
        self.assertEqual(model.calls[0]["model"], insight.MODELS["en"])

    def test_several_runs_are_merged_without_repeats_and_use_new_seeds(self):
        same = statement("8.6% 감소했습니다.", "week")
        other = statement("13.3% 증가했습니다.", "avg7")
        model = FakeModel([same], [same, other])
        result = self.make(model, runs=2, first_seed=5)
        self.assertEqual(len(result.candidates), 2)
        self.assertEqual([c["options"]["seed"] for c in model.calls], [5, 6])

    def test_citation_ids_are_fixed_by_the_schema(self):
        model = FakeModel([])
        self.make(model, runs=1)
        items = model.calls[0]["format"]["properties"]["insights"]["items"]
        self.assertEqual(items["properties"]["fact_ids"]["items"]["enum"], [f.id for f in PACK])

    def test_the_model_stays_within_the_memory_limit_settings(self):
        model = FakeModel([])
        self.make(model, runs=1)
        self.assertEqual(model.calls[0]["options"]["num_ctx"], 4096)

    def test_prompt_holds_the_facts(self):
        model = FakeModel([])
        self.make(model, runs=1)
        self.assertIn("week: 전주 같은 요일", model.calls[0]["prompt"])

    def test_broken_json_counts_as_nothing_not_a_crash(self):
        result = self.make(FakeModel("이건 JSON이 아닙니다"), runs=2)
        self.assertEqual(result.candidates, [])

    def test_citation_text_leaked_into_the_sentence_is_removed(self):
        model = FakeModel([statement("8.6% 감소했습니다 (fact_ids: week).", "week")])
        self.assertEqual(self.make(model, runs=1).candidates[0].statement, "8.6% 감소했습니다.")

    def test_no_server_and_no_model_are_explained(self):
        with mock.patch.object(insight, "post", side_effect=insight.InsightError("ollama_missing")):
            with self.assertRaises(insight.InsightError) as caught:
                insight.make_candidates(PACK, "ko", runs=1)
        self.assertEqual(caught.exception.code, "ollama_missing")

    def test_only_this_pc_is_ever_contacted(self):
        self.assertTrue(insight.URL.startswith("http://127.0.0.1:"))
        with mock.patch.object(insight.OPENER, "open", side_effect=urllib.error.URLError("refused")) as opened:
            with self.assertRaises(insight.InsightError) as caught:
                insight.post("generate", {}, 1)
        self.assertEqual(caught.exception.code, "ollama_missing")
        self.assertTrue(opened.call_args[0][0].full_url.startswith("http://127.0.0.1:"))

    def test_missing_model_names_the_model(self):
        missing = urllib.error.HTTPError("u", 404, "not found", None, None)
        with mock.patch.object(insight.OPENER, "open", side_effect=missing):
            with self.assertRaises(insight.InsightError) as caught:
                insight.post("generate", {"model": "abc"}, 1)
        self.assertEqual((caught.exception.code, caught.exception.params), ("model_missing", {"model": "abc"}))


class InsightPlaceholder(Base):
    KEY = "통찰:일별!A,B"
    ROWS = [["일자", "수송"]] + [[f"2026-01-{day:02d}", 100 + day] for day in range(5, 13)]

    def book(self, rows=None, name="insight.xlsx"):
        def fill(wb):
            sheet = wb.create_sheet("일별")
            for row in rows or self.ROWS:
                sheet.append(row)
        return make_workbook(os.path.join(self.dir, name), fill)

    def fill(self, insights=None, rows=None, name="insight.xlsx", key=None):
        if not self.slide.shapes:
            textbox_runs(self.slide, ["{{" + (key or self.KEY) + "}}"])
        self.prs.save(self.pptx)
        report = core.run(self.pptx, self.book(rows, name), self.out, insights=insights)
        return report, Presentation(self.out).slides[0].shapes[0].text_frame.text

    def chosen(self, report, text, key=None):
        request = report.requests[key or self.KEY]
        return {(key or self.KEY).casefold(): {"fingerprint": request.fingerprint, "text": text}}

    def test_core_never_touches_the_model_or_the_network(self):
        import ast
        with open(core.__file__, encoding="utf-8") as fh:
            names = {alias.name.split(".")[0] for node in ast.walk(ast.parse(fh.read())) if isinstance(node, ast.Import) for alias in node.names}
        self.assertFalse({"insight", "urllib", "http", "socket"} & names)
        with mock.patch.object(insight, "post", side_effect=AssertionError("called the model")):
            self.fill()

    def test_with_nothing_chosen_the_calculated_sentence_goes_in_and_it_says_so(self):
        report, text = self.fill()
        self.assertTrue(text.startswith("1월 12일(월) 수송 112."))
        item = self.item(report, self.KEY)
        self.assertEqual(item.issues[0].code, "insight_not_chosen")
        self.assertEqual(item.status, "warn")

    def test_request_carries_the_first_stage_results_for_the_model(self):
        report, _text = self.fill()
        request = report.requests[self.KEY]
        self.assertEqual(request.lang, "ko")
        self.assertIn("week", [fact.id for fact in request.facts])

    def test_chosen_sentence_goes_in_exactly_as_chosen(self):
        report, _text = self.fill()
        insights = self.chosen(report, "사람이 고른 문장입니다.")
        report, text = self.fill(insights)
        self.assertEqual(text, "사람이 고른 문장입니다.")
        self.assertEqual(self.item(report, self.KEY).status, "ok")

    def test_preview_and_report_hold_the_same_text(self):
        report, _text = self.fill()
        insights = self.chosen(report, "같은 문장")
        self.prs.save(self.pptx)
        preview = core.run(self.pptx, self.book(), insights=insights)
        self.assertEqual(self.item(preview, self.KEY).value, "같은 문장")

    def test_a_choice_for_old_data_is_not_used_on_new_data(self):
        report, _text = self.fill()
        insights = self.chosen(report, "어제 고른 문장")
        report, text = self.fill(insights, self.ROWS + [["2026-01-13", 300]], "newer.xlsx")
        self.assertNotIn("어제 고른 문장", text)
        self.assertTrue(text.startswith("1월 13일(화) 수송 300."))
        self.assertEqual(self.item(report, self.KEY).issues[0].code, "insight_stale")

    def test_the_word_insight_does_not_decide_the_language(self):
        report, text = self.fill(key="insight:일별!A,B")  # every word in the sheets and the template here is Korean
        self.assertEqual(report.requests["insight:일별!A,B"].lang, "ko")
        self.assertIn("수송", text)

    def test_english_template_and_sheet_route_to_english(self):
        def fill(wb):
            wb.remove(wb["요약"])
            wb.remove(wb["데이터 표"])
            wb.remove(wb["표시트"])
            sheet = wb.create_sheet("Daily")
            for row in [["Date", "Sales"]] + [[f"2026-01-{day:02d}", 100 + day] for day in range(5, 13)]:
                sheet.append(row)
        xlsx = make_workbook(os.path.join(self.dir, "english.xlsx"), fill)
        textbox_runs(self.slide, ["Weekly report for the board: {{insight:Daily!A,B}}"])
        self.prs.save(self.pptx)
        report = core.run(self.pptx, xlsx)
        self.assertEqual(report.requests["insight:Daily!A,B"].lang, "en")

    def test_bad_references_stay_visible(self):
        for key, code in (("통찰:일별!A", "calc_columns"), ("통찰:없는시트!A,B", "missing_sheet"), ("통찰:일별", "latest_bad")):
            with self.subTest(key=key):
                self.setUp()
                report, text = self.fill(key=key)
                self.assertEqual(self.item(report, key).issues[0].code, code)
                self.assertEqual(text, "{{" + key + "}}")
                self.assertEqual(report.requests, {})

    def test_formula_without_saved_result_is_reported(self):
        rows = [["일자", "수송"], ["2026-01-11", 1], ["2026-01-12", "=1+1"]]
        report, text = self.fill(rows=rows)
        self.assertEqual(self.item(report, self.KEY).issues[0].code, "formula_no_cache")
        self.assertEqual(text, "{{" + self.KEY + "}}")

    def test_choices_are_stored_beside_the_settings(self):
        path = core.settings_path(self.dir, core.INSIGHTS_FILE)
        self.assertEqual(os.path.dirname(path), self.dir)
        core.save_settings(path, {"key": {"fingerprint": "f", "text": "문장"}})
        self.assertEqual(core.load_settings(path)["key"]["text"], "문장")


class Issues(Base):
    def code(self, report, key):
        return self.item(report, key).issues[0].code

    def test_missing_name_stays_visible(self):
        textbox_runs(self.slide, ["값 {{없는이름}} 끝"])
        report, out = self.go()
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "값 {{없는이름}} 끝")
        self.assertEqual(self.code(report, "없는이름"), "missing_name")
        self.assertEqual(report.count("error"), 1)

    def test_missing_sheet_and_bad_reference(self):
        textbox_runs(self.slide, ["{{없는시트!A1}} {{요약!ZZZZ}}"])
        report, out = self.go()
        self.assertEqual(self.code(report, "없는시트!A1"), "missing_sheet")
        self.assertEqual(self.code(report, "요약!ZZZZ"), "bad_ref")
        self.assertIn("{{없는시트!A1}}", out.slides[0].shapes[0].text_frame.text)

    def test_range_for_text(self):
        textbox_runs(self.slide, ["{{요약!A1:B2}}"])
        report, _out = self.go()
        self.assertEqual(self.code(report, "요약!A1:B2"), "range_for_text")

    def test_formula_name_is_explained(self):
        textbox_runs(self.slide, ["{{동적}}"])
        report, _out = self.go()
        self.assertEqual(self.code(report, "동적"), "name_not_cells")

    def test_empty_cell_warns(self):
        textbox_runs(self.slide, ["[{{요약!H9}}]"])
        report, out = self.go()
        self.assertEqual(self.code(report, "요약!H9"), "empty_cell")
        self.assertEqual(self.item(report, "요약!H9").status, "warn")
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "[]")

    def test_formula_without_saved_result(self):
        def fill(wb):
            wb["요약"]["C1"] = "=1+1"
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        textbox_runs(self.slide, ["{{요약!C1}}"])
        report, out = self.go(xlsx)
        self.assertEqual(self.code(report, "요약!C1"), "formula_no_cache")
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "{{요약!C1}}")

    def test_error_value_in_text(self):
        def fill(wb):
            cell = wb["요약"]["C1"]
            cell.value, cell.data_type = "#DIV/0!", "e"
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        textbox_runs(self.slide, ["{{요약!C1}}"])
        report, out = self.go(xlsx)
        self.assertEqual(self.code(report, "요약!C1"), "error_value")
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "{{요약!C1}}")

    def test_unknown_format_warns(self):
        def fill(wb):
            wb["요약"]["C1"] = 0.5
            wb["요약"]["C1"].number_format = "# ?/?"
        xlsx = make_workbook(os.path.join(self.dir, "d2.xlsx"), fill)
        textbox_runs(self.slide, ["{{요약!C1}}"])
        report, _out = self.go(xlsx)
        self.assertEqual(self.code(report, "요약!C1"), "format_fallback")

    def test_table_marker_in_text_box_is_explained(self):
        textbox_runs(self.slide, ["{{table:매출범위}}"])
        report, out = self.go()
        self.assertEqual(self.code(report, "table:매출범위"), "marker_in_text")
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "{{table:매출범위}}")

    def test_same_key_is_one_item_with_all_places(self):
        textbox_runs(self.slide, ["{{매출}} {{매출}}"])
        second = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        textbox_runs(second, ["{{매출}}"])
        report, _out = self.go()
        self.assertEqual(len(report.items), 1)
        self.assertEqual(report.items[0].wheres, [("slide", 1), ("slide", 2)])

    def test_preview_matches_generation(self):
        textbox_runs(self.slide, ["{{매출}} {{없음}}"])
        add_table(self.slide, 2, 3, "{{table:매출범위}}")
        self.prs.save(self.pptx)
        view = lambda r: [(i.kind, i.key, i.value, i.source, i.status, i.wheres) for i in r.items]  # noqa: E731
        preview = core.run(self.pptx, self.xlsx)
        made = core.run(self.pptx, self.xlsx, self.out)
        self.assertEqual(view(preview), view(made))
        self.assertEqual(preview.output, "")
        self.assertEqual(made.output, self.out)

    def test_smartart_placeholder_is_reported(self):
        self.prs.save(self.pptx)
        patched = os.path.join(self.dir, "smart.pptx")
        with zipfile.ZipFile(self.pptx) as src, zipfile.ZipFile(patched, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = data.replace(b"</Types>", (
                        b'<Override PartName="/ppt/diagrams/data1.xml" ContentType='
                        b'"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/></Types>'))
                if info.filename == "ppt/slides/_rels/slide1.xml.rels":
                    data = data.replace(b"</Relationships>", (
                        b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        b'relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>'))
                dst.writestr(info, data)
            dst.writestr("ppt/diagrams/data1.xml", "<x>{{매출}}</x>")
        report = core.run(patched, self.xlsx)
        self.assertEqual(self.code(report, "/ppt/diagrams/data1.xml"), "unsupported_location")


class BadFiles(Base):
    def assert_fails(self, code, template=None, excel=None):
        self.prs.save(self.pptx)
        with self.assertRaises(core.SlideFillError) as ctx:
            core.run(template or self.pptx, excel or self.xlsx)
        self.assertEqual(ctx.exception.code, code)

    def write(self, name, data):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as fh:
            fh.write(data)
        return path

    def test_encrypted_or_drm_ole_excel(self):
        self.assert_fails("file_protected", excel=self.write("x.xlsx", core.OLE_MAGIC + b"\0" * 100))

    def test_encrypted_template(self):
        self.assert_fails("file_protected", template=self.write("x.pptx", core.OLE_MAGIC + b"\0" * 100))

    def test_foreign_wrapper_is_not_office(self):
        self.assert_fails("file_not_office", excel=self.write("x.xlsx", b"SOME-DRM-HEADER" + b"\0" * 100))

    def test_legacy_extensions(self):
        self.assert_fails("wrong_type_excel", excel=self.write("x.xls", core.OLE_MAGIC))
        self.assert_fails("wrong_type_template", template=self.write("x.ppt", core.OLE_MAGIC))

    def test_missing_file(self):
        self.assert_fails("file_missing", excel=os.path.join(self.dir, "none.xlsx"))

    def test_zip_that_is_not_a_workbook(self):
        path = os.path.join(self.dir, "x.xlsx")
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("hello.txt", "hi")
        self.assert_fails("file_unreadable", excel=path)


class KoreanBuiltinFormat(Base):
    def test_builtin_korean_date_id_is_not_a_bare_serial(self):
        """Korean Excel can store dates with built-in id 31, which openpyxl does not know."""
        plain = os.path.join(self.dir, "plain.xlsx")
        wb = openpyxl.Workbook()
        wb.active.title = "s"
        wb.active["A1"] = 46329
        wb.active["A1"].number_format = "0.000"
        wb.save(plain)
        patched = os.path.join(self.dir, "korean.xlsx")
        with zipfile.ZipFile(plain) as src, zipfile.ZipFile(patched, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "xl/styles.xml":
                    self.assertIn(b'numFmtId="164"', data)
                    data = data.replace(b'<xf numFmtId="164"', b'<xf numFmtId="31"')
                dst.writestr(info, data)
        textbox_runs(self.slide, ["{{s!A1}}"])
        _report, out = self.go(patched)
        self.assertEqual(out.slides[0].shapes[0].text_frame.text, "2026년 11월 03일")

    def test_table_matches_real_excel(self):
        """Displayed text recorded from Excel on Korean Windows for these built-in ids."""
        self.assertEqual(shown(46329.5875, core.KOREAN_BUILTIN_FORMATS[27]), "2026年 11月 03日")
        self.assertEqual(shown(46329.5875, core.KOREAN_BUILTIN_FORMATS[32]), "14시 06분")
        self.assertEqual(shown(46329.5875, core.KOREAN_BUILTIN_FORMATS[34]), "2026/11/03")
        self.assertEqual(shown(46329.5875, core.KOREAN_BUILTIN_FORMATS[35]), "2026-11-03")
        self.assertEqual(shown(46329.5875, core.KOREAN_BUILTIN_FORMATS[30]), "2026-11-03")


class Output(Base):
    @staticmethod
    def digest(path):
        with open(path, "rb") as fh:
            return hashlib.sha256(fh.read()).hexdigest()

    def test_sources_are_not_changed(self):
        textbox_runs(self.slide, ["{{매출}}"])
        self.prs.save(self.pptx)
        before = self.digest(self.pptx), self.digest(self.xlsx)
        core.run(self.pptx, self.xlsx, self.out)
        self.assertEqual((self.digest(self.pptx), self.digest(self.xlsx)), before)

    def test_preview_writes_nothing(self):
        self.prs.save(self.pptx)
        before = set(os.listdir(self.dir))
        core.run(self.pptx, self.xlsx)
        self.assertEqual(set(os.listdir(self.dir)), before)

    def test_output_name_never_reuses_existing_file(self):
        day = dt.date(2026, 11, 3)
        first = core.output_path(self.dir, self.pptx, day)
        self.assertEqual(os.path.basename(first), "template_2026-11-03.pptx")
        open(first, "w").close()
        second = core.output_path(self.dir, self.pptx, day)
        self.assertEqual(os.path.basename(second), "template_2026-11-03_2.pptx")
        open(second, "w").close()
        self.assertEqual(os.path.basename(core.output_path(self.dir, self.pptx, day)), "template_2026-11-03_3.pptx")

    def test_save_failure_is_reported_and_leaves_no_tmp(self):
        self.prs.save(self.pptx)
        bad = os.path.join(self.dir, "no_such_folder", "out.pptx")
        with self.assertRaises(core.SlideFillError) as ctx:
            core.run(self.pptx, self.xlsx, bad)
        self.assertEqual(ctx.exception.code, "cannot_save")


class Settings(Base):
    def test_round_trip(self):
        path = os.path.join(self.dir, "s.json")
        core.save_settings(path, {"template": "경로"})
        self.assertEqual(core.load_settings(path), {"template": "경로"})
        self.assertFalse(os.path.exists(path + ".tmp"))

    def test_missing_gives_empty(self):
        self.assertEqual(core.load_settings(os.path.join(self.dir, "none.json")), {})

    def test_broken_file_is_kept_aside(self):
        path = os.path.join(self.dir, "s.json")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("{broken")
        self.assertEqual(core.load_settings(path), {})
        self.assertTrue(os.path.exists(path + ".broken"))

    def test_saved_as_utf8(self):
        path = os.path.join(self.dir, "s.json")
        core.save_settings(path, {"언어": "ko"})
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(json.load(fh), {"언어": "ko"})


if __name__ == "__main__":
    unittest.main(verbosity=1)
