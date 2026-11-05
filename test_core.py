"""Unit tests for numfmt.py and core.py. Run: python -X utf8 test_core.py"""

import datetime as dt
import hashlib
import json
import os
import re
import shutil
import tempfile
import unittest
import urllib.error
import zipfile
from dataclasses import replace
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
import i18n
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
        for fmt in ("# ?/?", "[DBNum1]General", "000-0000", "[$-F800]dddd"):
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
        self.assertIn("끝난 기간 중 가장 최근은", issues[0])
        self.assertIn("아직 끝나지 않은 기간", self.value("해설:표!A,B", 표=rows)[0])

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
        frame = self.slide.shapes.add_table(2, 3, 0, 0, 8229600, 914400)
        frame._element.xpath(".//p:cNvPr")[0].set("descr", "{{표:구성비3:표!A,B}}")
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("S", [1, 2])
        chart = self.slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 914400 * 3, 914400, 914400, data)  # clear of the table once it has grown
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
        self.assertIn("2025-26", found["note"].result)  # a sentence calls it what the sheet does: it is not the year 2026
        self.assertNotIn("2026년", found["note"].result)
        monthly =[["월", "수"]] + [[f"2026-{m:02d}", m * m] for m in range(1, 13)]
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
        long_list = [facts.Item(f"회사{n}", 300 + n * 37 % 211) for n in range(150)]
        long_list[10:13] = [facts.Item("큰회사", 300), facts.Item("작은회사", 100), facts.Item("다른회사", 200)]  # adds up by chance
        self.assertIsNone(facts.without_parts(long_list))
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
        self.assertEqual(len(slides), 1)
        self.assertTrue(any(shape.has_chart for shape in slides[0].shapes))
        self.assertTrue(any(shape.has_table for shape in slides[0].shapes))

    def test_a_report_from_the_workbook_alone_reads_the_same_as_one_read_afresh(self):
        rows = [["월", "매출"]] + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), 100 + m] for m in range(14)]
        excel = self.book("엑셀만.xlsx", 월별=rows)
        source = core.ExcelSource(excel)
        shape, base, found = suggest.first_choice(source)
        path = os.path.join(self.dir, "엑셀만.pptx")
        suggest.make_starter(path, "엑셀만", found)
        reused = core.run(path, excel, source=source)
        fresh = core.run(path, excel)
        seen = lambda report: [(i.kind, i.key, i.value, i.status, i.source, i.wheres) for i in report.items]
        self.assertEqual(seen(reused), seen(fresh))
        self.assertEqual(suggest.describe(shape, suggest.default_value(base)), "월별 · 매출")

    def test_everything_a_laid_out_report_lists_has_a_plain_name(self):
        rows = [["월", "매출"]] + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), 100 + m] for m in range(14)]
        excel = self.book("이름.xlsx", 월별=rows)
        _shape, _base, found = suggest.first_choice(core.ExcelSource(excel))
        path = os.path.join(self.dir, "이름.pptx")
        suggest.make_starter(path, "이름", found)
        named = suggest.labels(found)
        report = core.run(path, excel)
        self.assertTrue(report.items)
        for item in report.items:
            self.assertIn(item.key.casefold(), named)

    def months(self, *columns, count=20):
        rows = [["월"] + [name for name, _start in columns]]
        return rows + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1)] + [start + m for _name, start in columns] for m in range(count)]

    def planned(self, name, **sheets):
        excel = self.book(name, **sheets)
        plan = suggest.plan_report(core.ExcelSource(excel))
        path = os.path.join(self.dir, name + ".pptx")
        suggest.make_report(path, "제목", plan.sections)
        report = core.run(path, excel, os.path.join(self.dir, name + "_out.pptx"))
        self.assertEqual(report.count("error"), 0, [(i.key, [x.code for x in i.issues]) for i in report.items])
        slides = Presentation(os.path.join(self.dir, name + "_out.pptx")).slides
        words = [" ".join(shape.text_frame.text for shape in slide.shapes if shape.has_text_frame) for slide in slides]
        return plan, slides, words, report

    def test_two_sides_of_a_whole_both_reach_the_report(self):
        plan, slides, words, report = self.planned("양쪽.xlsx", 매출=self.months(("내국인", 100), ("외국인", 500)))
        self.assertEqual(plan.what, "매출 · 내국인, 외국인")
        self.assertIn("내국인 119", words[0])
        self.assertIn("외국인 519", words[0])
        table = next(shape for shape in slides[0].shapes if shape.has_table).table
        self.assertEqual([table.cell(0, c).text for c in range(3)], ["월", "내국인", "외국인"])
        self.assertEqual(table.cell(len(table.rows) - 1, 2).text, "519")
        self.assertEqual(table.cell(len(table.rows) - 1, 2).text_frame.paragraphs[0].runs[0].font.size.pt, suggest.TABLE_POINTS)
        chart = next(shape for shape in slides[0].shapes if shape.has_chart).chart
        self.assertEqual([plot.name for plot in chart.plots[0].series], ["내국인", "외국인"])
        self.assertEqual({named[1] for named in suggest.labels(plan.items).values()}, {"내국인", "외국인", ""})
        for item in report.items:
            self.assertIn(item.key.casefold(), suggest.labels(plan.items))

    def test_a_named_total_speaks_for_its_parts(self):
        plan, _slides, words, _report = self.planned("합계.xlsx", 매출=self.months(("내국인", 100), ("외국인", 500), ("합계", 600)))
        self.assertEqual(plan.what, "매출 · 합계")
        self.assertEqual(plan.figures_left, 2)
        self.assertNotIn("내국인 119", words[0])

    def test_many_figures_are_left_out_and_counted(self):
        plan, _slides, _words, _report = self.planned("역.xlsx", 승차=self.months(*[(f"역{n}", 100 + n) for n in range(30)]))
        self.assertEqual(plan.what, "승차 · 역0")
        self.assertEqual(plan.figures_left, 29)

    def test_a_figure_on_another_scale_stays_off_the_chart_but_in_the_table(self):
        plan, slides, words, _report = self.planned("규모.xlsx", 운항=self.months(("편", 10), ("여객", 1_000_000)))
        self.assertIn("여객", words[0])
        table = next(shape for shape in slides[0].shapes if shape.has_table).table
        self.assertEqual(len(table.columns), 3)
        chart = next(shape for shape in slides[0].shapes if shape.has_chart).chart
        self.assertEqual([plot.name for plot in chart.plots[0].series], ["편"])

    def test_a_sheet_built_the_same_way_gets_its_own_slides(self):
        plan, slides, words, _report = self.planned("두시트.xlsx", 매출=self.months(("내국인", 100), ("외국인", 500)), 이용객=self.months(("내국인", 30), ("외국인", 70)))
        self.assertEqual(len(plan.sections), 2)
        self.assertEqual(plan.what, "매출 · 내국인, 외국인 / 이용객 · 내국인, 외국인")
        self.assertEqual(len(slides), 2)
        self.assertIn("이용객", words[1])
        self.assertIn("외국인 89", words[1])
        self.assertEqual(plan.sheets_left, 0)

    def test_a_sheet_built_differently_is_left_out_and_counted(self):
        plan, slides, _words, _report = self.planned("다른시트.xlsx", 매출=self.months(("내국인", 100), ("외국인", 500)), 인원=self.months(("정규직", 5), ("계약직", 6)))
        self.assertEqual(len(plan.sections), 1)
        self.assertEqual(plan.sheets_left, 1)
        self.assertEqual(len(slides), 1)

    def test_many_sheets_built_alike_are_not_picked_from_at_random(self):
        sheets = {f"{m}월": self.months(("인원", 100 + m)) for m in range(1, 13)}
        plan, slides, _words, _report = self.planned("월별시트.xlsx", **sheets)
        self.assertEqual(len(plan.sections), 1)
        self.assertEqual(plan.sheets_left, 11)
        self.assertEqual(len(slides), 1)

    def test_when_a_slide_is_full_the_readings_go_and_every_figure_stays(self):
        long = "가" * 400
        sentences = [suggest.Suggestion(code, f"{code}{n}", "text", result=long if code == "reading" else "짧은 문장") for n in range(4) for code in ("note", "reading")]
        kept = suggest._fit(sentences, suggest.TEXT_WIDTH)
        self.assertEqual([item.key for item in kept if item.code == "note"], ["note0", "note1", "note2", "note3"])
        self.assertLess(len(kept), len(sentences))
        self.assertLessEqual(sum(suggest._height(i, suggest.TEXT_WIDTH) + suggest.SENTENCE_GAP for i in kept), suggest.SLIDE_BOTTOM - suggest.CONTENT_TOP)

    def test_a_sheet_name_ending_in_a_space_can_be_read_and_referred_to(self):
        excel = self.book("공백.xlsx", **{"매출 ": self.months(("매출", 100))})
        source = core.ExcelSource(excel)
        self.assertEqual(suggest.read_shape(source, "매출 ").sheet, "매출 ")
        self.assertEqual(source.text_value("최신:매출 !A,B")[0], "119")
        self.assertEqual(source.text_value("최신:매출!A,B")[0], "119")

    def test_sheets_that_continue_one_another_are_one_section(self):
        rows = lambda year: [["월", "매출"]] + [[dt.datetime(year, m, 1), 100 + m] for m in range(1, 13)]
        plan, _slides, _words, _report = self.planned("연도.xlsx", **{"실적 2024": rows(2024), "실적 2025": rows(2025)})
        self.assertEqual(len(plan.sections), 1)
        self.assertEqual(plan.sheets_left, 0)

    def test_the_heading_says_what_was_counted(self):
        unit = [["연월별 실적 (단위: 백만원)", None], ["월", "매출"]] + [[dt.datetime(2025, m, 1), 100 + m] for m in range(1, 13)]
        plan, _slides, words, _report = self.planned("단위.xlsx", data=unit)
        self.assertEqual(plan.sections[0].heading, "연월별 실적 (단위: 백만원)")
        self.assertIn("단위: 백만원", words[0])
        plan, _slides, words, _report = self.planned("이름.xlsx", **{"매출액(단위 백만불)": self.months(("내국인", 100))})
        self.assertEqual(plan.sections[0].heading, "매출액(단위 백만불)")
        plan, _slides, _words, _report = self.planned("없음.xlsx", Sheet1=self.months(("매출", 100)))
        self.assertEqual(plan.sections[0].heading, "")

    def test_a_column_of_units_is_read_only_when_it_says_one_thing(self):
        rows = lambda *units: [["월", "매출", "단위"]] + [[dt.datetime(2025, m, 1), 100 + m, units[m % len(units)]] for m in range(1, 13)]
        source = core.ExcelSource(self.book("열.xlsx", Sheet1=rows("원")))
        self.assertEqual(suggest.heading(source, suggest.Shape("Sheet1")), "단위: 원")
        source = core.ExcelSource(self.book("여러열.xlsx", Sheet1=rows("원", "명")))
        self.assertEqual(suggest.heading(source, suggest.Shape("Sheet1")), "")
        source = core.ExcelSource(self.book("이름만.xlsx", Sheet1=[["월", "유통단계별단위명"], [dt.datetime(2025, 1, 1), "20kg"]]))
        self.assertEqual(suggest.heading(source, suggest.Shape("Sheet1")), "")

    def test_what_the_helper_lists_for_a_report_is_what_the_report_holds(self):
        excel = self.book("목록.xlsx", 매출=self.months(("매출", 100)))
        _shape, _base, found = suggest.first_choice(core.ExcelSource(excel))
        path = os.path.join(self.dir, "목록.pptx")
        suggest.make_starter(path, "목록", found)
        listed = {item.key.casefold() for item in suggest.in_report(found)}
        named = suggest.labels(suggest.in_report(found))
        self.assertTrue({item.code for item in found} - {item.code for item in suggest.in_report(found)}, "some are only for a template")
        report = core.run(path, excel)
        self.assertEqual({item.key.casefold() for item in report.items}, set(named))
        self.assertTrue(listed)

    def test_a_table_is_listed_by_its_rows_of_data(self):
        excel = self.book("크기.xlsx", 매출=self.months(("매출", 100)))
        _shape, _base, found = suggest.first_choice(core.ExcelSource(excel))
        path = os.path.join(self.dir, "크기.pptx")
        suggest.make_starter(path, "크기", found)
        report = core.run(path, excel)
        table = next(item for item in report.items if item.kind == "table")
        self.assertEqual(table.value, "6 × 2")
        self.assertEqual(table.value, next(item for item in found if item.code == "table").result)

    def test_measures_under_each_year_are_not_added_together(self):
        years = [["발생년"] + [cell for year in range(2020, 2026) for cell in (str(year), None)], ["유형"] + ["사고건수", "사망자수"] * 6]
        rows = years + [["차대차"] + [10, 1] * 6, ["합계"] + [n for year in range(6) for n in (100 + year, 7)]]
        plan, _slides, words, _report = self.planned("연도가로.xlsx", 사고=rows)
        self.assertIn("2025년 합계 (사고건수) 105", words[0])
        self.assertNotIn("112", words[0])

    def test_two_levels_under_each_year_are_refused_rather_than_guessed(self):
        head = [["발생년"] + [cell for year in range(2020, 2026) for cell in (str(year), None, None, None)], ["성별"] + ["남", None, "여", None] * 6, ["연령"] + ["사망자수", "부상자수"] * 12]
        source = core.ExcelSource(self.book("두단.xlsx", 사고=head + [["1세"] + [1, 2, 3, 4] * 6, ["2세"] + [2, 3, 4, 5] * 6, ["합계"] + [3, 5, 7, 9] * 6]))
        self.assertIsNone(suggest.plan_report(source))

    def test_a_column_of_days_is_not_a_figure(self):
        rows = [["지정일자", "설치일자", "기관명"]] + [[f"2020{m:02d}{d:02d}", f"2020{m:02d}{d:02d}", f"기관{m}{d}"] for m in range(1, 13) for d in (1, 15)]
        source = core.ExcelSource(self.book("날짜열.xlsx", 현황=rows))
        shape = suggest.read_shape(source, "현황")
        names = [value.name for base in shape.bases if base.kind == "date" for value in base.values]
        self.assertNotIn("설치일자", names)

    def merged_book(self, name, rows, merged):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "표"
        for row in rows:
            ws.append(row)
        for cells in merged:
            ws.merge_cells(cells)
        path = os.path.join(self.dir, name)
        wb.save(path)
        wb.close()
        return path

    def language_of(self, excel, title="자료"):
        source = core.ExcelSource(excel)
        plan = suggest.plan_report(source)
        return next(iter(suggest.ask_insights(source, plan.sections, title).values()))[0].lang

    def test_the_language_of_a_report_follows_what_the_sheet_holds(self):
        korean = [["지역", "시설", "입장객"]] + [["동구", f"{name} 공원", 100 + n * 37 % 91] for n, name in enumerate(("가람", "나루", "다솜", "라온", "마루", "바다"))]
        self.assertEqual(self.language_of(self.book("한글.xlsx", tourStatConfirmManageBizData=korean), "tourStatConfirmManageBiz_2021"), "ko")
        english = [["Region", "Site", "Visitors"]] + [["East", f"{name} Park", 100 + n * 37 % 91] for n, name in enumerate(("Oak", "Elm", "Ash", "Fir", "Yew", "Bay"))]
        self.assertEqual(self.language_of(self.book("영어.xlsx", 자료=english), "시설 현황 보고서"), "en")

    def test_short_korean_names_do_not_lose_to_long_english_ones(self):
        rows = [["회사", "구분", "매출"]] + [[name, "국내", 100 + n * 37 % 91] for n, name in enumerate(("Samsung Electronics", "Hyundai Motor Company", "Hanwha Solutions", "Kakao Corporation", "Naver Corporation", "Coupang Incorporated"))]
        self.assertEqual(self.language_of(self.book("혼합.xlsx", 표=rows)), "ko")

    def test_a_count_in_an_english_report_is_not_called_in_korean(self):
        rows = [["State", "Age"]] + [[state, 30 + n] for state, size in (("VIC", 9), ("NSW", 6), ("QLD", 4)) for n in range(size)]
        source = core.ExcelSource(self.book("셈.xlsx", 표=rows))
        plan = suggest.plan_report(source)
        request = next(iter(suggest.ask_insights(source, plan.sections, "Report").values()))[0]
        self.assertEqual(request.lang, "en")
        self.assertNotRegex(next(fact for fact in request.facts if fact.id == "subject").text, "건수")

    def doubts(self, rows=None, sheet="표", path=None):
        source = core.ExcelSource(path or self.book("의심.xlsx", **{sheet: rows}))
        return [code for code, _params in suggest.plan_report(source).doubts]

    def test_a_plain_table_raises_no_doubt(self):
        self.assertEqual(self.doubts([["지역", "매출"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(8)]), [])
        self.assertEqual(self.doubts([["월", "매출"]] + [[f"2025-{m:02d}", 100 + m * 37 % 91] for m in range(1, 10)]), [])

    def test_adding_up_values_that_are_nearly_the_same_raises_a_doubt(self):
        rows = [["연월", "회사", "발량"]] + [[f"2025-{m:02d}", f"회사{c}", 42.4 + (m + c) % 3 * 0.1] for m in range(1, 7) for c in range(10)]
        self.assertIn("sum_of_levels", self.doubts(rows))

    def test_something_called_a_level_is_averaged_and_not_asked_about(self):
        rows = [["연월", "지점", "수질 BOD"]] + [[f"2025-{m:02d}", f"지점{c}", 10 + m * c] for m in range(1, 7) for c in range(4)]
        self.assertNotIn("sum_of_levels", self.doubts(rows))
        path = self.book("수질.xlsx", 표=rows)
        note = next(item for item in suggest.plan_report(core.ExcelSource(path)).sections[0].sentences if item.code == "note")
        self.assertIn("평균 수질 BOD", note.result)
        self.assertIn("평균 수질 BOD 19.", note.result)  # 10, 16, 22 and 28 averaged, not one of them and not their sum

    def test_a_level_is_averaged_within_one_kind_of_measurement(self):
        rows = [["연월", "항목", "지점", "농도"]] + [[f"2025-{m:02d}", kind, place, size * m] for m in range(1, 7) for kind, size in (("BOD", 1), ("COD", 100)) for place in ("가", "나")]
        path = self.book("수질.xlsx", 표=rows)
        note = next(item for item in suggest.plan_report(core.ExcelSource(path)).sections[0].sentences if item.code == "note")
        self.assertIn("BOD", note.result)  # BOD and COD together would mean nothing

    def test_a_level_reads_as_a_range_not_as_shares_of_a_sum(self):
        rows = [["연월", "회사", "발열량"]] + [[f"2025-{m:02d}", f"회사{c}", 42 + c / 10 + m / 100] for m in range(1, 7) for c in range(10)]
        path = self.book("가스.xlsx", 표=rows)
        by_item = next(item for item in suggest.plan_report(core.ExcelSource(path)).sections[0].sentences if item.code == "by_item")
        self.assertIn("가장 높은 항목은 회사9", by_item.result)
        self.assertNotRegex(by_item.result, "합계|전체의")

    def test_an_average_over_time_that_adds_up_raises_no_doubt(self):
        rows = [["번호", "분류", "분기별 평균 수출금액(USD)"]] + [[n + 1, name, size] for n, (name, size) in enumerate((("전자", 100), ("전자", 200), ("화학", 50), ("화학", 70), ("전자", 300), ("화학", 80), ("예술", 9), ("예술", 11)))]
        self.assertNotIn("sum_of_levels", self.doubts(rows))

    def test_counting_rows_raises_a_doubt(self):
        rows = [["부서", "Age"]] + [[name, 30 + n] for name, size in (("영업", 9), ("개발", 6), ("인사", 4)) for n in range(size)]
        self.assertIn("count_only", self.doubts(rows))

    def test_reading_only_the_first_of_stacked_tables_raises_a_doubt(self):
        one = [["구분", "계"]] + [[name, size] for name, size in (("동구", 163), ("중구", 150), ("서구", 205), ("대덕", 119))] + [["합계", 637]]
        two = [["구분", "계"]] + [[name, size] for name, size in (("동구", 1), ("중구", 2), ("서구", 3), ("대덕", 4))] + [["합계", 10]]
        self.assertIn("first_table", self.doubts(one + [[None]] + two))

    def test_a_heading_below_the_first_figure_raises_a_doubt(self):
        rows = [["계", "임대"], [410, 385], [None, None], ["구분", "점포"], [None, "(수)"]] + [[f"{n}동", 30 + n * 7 % 11] for n in range(1, 11)]
        self.assertIn("heading_below", self.doubts(rows))

    def test_figures_with_no_date_of_their_own_raise_a_doubt(self):
        rows = [["수거일", "품목", "수량"]]
        for day in range(1, 8):
            rows.append([f"2025-07-{day:02d}", "장롱", 3 + day])
            rows += [[None, "소파", 2 + day], [None, "책상", 4 + day]]
        self.assertIn("dates_sparse", self.doubts(rows))

    def test_a_sheet_that_states_per_cent_is_read_in_points(self):
        rows = [["구분", 2019, 2020, 2021, 2022, 2023], ["Total children", 83, 84, 85, 85.8, 87], [], ["Unit:"], ["Per cent"]]
        self.assertEqual(self.doubts(rows), [])
        section = suggest.plan_report(core.ExcelSource(self.book("비율.xlsx", 표=rows))).sections[0]
        note = next(item for item in section.sentences if item.code == "note").result
        self.assertIn("87%", note)
        self.assertIn("85.8% 대비 +1.2포인트", note)
        self.assertIn("Unit: Per cent", section.heading)

    def test_a_percent_sign_somewhere_on_the_sheet_is_only_a_doubt(self):
        rows = [["구분", 2019, 2020, 2021, 2022, 2023], ["Total children", 83, 84, 85, 85.8, 87], [], ["Per cent"]]
        self.assertIn("maybe_percent", self.doubts(rows))
        note = next(item for item in suggest.plan_report(core.ExcelSource(self.book("비율2.xlsx", 표=rows))).sections[0].sentences if item.code == "note").result
        self.assertNotIn("%,", note.replace("% 증가", ""))

    def test_an_example_sheet_raises_a_doubt(self):
        self.assertIn("example_sheet", self.doubts([["지역", "매출"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(8)], sheet="Example"))

    def test_several_tables_in_one_column_raise_a_doubt(self):
        rows = [["Table 1: Gender", "Total"], ["Female", 121], ["Male", 134], [], ["Table 2: Age", "Total"], ["Under 20", 110], ["21 - 24", 90], ["25 - 29", 55]]
        self.assertIn("several_tables", self.doubts(rows))

    def test_a_total_that_does_not_add_up_raises_a_doubt(self):
        rows = [["지역", "값"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(6)] + [["합계", 5000]]
        self.assertIn("total_differs", self.doubts(rows))

    def test_doubts_from_several_files_say_which_file(self):
        good = self.book("좋음.xlsx", 표=[["지역", "매출"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(8)])
        bad = self.book("나쁨.xlsx", Example=[["지역", "매출"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(8)])
        plan = suggest.plan_files(core.Sources([good, bad]))
        self.assertEqual([(code, params.get("file")) for code, params in plan.doubts], [("example_sheet", "나쁨")])

    def test_every_doubt_has_words_in_both_languages(self):
        for code in suggest.DOUBTS:
            for lang in ("en", "ko"):
                self.assertTrue(i18n.text(lang, "doubt_" + code, figure="가"), code)

    def test_a_total_under_a_total_is_one_total(self):
        rows = [["구분", "합  계", None, None, "세대"], [None, "계", "남", "여", None]] + [[f"시{n}", 100 + n * 37 % 91, 50, 50, 10 + n * 3 % 7] for n in range(6)] + [["총  계", 700, 300, 300, 80]]
        source = core.ExcelSource(self.merged_book("합계계.xlsx", rows, ["B1:D1", "A1:A2", "E1:E2"]))
        names = [value.name for value in suggest.read_shape(source, "표").bases[0].values]
        self.assertEqual("".join(names[0].split()), "합계")

    def test_a_financial_year_ends_before_its_calendar_year_does(self):
        rows = [["Fin. Year", "Items"]] + [[f"{year}-{(year + 1) % 100:02d}", 1000 + year % 7 * 13] for year in range(2018, 2026)]
        path = self.book("회계.xlsx", 표=rows)
        for saved, expected in ((dt.datetime(2025, 10, 3), True), (dt.datetime(2026, 10, 7), False)):
            source = core.ExcelSource(path)
            source.wb.properties.modified = saved
            series = source._series(source._parse_ref("표!A,B", "x"))
            self.assertEqual(any(issue.code == "latest_unfinished" for issue in series.notes), expected, saved)

    def test_measures_of_different_units_under_one_name_are_not_added(self):
        rows = [["시도", "시군구", "Measures", "합계"]]
        for place, count, floors, area in (("괴산군", 25, 30, 900.5), ("진천군", 27, 36, 90597.03), ("청주시", 10, 15, 75028.82), ("음성군", 40, 44, 35270.1)):
            rows += [["충청북도", place, "건수", count], ["충청북도", place, "동수", floors], ["충청북도", place, "연면적(㎡)", area]]
        said = " ".join(item.result for item in suggest.plan_report(core.ExcelSource(self.book("건축.xlsx", 표=rows))).sections[0].sentences)
        self.assertIn("건수", said)
        self.assertIn("음성군", said)
        self.assertNotIn("90,660", said)
        self.assertNotIn("90660", said)

    def test_measures_under_each_date_are_not_added_either(self):
        rows = [["월", "지표", "값"]]
        for month in range(1, 8):
            rows += [[f"2025-{month:02d}", "건수", 10 + month], [f"2025-{month:02d}", "연면적", 5000 + month * 7]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertIn("B=건수", found["latest"].key)
        self.assertEqual(found["latest"].result, "17")

    def test_a_column_that_goes_with_each_name_is_not_a_kind_to_choose(self):
        rows = [["대륙", "국가", "구분", "값"]]
        for continent, country, size in (("아시아", "일본", 30), ("아시아", "중국", 50), ("유럽", "독일", 20), ("유럽", "영국", 10)):
            rows += [[continent, country, "월계", size], [continent, country, "누계", size * 10]]
        said = " ".join(item.result for item in suggest.plan_report(core.ExcelSource(self.book("대륙2.xlsx", 표=rows))).sections[0].sentences)
        self.assertIn("중국", said)
        self.assertIn("독일", said)

    def test_one_kind_is_enough_and_the_other_can_add_up(self):
        rows = [["지역", "구분", "성별", "값"]]
        for place, size in (("가", 10), ("나", 20), ("다", 30)):
            for kind, scale in (("월계", 1), ("누계", 10)):
                for sex, part in (("남", 1), ("여", 2)):
                    rows.append([place, kind, sex, size * scale * part])
        said = " ".join(item.result for item in suggest.plan_report(core.ExcelSource(self.book("두갈래.xlsx", 표=rows))).sections[0].sentences)
        self.assertIn("다(90)", said)

    def test_a_kind_only_some_names_have_still_splits_them(self):
        rows = [["공항", "노선", "값"]] + [[name, kind, size * scale] for name, size in (("김포", 10), ("김해", 8)) for kind, scale in (("국내선", 1), ("국제선", 3), ("화물", 5), ("전세", 7))] + [[name, "국내선", 3 + n] for n, name in enumerate(("광주", "대구", "여수", "울산", "제주", "청주"))]
        source = core.ExcelSource(self.book("공항.xlsx", 표=rows))
        shape = suggest.read_shape(source, "표")
        base = next(b for b in shape.bases if b.spec == "A")
        value = suggest.default_value(base)
        self.assertEqual(suggest.default_conditions(source, shape, base, value, "sum"), [("B", "국내선")])

    def test_a_pivot_tables_indented_lines_are_parts(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "표"
        lines = [("Term 3", 0, 30), ("North", 1, 20), ("Abbey", 2, 8), ("Bourke", 2, 12), ("South", 1, 10), ("Coast", 2, 10), ("Term 4", 0, 70), ("North", 1, 25), ("Abbey", 2, 25), ("South", 1, 45), ("Coast", 2, 15), ("Dale", 2, 30), ("Grand Total", 0, 100)]
        ws.append(["Count", "Column Labels"])
        ws.append(["Row Labels", "Violence", "Grand Total"])
        for name, indent, size in lines:
            ws.append([name, size // 2, size])
            ws.cell(ws.max_row, 1).alignment = openpyxl.styles.Alignment(indent=indent)
        path = os.path.join(self.dir, "피벗.xlsx")
        wb.save(path)
        source = core.ExcelSource(path)
        found = source._comparison(source._parse_ref("표!A,C", "x"))
        self.assertEqual([item.label for item in found.items], ["Term 3", "Term 4"])
        self.assertEqual(found.total, 100)
        self.assertEqual(found.header, "Grand Total")

    def test_headings_with_no_figure_are_not_an_outer_level(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "표"
        ws.append(["항목", "2026 $m"])
        for name, indent, size in (("Expenses by function", 0, None), ("Health", 1, 45), ("Defence", 1, 25), ("Revenue by source", 0, None), ("Taxation", 1, 19), ("Fees", 1, 11), ("Total", 0, 100)):
            ws.append([name, size])
            ws.cell(ws.max_row, 1).alignment = openpyxl.styles.Alignment(indent=indent)
        path = os.path.join(self.dir, "소제목.xlsx")
        wb.save(path)
        source = core.ExcelSource(path)
        self.assertEqual(len(source._comparison(source._parse_ref("표!A,B", "x")).items), 4)

    def test_one_line_set_in_is_not_a_pivot_of_one(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "표"
        ws.append(["이름", "값"])
        for name, indent, size in (("가", 1, 10), ("나", 1, 20), ("다", 1, 30), ("계", 0, 60)):
            ws.append([name, size])
            ws.cell(ws.max_row, 1).alignment = openpyxl.styles.Alignment(indent=indent)
        path = os.path.join(self.dir, "한칸.xlsx")
        wb.save(path)
        source = core.ExcelSource(path)
        self.assertEqual(len(source._comparison(source._parse_ref("표!A,B", "x")).items), 3)

    def test_pivot_caches_are_not_read(self):
        caches = openpyxl.reader.workbook.WorkbookParser.pivot_caches.fget(None)
        self.assertIsNone(caches[7])

    def test_more_identifier_names(self):
        for name in ("CRED_LIC_PCODE", "CRED_LIC_NUM", "CRED_LIC_AFSL_NUM", "respondent", "Respondent ID", "STORE_NBR"):
            self.assertFalse(suggest._names_a_figure(name), name)
        for name in ("No of Units", "No. of staff", "Number", "Numerous"):
            self.assertTrue(suggest._names_a_figure(name), name)

    def test_tables_stacked_in_one_sheet_read_the_first(self):
        rows = [["경로당 수 및 회원현황"], ["구분", "계", "회원"]] + [[name, size, people] for name, size, people in (("동구지회", 163, 5804), ("중구지회", 150, 5200), ("서구지회", 205, 7100), ("대덕구지회", 119, 5402))] + [["합계", 637, 23506]]
        rows += [[None], ["구분", "계", "회원"]] + [[name, size, people] for name, size, people in (("동구지회", 1, 30), ("중구지회", 2, 40), ("서구지회", 3, 50), ("대덕구지회", 4, 60))] + [["합계", 10, 180]]
        plan = suggest.plan_report(core.ExcelSource(self.book("쌓인표.xlsx", 표=rows)))
        said = " ".join(item.result for item in plan.sections[0].sentences)
        self.assertIn("서구지회(205)", said)
        self.assertIn("637", said)
        self.assertNotIn("구분", said)
        self.assertNotIn("건수", said)

    def test_rows_nested_in_a_long_list_keep_only_the_top_level(self):
        items, rows = [], []
        for term, regions in (("Term 3", 60), ("Term 4", 60)):
            lines = []
            for region in range(regions):
                schools = [3 + (region + school) % 5 for school in range(2)]
                lines += [(f"{term} region {region}", sum(schools))] + [(f"{term} school {region}-{school}", size) for school, size in enumerate(schools)]
            rows += [(term, sum(size for name, size in lines if "region" in name))] + lines
        found = facts.without_parts([facts.Item(name, size) for name, size in rows])
        self.assertEqual([item.label for item in found], ["Term 3", "Term 4"])

    def test_small_numbers_that_happen_to_add_up_are_not_parts(self):
        items = [facts.Item(*pair) for pair in (("가병원", 21), ("나병원", 2), ("다의원", 1), ("라의원", 1), ("마의원", 3), ("바의원", 2), ("사의원", 1))]
        self.assertIsNone(facts.without_parts(items))

    def test_a_chance_match_in_a_long_list_still_removes_nothing(self):
        long_list = [facts.Item(f"회사{n}", 300 + n * 37 % 211) for n in range(150)]
        long_list[10:13] = [facts.Item("큰회사", 300), facts.Item("작은회사", 100), facts.Item("다른회사", 200)]
        self.assertIsNone(facts.without_parts(long_list))

    def test_an_average_over_time_adds_up_across_things(self):
        rows = [["번호", "분류", "분기별 평균 수출금액(USD)", "분기평균성장률(%)"]] + [[n + 1, name, size, 1.5] for n, (name, size) in enumerate((("전자", 100), ("전자", 200), ("화학", 50), ("화학", 70), ("전자", 300), ("화학", 80), ("예술", 9), ("예술", 11)))]
        source = core.ExcelSource(self.book("수출.xlsx", 표=rows))
        said = " ".join(item.result for item in suggest.plan_report(source).sections[0].sentences)
        self.assertIn("전자", said)
        self.assertIn("600", said)
        for name, expected in (("월평균 판매량", "sum"), ("평균 단가", "mean"), ("평균 점수", "mean"), ("연평균 수입액", "sum")):
            base = suggest.Base("name", "A", "이름", many_rows=True)
            self.assertEqual(suggest.default_how(base, suggest.Value("C", name)), expected, name)

    def test_a_sheet_wider_than_its_cells_makes_no_new_cells(self):
        path = self.book("넓은.xlsx", 표=[["구분", "값"]] + [[f"항목{n}", 10 + n * 37 % 91] for n in range(8)])
        wb = openpyxl.load_workbook(path)
        wb["표"].cell(2, 16000).value = None
        wb["표"].cell(3, 16000).value = "끝"
        wb.save(path)
        source = core.ExcelSource(path)
        ws = source._sheet("표")
        before = len(ws._cells)
        self.assertFalse(source._has_formula(ws, 2, 2))
        self.assertLess(len(ws._cells), before + 100)

    def test_names_that_are_not_figures(self):
        for name in ("ISBN", "ISSN", "SKU", "바코드", "EmployeeNumber", "사번", "학번", "발생차수", "회차", "순위", "등수", "Rank", "등급", "Group End Date", "Start_Date", "일자", "지정일", "생년월일", "기준일", "발행년도", "Year"):
            self.assertFalse(suggest._names_a_figure(name), name)
        for name in ("차수별 건수", "등급별 인원", "일자별 매출", "Update count", "유효 측정일수", "Datasets", "매출액", "연간 수입", "Sales rank score", "2023년도 수량", "2025년도", "Year 2024 sales"):
            self.assertTrue(suggest._names_a_figure(name), name)

    def test_a_book_list_counts_books_and_does_not_add_their_numbers(self):
        rows = [["No", "서명", "발행년", "ISBN"]] + [[n + 1, f"책{n}", 2020 + n % 6, 9791190000000 + n * 7919] for n in range(60)]
        plan = suggest.plan_report(core.ExcelSource(self.book("책.xlsx", 표=rows)))
        self.assertNotIn("ISBN", plan.what)
        self.assertNotRegex(" ".join(item.result for item in plan.sections[0].sentences), r"\d{9,}|\d{1,3}(?:,\d{3}){3,}")

    def test_long_numbers_under_each_date_are_identifiers_whatever_they_are_called(self):
        rows = [["연도", "등록", "금액"]] + [[2020 + n % 6, 9791190000000 + n * 7919, 1000 + n * 37 % 211] for n in range(60)]
        plan = suggest.plan_report(core.ExcelSource(self.book("긴값.xlsx", 표=rows)))
        self.assertIn("금액", plan.what)
        self.assertNotIn("등록", plan.what)

    def test_a_running_number_that_starts_again_is_still_a_running_number(self):
        numbers = list(range(1, 41)) + list(range(1, 21))
        self.assertTrue(suggest._steps_by_one(numbers))
        self.assertFalse(suggest._steps_by_one([3, 4, 9, 2, 8, 1, 7, 5, 6, 10]))
        rows = [["분류", "지역", "번호표", "규모"]] + [["가" if n < 40 else "나", f"지역{n}", number, 100 + n * 37 % 91] for n, number in enumerate(numbers)]
        plan = suggest.plan_report(core.ExcelSource(self.book("차수.xlsx", 표=rows)))
        self.assertIn("규모", plan.what)
        self.assertNotIn("번호표", plan.what)

    def test_ages_are_not_added_even_beside_a_column_of_kinds(self):
        rows = [["부서", "구분", "Age"]] + [[name, "재직" if n % 2 else "퇴직", 30 + n * 7 % 11] for name, size in (("영업", 9), ("개발", 6), ("인사", 4)) for n in range(size)]
        source = core.ExcelSource(self.book("나이.xlsx", 표=rows))
        shape, base, _found = suggest.first_choice(source)
        self.assertEqual(suggest.default_how(base, suggest.default_value(base)), "count")
        self.assertEqual(suggest.plan_report(source).what, "표")

    def test_a_list_of_unlike_things_has_no_figure(self):
        rows = [["Name", "Value"], ["Semi major axis", 6378137], ["Inverse flattening", 298.257222101], ["Flattening", 0.00335281], ["Semi-minor axis", 6356752.31], ["Eccentricity", 0.00669438], ["e", 0.08181919], ["False northing", 10000000], ["Scale", 0.9996], ["False easting", 500000]]
        self.assertIsNone(suggest.plan_report(core.ExcelSource(self.book("상수.xlsx", Engine=rows))))

    def test_figures_of_very_different_sizes_are_still_figures(self):
        rows = [["분야", "예산"]] + [[f"분야{n}", size] for n, size in enumerate((1280000000000, 724000000000, 9662766000, 5000, 0, 320000000000))]
        self.assertEqual(suggest.plan_report(core.ExcelSource(self.book("예산.xlsx", 표=rows))).what, "표 · 예산")

    def test_a_year_beside_named_things_is_not_a_figure(self):
        rows = [["분야", "회계연도", "예산"]] + [[f"분야{n}", 2025, 1000 + n * 37 % 91] for n in range(8)]
        self.assertEqual(suggest.plan_report(core.ExcelSource(self.book("연도열.xlsx", 표=rows))).what, "표 · 예산")

    def test_a_subtotal_row_beside_its_parts_is_not_added_to_them(self):
        rows = [["Year", "Group", "Item", "Items"]]
        for year, parts in ((2022, (40, 60)), (2023, (50, 70)), (2024, (55, 85)), (2025, (30, 50))):
            rows += [[year, "FY Subtotal", None, sum(parts)], [year, "Auslan", "Interpreting", parts[0]], [year, "Support", "Mentoring", parts[1]]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertIn("B=FY Subtotal", found["latest"].key)
        self.assertEqual(found["latest"].result, "80")

    def test_marked_whole_rows_under_names_are_chosen_and_nothing_is_added_twice(self):
        rows = [["항만명", "품목", "구분", "물량"]]
        for port, size in (("평택", 341), ("광양", 216), ("인천", 90), ("동해", 45)):
            for kind, share in (("4:환적", 1), ("2:수출", 3), ("5합계", 4)):
                rows += [[port, "00:소계", kind, size * share], [port, "01:육류", kind, size * share // 4], [port, "02:곡물", kind, size * share - size * share // 4]]
        source = core.ExcelSource(self.book("항만.xlsx", 표=rows))
        plan = suggest.plan_report(source)
        said = " ".join(item.result for item in plan.sections[0].sentences)
        self.assertIn("평택", said)
        self.assertIn(f"{341 * 4:,}", said)
        self.assertNotIn(f"{341 * 8:,}", said)
        self.assertIn("5합계", said)
        self.assertIn("00:소계", said)

    def test_a_row_that_equals_the_rest_is_a_whole_whatever_it_is_called(self):
        rows = [["연도", "성별", "인원"]]
        for year, (men, women, other) in zip(range(2020, 2026), ((60, 40, 5), (61, 43, 4), (70, 50, 6), (72, 55, 3), (66, 52, 2), (80, 60, 7))):
            rows += [[year, "남", men], [year, "여", women], [year, "미상", other], [year, "모두", men + women + other]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertIn("B=모두", found["latest"].key)
        self.assertEqual(found["latest"].result, "147")

    def test_two_parts_of_much_the_same_size_are_not_a_whole(self):
        rows = [["연도", "성별", "인원"]]
        for year, size in zip(range(2020, 2026), (100, 110, 120, 130, 140, 150)):
            rows += [[year, "남", size], [year, "여", size + 1]]
        _s, _shape, _base, found = self.first(표=rows)
        self.assertNotIn("B=", found["latest"].key)
        self.assertEqual(found["latest"].result, "301")

    def test_a_word_that_only_ends_like_a_total_is_a_part(self):
        for word in ("가계", "회계", "subtotals of sales", "통계"):
            self.assertIsNone(suggest.ALL_ROWS_WORD.fullmatch(word), word)
        for word in ("5합계", "00:소계", "소계", "FY Subtotal", "Sub-total", "1. 계", "All Regions", "Grand Total", "전체"):
            self.assertIsNotNone(suggest.ALL_ROWS_WORD.fullmatch(word), word)

    def test_kinds_without_a_whole_still_take_the_first_under_names(self):
        rows = [["국가", "구분", "품목", "값"]]
        for name, size in (("일본", 10), ("중국", 30), ("대만", 20)):
            rows += [[name, "월계", "소계", size], [name, "월계", "가", size - 1], [name, "월계", "나", 1], [name, "누계", "소계", size * 10], [name, "누계", "가", size * 10 - 1], [name, "누계", "나", 1]]
        source = core.ExcelSource(self.book("두구분.xlsx", 표=rows))
        said = " ".join(item.result for item in suggest.plan_report(source).sections[0].sentences)
        self.assertIn("중국", said)
        self.assertIn("30", said)
        self.assertNotIn("330", said)

    def test_a_total_merged_across_levels_is_not_a_row_to_choose(self):
        rows = [["대륙", "국가", "구분", "2025년"], ["총계", None, "월계", 600], [None, None, "누계", 6000]]
        for continent, country, size in (("아시아", "소계", 500), (None, "일본", 300), (None, "중국", 200), ("서구", "소계", 100), (None, "미국", 70), (None, "영국", 30)):
            rows += [[continent, country, "월계", size], [None, None, "누계", size * 10]]
        merged = ["A2:B3", "A4:A9", "A10:A15"] + [f"B{row}:B{row + 1}" for row in range(4, 16, 2)]
        said = " ".join(item.result for item in suggest.plan_report(core.ExcelSource(self.merged_book("대륙국가.xlsx", rows, merged))).sections[0].sentences)
        self.assertIn("일본", said)
        self.assertIn("전체 600", said)
        self.assertNotIn("누계 ", said)

    def test_a_kind_of_figure_is_not_what_the_report_compares(self):
        rows = [["지역", "구분", "2025년"]]
        for name, size in (("아시아", 100), ("유럽", 20), ("미주", 15), ("중동", 9), ("아프리카", 7), ("대양주", 4), ("기타", 2), ("총계", 157)):
            rows += [[name, "월계", size], [name, "누계", size * 12], [name, "전년", size * 11]]
        said = " ".join(item.result for item in suggest.plan_report(core.ExcelSource(self.book("대륙.xlsx", 표=rows))).sections[0].sentences)
        self.assertIn("아시아", said)
        self.assertNotIn("누계 ", said)

    def test_a_number_that_always_goes_with_one_name_is_a_code(self):
        species = [("가리비", 23270007), ("상어", 37017001), ("참치", 37441004), ("도미", 37353001), ("민어", 37354001), ("갈치", 37441002), ("고등어", 37441003), ("홍어", 37031001)]
        rows = [["연도", "어종", "어종분류", "어획량"]]
        for year in range(2019, 2025):
            for number, (fish, code) in enumerate(species):
                rows.append([year, fish, code, 1000 + year % 7 * 100 + number * 37])
        plan = suggest.plan_report(core.ExcelSource(self.book("어획.xlsx", 어획=rows)))
        self.assertEqual(plan.what, "어획 · 어획량")

    def test_an_area_that_repeats_with_its_name_is_not_a_figure_but_its_parts_add_up(self):
        rows = [["지역", "지역면적", "등급코드", "등급별 면적"]]
        for name, parts in (("가", (10.5, 20.5)), ("나", (5.25, 5.25, 30.0)), ("다", (1.5, 2.5)), ("라", (7.0, 8.0)), ("마", (3.5, 4.75))):
            rows += [[name, sum(parts), grade, part] for grade, part in enumerate(parts)]
        source = core.ExcelSource(self.book("면적.xlsx", 면적=rows))
        plan = suggest.plan_report(source)
        self.assertIn("등급별 면적", plan.what)
        self.assertNotIn("지역면적", plan.what)
        self.assertIn("나(40.5)", " ".join(item.result for item in plan.sections[0].sentences))

    def test_codes_named_as_such_are_not_figures(self):
        for heading in ("ABN", "FILE NUMBER", "CUSTOMER_ID", "Postcode", "우편번호", "사업자등록번호", "위도", "Longitude"):
            rows = [["이름", heading, "직원수"]] + [[f"회사{n}", 880000 + n * 137, 50 + n * 3 % 17] for n in range(24)]
            plan = suggest.plan_report(core.ExcelSource(self.book("번호.xlsx", 표=rows)))
            self.assertEqual(plan.what, "표 · 직원수", heading)

    def test_a_count_of_something_is_still_a_figure(self):
        rows = [["이름", "Number of individuals"]] + [[f"회사{n}", 880000 + n * 137] for n in range(24)]
        plan = suggest.plan_report(core.ExcelSource(self.book("인원.xlsx", 표=rows)))
        self.assertEqual(plan.what, "표 · Number of individuals")

    def test_a_count_under_a_name_that_repeats_is_added_up(self):
        rows = [["State", "Postcode", "Number of individuals"]] + [[state, 3000 + number * 10 + n, 100 + n - number * 10] for number, state in enumerate(("VIC", "NSW", "QLD")) for n in range(6)]
        source = core.ExcelSource(self.book("주.xlsx", 표=rows))
        plan = suggest.plan_report(source)
        self.assertIn("Number of individuals", plan.what)
        self.assertIn("615", " ".join(item.result for item in plan.sections[0].sentences))

    def test_long_numbers_each_different_are_identifiers_whatever_they_are_called(self):
        rows = [["이름", "등록", "직원수"]] + [[f"회사{n}", 11000073870 + n * 1000003, 50 + n * 3 % 17] for n in range(24)]
        plan = suggest.plan_report(core.ExcelSource(self.book("긴번호.xlsx", 표=rows)))
        self.assertEqual(plan.what, "표 · 직원수")

    def test_figures_of_one_size_are_still_figures(self):
        rows = [["업종", "월급여"]] + [[f"업종{n}", 3000000 + n * 12347] for n in range(24)]
        plan = suggest.plan_report(core.ExcelSource(self.book("급여.xlsx", 표=rows)))
        self.assertEqual(plan.what, "표 · 월급여")

    def test_a_series_that_stays_in_one_size_is_still_a_figure(self):
        rows = [["연도", "구분", "금액"]] + [[2000 + n, f"구분{n % 3}", 1234567890 + n * 991] for n in range(24)]
        plan = suggest.plan_report(core.ExcelSource(self.book("큰값.xlsx", 표=rows)))
        self.assertIn("금액", plan.what)

    def test_a_heading_that_is_only_punctuation_names_no_figure(self):
        rows = [["항목", "-", "건수"]] + [[f"항목{n}", n % 3 + 1, 10 + n * 7] for n in range(12)]
        plan = suggest.plan_report(core.ExcelSource(self.book("빈머리.xlsx", 표=rows)))
        self.assertEqual(plan.what, "표 · 건수")

    def test_a_sheet_that_says_it_is_a_table_of_contents_is_not_data(self):
        rows = [["Taxation statistics: Companies detailed table index"], ["Label", "Item"]] + [[f"항목{n}", n % 9 + 1] for n in range(12)]
        self.assertIsNone(suggest.plan_report(core.ExcelSource(self.book("색인.xlsx", Companies=rows))))

    def test_a_price_index_is_data(self):
        rows = [["Consumer price index"], ["Month", "Index"]] + [[dt.datetime(2024, m, 1), 100 + m] for m in range(1, 13)]
        self.assertIsNotNone(suggest.plan_report(core.ExcelSource(self.book("물가.xlsx", 물가=rows))))

    def test_measures_under_merged_headings_say_which_heading(self):
        rows = [["구 분", "일반수출신고", None, "간이수출신고", None], [None, "건수", "금액", "건수", "금액"]] + [[f"{year}년", 100 + year % 9, 50, 3000 + year % 7, 700] for year in range(2019, 2026)]
        source = core.ExcelSource(self.merged_book("머리.xlsx", rows, ["B1:C1", "D1:E1", "A1:A2"]))
        plan = suggest.plan_report(source)
        self.assertIn("일반수출신고 건수", plan.what)
        self.assertIn("간이수출신고 건수", plan.what)

    def test_a_heading_only_merged_over_others_is_not_borrowed_from_a_title(self):
        rows = [["전체 현황 보고"], ["지역", "건수", "건수"]] + [[f"지역{n}", 10 + n * 3 % 7, 20 + n * 5 % 7] for n in range(6)]
        source = core.ExcelSource(self.merged_book("제목.xlsx", rows, []))
        names = [value.name for base in suggest.read_shape(source, "표").bases for value in base.values]
        self.assertTrue(names)
        self.assertFalse(any("현황" in name for name in names), names)

    def test_a_label_for_each_level_is_not_read_as_one_name(self):
        rows = [["지역", "시설", "유형", "구분", "입장객"]]
        for region in ("동구", "서구"):
            for number in range(5):
                rows.append([region if number == 0 else None, f"{region}시설{number}", "무료", "내국인", 100 + number])
            rows.append([None, None, None, None, 510])
        rows += [["유료", None, None, None, 0], ["전 체", None, None, None, 1020]]
        plan = suggest.plan_report(core.ExcelSource(self.merged_book("계층.xlsx", rows, ["A2:A7", "A8:A13", "A14:D14", "A15:D15"])))
        said = " ".join(item.result for item in plan.sections[0].sentences)
        self.assertIn("시설4", said)
        self.assertNotIn("내국인 ", said)
        self.assertNotIn("유료", said)

    def test_a_company_called_total_is_not_a_total(self):
        rows = [["이름", "직원수"]] + [[name, size] for name, size in (("Total Tools Importing Pty Ltd", 123), ("TotalEnergies", 300), ("Alpha", 500), ("Beta", 77), ("Gamma", 210))]
        source = core.ExcelSource(self.book("회사.xlsx", 표=rows))
        found = source._comparison(source._parse_ref("표!A,B", "x"))
        self.assertEqual(len(found.items), 5)
        rows += [["Total staff", 1210]]
        source = core.ExcelSource(self.book("회사2.xlsx", 표=rows))
        found = source._comparison(source._parse_ref("표!A,B", "x"))
        self.assertEqual((len(found.items), found.total), (5, 1210))

    def test_a_stated_total_merged_in_from_the_left_is_still_the_total(self):
        rows = [["지역", "시설", "입장객"]] + [[None, f"시설{n}", 100 + n * 3 % 7] for n in range(6)] + [["전 체", None, sum(100 + n * 3 % 7 for n in range(6))]]
        plan = suggest.plan_report(core.ExcelSource(self.merged_book("총계병합.xlsx", rows, ["A8:B8"])))
        self.assertIn("전체", " ".join(item.result for item in plan.sections[0].sentences))
        self.assertNotIn("항목 합계", " ".join(item.result for item in plan.sections[0].sentences))

    def test_main_figures_merged_in_at_the_top_stay(self):
        rows = [["구분", None, "수량"], ["총생산량", None, 900], ["일평균", None, 300]] + [["유효", f"세부{n}", 100 + n * 3 % 7] for n in range(4)]
        source = core.ExcelSource(self.merged_book("위쪽병합.xlsx", rows, ["A1:B1", "A2:B2", "A3:B3"]))
        self.assertEqual(core._rows_of_an_outer_level(source._sheet("표"), source._parse_ref("표!B,C", "x"), [2]), frozenset())

    def test_labels_merged_in_on_every_row_are_still_the_names(self):
        rows = [["구분", None, "값"]] + [[f"항목{n}", None, 10 + n * 3 % 7] for n in range(6)]
        merged = [f"A{row}:B{row}" for row in range(1, 8)]
        plan = suggest.plan_report(core.ExcelSource(self.merged_book("가로병합.xlsx", rows, merged)))
        self.assertIn("항목", " ".join(item.result for item in plan.sections[0].sentences))

    def test_a_place_named_like_a_coordinate_is_not_a_coordinate(self):
        for name in ("위도항", "경도항", "남경도"):
            self.assertIsNone(suggest.ID_NAME.search(name), name)
        for name in ("위도", "경도", "시설 위도", "Latitude", "SITE_ID"):
            self.assertIsNotNone(suggest.ID_NAME.search(name), name)
        for name in ("Number of staff", "Total number", "Number", "Average number"):
            self.assertIsNone(suggest.ID_NAME.search(name), name)

    def test_a_part_under_the_heading_all_is_not_the_total(self):
        rows = [["회사", "전체", None, None, "개발", None], [None, "임원", "직원", "소계", "임원", "직원"]] + [[f"회사{n}", 5 + n, 100 + n * 7, 105 + n * 8, 2 + n, 50 + n * 3] for n in range(6)]
        plan = suggest.plan_report(core.ExcelSource(self.merged_book("전체.xlsx", rows, ["B1:D1", "E1:F1", "A1:A2"])))
        self.assertIn("전체 소계", plan.what)
        self.assertNotIn("전체 임원", plan.what)

    def test_a_title_merged_across_the_sheet_is_not_a_total_heading(self):
        rows = [["목록"], ["시설", "1월", "2월", "합계"]] + [[f"시설{n}", 10 + n * 3 % 7, 20 + n * 5 % 7, 30 + n * 4 % 7] for n in range(6)]
        plan = suggest.plan_report(core.ExcelSource(self.merged_book("머리제목.xlsx", rows, ["A1:D1"])))
        self.assertEqual(plan.what, "표 · 합계")

    def test_a_code_letter_over_a_name_is_left_off(self):
        rows = [["구분", "C", None, "D", None], ["월", "제조업", "제조업", "서비스", "서비스"]] + [[f"2025-{m:02d}", 100 + m * 3 % 7, 3, 200 + m * 5 % 7, 4] for m in range(1, 8)]
        source = core.ExcelSource(self.merged_book("글자.xlsx", rows, ["B1:C1", "D1:E1"]))
        names = [value.name for base in suggest.read_shape(source, "표").bases for value in base.values]
        self.assertIn("제조업", names)

    def test_a_sheet_that_breaks_the_first_down_is_left_out(self):
        whole = [["연도", "환자수"]] + [[2020 + n, 100 + n] for n in range(5)]
        parts = [["연도", "성별", "환자수"]] + [[2020 + n, sex, 50 + n] for n in range(5) for sex in ("남", "여")]
        plan = suggest.plan_report(core.ExcelSource(self.book("쪼갠.xlsx", 전체=whole, 성별=parts)))
        self.assertEqual(len(plan.sections), 1)
        self.assertEqual(plan.sheets_left, 1)

    def test_the_days_of_the_week_are_not_items(self):
        rows = [["일자", "요일구분", "노선", "횟수"]]
        for day in range(1, 29):
            for line in ("경부선", "호남선"):
                rows.append([dt.datetime(2025, 1, day), "월화수목금토일"[day % 7], line, day])
        _shape, _base, found = suggest.first_choice(core.ExcelSource(self.book("요일.xlsx", 운행=rows)))
        by_item = next(item for item in found if item.code == "by_item")
        self.assertIn("경부선", by_item.result)

    def test_columns_of_years_are_one_figure_at_different_times(self):
        rows = [["구분", "2025년(원)", "2024년(원)", "2023년(원)"], ["가", 30, 20, 10], ["나", 60, 50, 40], ["다", 90, 80, 70]]
        plan = suggest.plan_report(core.ExcelSource(self.book("연도열.xlsx", 표=rows)))
        self.assertEqual(plan.what, "표 · 2025년(원)")
        self.assertEqual(plan.figures_left, 2)

    def test_lists_with_repeating_numbers_are_listed_and_not_ranked(self):
        rows = [["과목", "2025년"], ["I. 유동자산", 900], ["1. 현금", 400], ["2. 예금", 500], ["II. 비유동자산", 500], ["I. 유동부채", 300], ["1. 미지급금", 300], ["II. 비유동부채", 700]]
        excel = self.book("재무.xlsx", 재무=rows)
        source = core.ExcelSource(excel)
        plan = suggest.plan_report(source)
        words = [item.result for item in plan.sections[0].sentences]
        self.assertEqual(words, ["2025년: I. 유동자산 900, II. 비유동자산 500, I. 유동부채 300, II. 비유동부채 700."])
        self.assertIsNone(plan.sections[0].chart)
        self.assertEqual(source.text_value("읽는법:재무!A,B")[0], facts.SEPARATE_LISTS["ko"])

    def test_a_title_alone_at_the_top_says_what_was_counted(self):
        rows = [["분기별 승객 통계 (시드니 노선)"], [], ["구간"] + [dt.datetime(2020 + q // 4, q % 4 * 3 + 3, 1) for q in range(20)]] + [[f"노선{n}"] + [100 + q * 7 % 11 for q in range(20)] for n in range(30)]
        source = core.ExcelSource(self.book("제목.xlsx", 표=rows))
        plan = suggest.plan_report(source)
        request = next(iter(suggest.ask_insights(source, plan.sections, "제목").values()))[0]
        self.assertIn("분기별 승객 통계", request.facts[1].text)
        self.assertIn("전체의 합계", request.facts[0].text)

    def test_a_chosen_sentence_replaces_the_reading_only_for_the_same_figures(self):
        excel = self.book("통찰.xlsx", 매출=self.months(("내국인", 100), ("외국인", 500)))
        source = core.ExcelSource(excel)
        plan = suggest.plan_report(source)
        asks = suggest.ask_insights(source, plan.sections, "통찰")
        self.assertEqual({figure for _request, figure in asks.values()}, {"내국인", "외국인"})
        for request, _figure in asks.values():
            self.assertEqual(request.facts[0].id, "subject")
            self.assertIn(request.facts[0].figure, ("내국인", "외국인"))
            self.assertIn("매출", request.facts[1].text)  # 외국인 alone would read as a number of foreigners
        request = next(request for request, figure in asks.values() if figure == "외국인")
        fresh = {request.key.casefold(): {"fingerprint": request.fingerprint, "text": "고른 문장."}}
        path = os.path.join(self.dir, "통찰.pptx")
        out = os.path.join(self.dir, "통찰_out.pptx")
        suggest.make_report(path, "통찰", suggest.choose_insights(plan.sections, asks, fresh))
        report = core.run(path, excel, out, insights=fresh)
        self.assertEqual(report.count("error"), 0, [(i.key, i.status) for i in report.items])
        self.assertEqual(next(i for i in report.items if i.key == request.key).status, "ok")
        words = " ".join(shape.text_frame.text for shape in Presentation(out).slides[0].shapes if shape.has_text_frame)
        self.assertIn("고른 문장.", words)
        self.assertIn("외국인 519", words)  # its main sentence stays
        stale = {request.key.casefold(): {"fingerprint": "다른 수치", "text": "고른 문장."}}
        kept = suggest.choose_insights(plan.sections, asks, stale)
        self.assertEqual([item.code for item in kept[0].sentences], [item.code for item in plan.sections[0].sentences])

    def test_large_plain_numbers_get_commas_only_when_the_report_was_laid_out_here(self):
        excel = self.book("쉼표.xlsx", 표=self.months(("매출", 493434)))
        _shape, _base, found = suggest.first_choice(core.ExcelSource(excel))
        path = os.path.join(self.dir, "쉼표.pptx")
        suggest.make_starter(path, "쉼표", found)
        cell = lambda name: next(s for s in Presentation(name).slides[0].shapes if s.has_table).table.cell(1, 1).text
        core.run(path, excel, os.path.join(self.dir, "그대로.pptx"))
        core.run(path, excel, os.path.join(self.dir, "쉼표들어감.pptx"), thousands=True)
        self.assertEqual(cell(os.path.join(self.dir, "그대로.pptx")), "493448")
        self.assertEqual(cell(os.path.join(self.dir, "쉼표들어감.pptx")), "493,448")

    def test_the_title_comes_from_the_name_unless_it_says_nothing(self):
        section = suggest.Section("매출(단위 원)", [])
        self.assertEqual(suggest.title_of(r"C:\자료\면세점_매출_액.xlsx", [section]), "면세점 매출 액")
        self.assertEqual(suggest.title_of(r"C:\자료\data.xlsx", [section]), "매출(단위 원)")
        self.assertEqual(suggest.title_of(r"C:\자료\Sheet1.xlsx", [suggest.Section("", [])]), "Sheet1")

    def test_english_words_follow_the_language(self):
        source = core.ExcelSource(self.book("en.xlsx", Daily=[["Date", "Sales"]] + [[dt.datetime(2026, 1, d), d] for d in range(1, 15)]))
        _shape, _base, found = suggest.first_choice(source, lang="en")
        self.assertEqual({item.code: item.key for item in found}["latest"], "latest:Daily!A,B")
        self.assertTrue(any(item.code == "week" and item.key == "week:Daily!A,B" for item in found))


class OtherFileKinds(Base):
    """samples/formats holds one workbook saved by real Excel in each kind (tools/make_format_samples.ps1)."""
    FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "formats")

    def test_every_kind_reads_the_same_figures(self):
        sentences = {}
        for ext in (".xlsx", ".xls", ".xlsb", ".ods"):
            source = core.ExcelSource(os.path.join(self.FOLDER, "monthly" + ext))
            self.assertEqual([ws.title for ws in source.wb.worksheets], ["월별", "지점"], ext)
            sentences[ext] = source.text_value("해설:월별!A,B")[0]
        self.assertEqual(len(set(sentences.values())), 1, sentences)
        self.assertIn("2026년 6월 매출 1,829", sentences[".xls"])

    def test_an_xls_keeps_its_number_formats_and_merged_cells(self):
        source = core.ExcelSource(os.path.join(self.FOLDER, "monthly.xls"))
        rows = source.range_cells("최신1:월별!A,B,C")[0]
        self.assertEqual([cell[1] for cell in rows[-1]][1:], ["1,829", "87.0%"])
        self.assertEqual(str(next(iter(source.wb["지점"].merged_cells.ranges))), "D1:E1")

    def text_file(self, name, raw):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as fh:
            fh.write(raw)
        return path

    def test_csv_in_either_encoding_with_grouped_numbers(self):
        body = "연월,매출,코드\n2026-01,\"1,200\",007\n2026-02,\"1,300\",008\n2026-03,\"1,400\",009\n"
        for encoding in ("utf-8-sig", "cp949"):
            path = self.text_file(f"자료_{encoding}.csv", body.encode(encoding))
            ws = core.ExcelSource(path).wb.worksheets[0]
            self.assertEqual(ws.title, f"자료_{encoding}")
            self.assertEqual(ws["A1"].value, "연월")
            self.assertEqual(ws["B2"].value, 1200)
            self.assertEqual(ws["C2"].value, "007")  # a code keeps its zeros

    def test_a_cell_that_looks_like_a_formula_stays_text(self):
        ws = core.ExcelSource(self.text_file("수식.csv", "이름,값\n=1+1,3\n".encode("utf-8-sig"))).wb.worksheets[0]
        self.assertEqual(ws["A2"].data_type, "s")

    def test_tabs_and_semicolons_are_recognised(self):
        for name, mark in (("a.tsv", "\t"), ("b.csv", ";"), ("c.txt", "\t")):
            ws = core.ExcelSource(self.text_file(name, f"연월{mark}매출\n2026-01{mark}5\n".encode("utf-8-sig"))).wb.worksheets[0]
            self.assertEqual(ws["B2"].value, 5, name)

    def test_an_xls_that_is_really_an_html_table(self):
        page = "<html><meta charset=utf-8><body><table><tr><th>연월</th><th colspan=2>매출</th></tr><tr><td>2026-01</td><td>1,200</td><td>9</td></tr></table></body></html>"
        ws = core.ExcelSource(self.text_file("내보내기.xls", page.encode("utf-8-sig"))).wb.worksheets[0]
        self.assertEqual(ws["A1"].value, "연월")
        self.assertEqual(ws["B2"].value, 1200)
        self.assertEqual(ws["C2"].value, 9)  # the wide heading does not shift the columns under it

    def test_an_xls_that_is_really_an_xlsx(self):
        with open(os.path.join(self.FOLDER, "monthly.xlsx"), "rb") as fh:
            path = self.text_file("이름만.xls", fh.read())
        self.assertEqual(core.ExcelSource(path).wb.worksheets[0].title, "월별")


class SeveralFiles(Base):
    def book(self, name, **tabs):
        path = os.path.join(self.dir, name)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows in tabs.items():
            ws = wb.create_sheet(title)
            for row in rows:
                ws.append(row)
        wb.save(path)
        return path

    def months(self, name, start):
        return [["월", name]] + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), start + m] for m in range(20)]

    def setUp(self):
        super().setUp()
        self.first = self.book("매출.xlsx", 표=self.months("매출", 100))
        self.second = self.book("고객.xlsx", 표=self.months("고객수", 900))

    def test_a_report_holds_an_overview_and_each_files_slide(self):
        source = core.Sources([self.first, self.second])
        plan = suggest.plan_files(source)
        path, out = os.path.join(self.dir, "t.pptx"), os.path.join(self.dir, "o.pptx")
        suggest.make_report(path, "매출 외 1개", plan.sections)
        report = core.run(path, [self.first, self.second], out, thousands=True)
        self.assertEqual(report.count("error"), 0, [(i.key, [x.code for x in i.issues]) for i in report.items])
        slides = Presentation(out).slides
        words = [" ".join(shape.text_frame.text for shape in slide.shapes if shape.has_text_frame) for slide in slides]
        self.assertEqual(len(slides), 3)
        self.assertIn("매출 119", words[0])
        self.assertIn("고객수 919", words[0])
        self.assertIn("고객", words[2])
        self.assertIn("고객수 919", words[2])
        self.assertNotIn("매출 119", words[2])  # each slide reads its own file
        self.assertNotIn("{{", " ".join(words))
        self.assertTrue(any(item.source.startswith("고객") for item in report.items))

    def test_a_file_with_nothing_to_read_is_named_and_the_rest_still_report(self):
        memo = self.book("메모.xlsx", 메모=[["그냥 글자"]])
        plan = suggest.plan_files(core.Sources([self.first, memo]))
        self.assertEqual(plan.files_left, ["메모"])
        self.assertEqual(len(plan.sections), 2)

    def test_same_sheet_names_in_different_files_do_not_mix(self):
        source = core.Sources([self.first, self.second])
        self.assertIn("매출 119", source.text_value("해설:@0@표!A,B")[0])
        self.assertIn("고객수 919", source.text_value("해설:@1@표!A,B")[0])

    def test_the_sentence_about_all_files_marks_where_each_fact_came_from(self):
        source = core.Sources([self.first, self.second])
        plan = suggest.plan_files(source)
        asks = suggest.ask_insights(source, plan.sections, "제목")
        request = asks[core.ALL_FILES][0]
        self.assertIn("2개", request.facts[0].text)
        self.assertTrue(any(fact.id.startswith("file1_") and "[매출]" in fact.text for fact in request.facts))
        self.assertTrue(any(fact.id.startswith("file2_") and "[고객 (고객수)]" in fact.text for fact in request.facts))
        self.assertEqual(request.facts[1].id, "cross_direction")  # what was calculated across the files comes before each file's own
        self.assertTrue(all(len(fact.id) for fact in request.facts))
        self.assertEqual(len({fact.id for fact in request.facts}), len(request.facts))

    def test_a_chosen_sentence_about_all_files_goes_first_and_only_for_the_same_figures(self):
        source = core.Sources([self.first, self.second])
        plan = suggest.plan_files(source)
        asks = suggest.ask_insights(source, plan.sections, "제목")
        request = asks[core.ALL_FILES][0]
        saved = {request.key.casefold(): {"fingerprint": request.fingerprint, "text": "둘 다 늘었습니다."}}
        chosen = suggest.choose_insights(plan.sections, asks, saved)
        self.assertEqual(chosen[0].sentences[0].result, "둘 다 늘었습니다.")
        path, out = os.path.join(self.dir, "t2.pptx"), os.path.join(self.dir, "o2.pptx")
        suggest.make_report(path, "제목", chosen)
        report = core.run(path, [self.first, self.second], out, insights=saved)
        self.assertEqual(next(i for i in report.items if core.ALL_FILES in i.key).status, "ok")
        words = " ".join(shape.text_frame.text for shape in Presentation(out).slides[0].shapes if shape.has_text_frame)
        self.assertIn("둘 다 늘었습니다.", words)
        stale = {request.key.casefold(): {"fingerprint": "다른 수치", "text": "둘 다 늘었습니다."}}
        self.assertNotEqual(suggest.choose_insights(plan.sections, asks, stale)[0].sentences[0].result, "둘 다 늘었습니다.")

    def test_a_figure_that_is_only_part_of_a_file_is_named(self):
        whole = self.book("지하철 승차.xlsx", 표=self.months("설화명곡", 100))
        other = self.book("지하철 하차.xlsx", 표=self.months("설화명곡", 90))
        source = core.Sources([whole, other])
        plan = suggest.plan_files(source)
        lead = plan.sections[0].sentences[0]
        self.assertEqual(lead.code, "cross")
        self.assertIn("지하철 승차 (설화명곡)", lead.result)  # or it reads as all of the subway

    def test_ticked_figures_make_a_section_of_exactly_those(self):
        excel = self.book("틱.xlsx", 표=[["월", "가", "나", "다"]] + [[dt.datetime(2025 + m // 12, m % 12 + 1, 1), 100 + m, 200 + m, 300 + m] for m in range(20)])
        source = core.ExcelSource(excel)
        shape, base, found = suggest.first_choice(source)
        figures = [value for value in base.values if value.name in ("나", "다")]
        section = suggest.section_for(source, shape, base, figures, found)
        words = " ".join(item.result for item in section.sentences)
        self.assertIn("나 219", words)
        self.assertIn("다 319", words)
        self.assertNotIn("가 119", words)
        self.assertEqual([cell[1] for cell in suggest.preview_rows(source, section.table)[0]][1:], ["나", "다"])


class TriedReports(Base):
    """What going through a whole report by eye turned up, on 2026-11-05: right sums, wrong readings."""

    def book(self, name, **tabs):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows in tabs.items():
            ws = wb.create_sheet(title)
            for row in rows:
                ws.append(row)
        path = os.path.join(self.dir, name)
        wb.save(path)
        return path

    def docx(self, name, lines):
        path = os.path.join(self.dir, name)
        body = "".join(f"<w:p><w:r><w:t>{line}</w:t></w:r></w:p>" for line in lines)
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("word/document.xml", f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>')
        return path

    def note(self, section):
        return next(item for item in section.sentences if item.code == "note").result

    def yearly(self, name, start=2015, end=2024):
        return self.book(f"{name}.xlsx", 자료=[["연도", "매출"]] + [[year, 100 + year * 7 % 13] for year in range(start, end + 1)])

    def test_a_first_slide_with_many_files_goes_on_to_another_rather_than_off_the_bottom(self):
        paths = [self.yearly(f"파일{number}") for number in range(9)]
        source = core.Sources(paths)
        plan = suggest.plan_files(source)
        out = os.path.join(self.dir, "많음.pptx")
        suggest.make_report(out, "제목", plan.sections)
        slides = Presentation(out).slides
        height = Presentation(out).slide_height
        words = []
        for slide in slides:
            for shape in slide.shapes:
                self.assertLessEqual(shape.top + shape.height, height, shape.name)
                if shape.has_text_frame:
                    words.append(shape.text_frame.text)
        for number in range(9):
            self.assertIn(f"파일{number}", words)
        self.assertGreater(len(slides), 9 + 1)  # the overview took more than one
        self.assertTrue(any(word.endswith("(1/2)") for word in words))

    def test_a_label_is_not_counted_twice_in_a_sentences_height(self):
        item = suggest.Suggestion("note", "키", "text", result="짧은 문장", label="파일")
        out = os.path.join(self.dir, "높이.pptx")
        suggest.make_report(out, "제목", [suggest.Section("", [item, replace(item, label="다음")])])
        tops = sorted(shape.top for shape in Presentation(out).slides[0].shapes if shape.has_text_frame and shape.text_frame.text == "다음" or shape.has_text_frame and shape.text_frame.text == "파일")
        self.assertAlmostEqual((tops[1] - tops[0]) / 914400, suggest.LABEL_HEIGHT + 18 * suggest.LINE_HEIGHT / 72 + suggest.SENTENCE_GAP, places=2)

    def test_yearly_files_say_no_month_in_the_comparison_across_files(self):
        request = core.Sources([self.yearly("하나"), self.book("둘.xlsx", 자료=[["월", "매출"]] + [[f"2023-{m:02d}", 100 + m] for m in range(1, 13)])])
        plan = suggest.plan_files(request)
        cross = next(item for item in plan.sections[0].sentences if item.code == "cross").result
        self.assertIn("2024년,", cross)
        self.assertNotIn("2024년 1월", cross)
        self.assertIn("2023년 12월", cross)

    def test_no_change_is_still_a_change_to_list(self):
        one = self.book("같음.xlsx", 자료=[["연도", "매출"]] + [[year, 300] for year in range(2020, 2025)])
        two = self.book("늘음.xlsx", 자료=[["연도", "매출"]] + [[year, 100 + (year - 2020) ** 2] for year in range(2020, 2025)])
        cross = next(item for item in suggest.plan_files(core.Sources([one, two])).sections[0].sentences if item.code == "cross").result
        self.assertIn("같음 (매출) 변동 없음", cross)

    def test_a_figure_that_stops_early_gives_way_to_the_grand_total(self):
        years = ["2019-20", "2020-21", "2021-22", "2022-23", "2023-24"]
        rows = [["Program", "Category"] + years,
                ["Child Program Total", None, 3000, 3100, 3006, None, None],
                ["Family", "Family", 50000, 51000, 52000, 53000, 54000],
                ["Migration Program Total", None, 137000, 157000, 140000, 195000, 190000],
                ["Grand Total", None, 140000, 160000, 143000, 195000, 190000]]
        path = self.book("이민.xlsx", Outcome=rows)
        section = suggest.plan_report(core.ExcelSource(path)).sections[0]
        self.assertEqual(section.figures[0].name, "Grand Total")
        self.assertIn("2023-24 Grand Total 190,000", self.note(section))

    def test_a_total_that_reaches_the_newest_period_stays_first(self):
        rows = [["구분"] + [f"202{y}년" for y in range(0, 6)], ["남 합계"] + [10 + y for y in range(6)], ["여 합계"] + [50 + y for y in range(6)]]
        base = suggest.read_shape(core.ExcelSource(self.book("합계.xlsx", 표=rows)), "표").bases[0]
        self.assertEqual(suggest.default_value(base).name, "남 합계")

    def test_a_period_still_running_is_not_compared_with_ended_ones(self):
        rows = [["Fin. Year", "수"]] + [[f"{y}-{(y + 1) % 100:02d}", 8000 + y] for y in range(2018, 2025)] + [["2025-26", 960]]
        wb = openpyxl.Workbook()
        wb.active.title = "자료"
        for row in rows:
            wb.active.append(row)
        info = wb.create_sheet("Information")
        info.append(["Data as at 30-Sep-2025, for applications created on or after 01-Jul-2018."])
        wb.properties.modified = dt.datetime(2025, 10, 3)
        path = os.path.join(self.dir, "미완결.xlsx")
        wb.save(path)
        wb2 = openpyxl.load_workbook(path)  # saving stamps the time, so it is put back in the file itself
        source = core.ExcelSource(path)
        source.wb.properties.modified = dt.datetime(2025, 10, 3)
        note = source.text_value("해설:자료!A,B")[0]
        read = source.text_value("읽는법:자료!A,B")[0]
        wb2.close()
        self.assertIn("2025-26은 자료 기준일 2025-09-30에 아직 끝나지 않은 기간", note)
        self.assertIn("2025-26 수 960.", note)
        self.assertNotIn("감소", note)
        self.assertNotIn("끝나지 않은", read)  # said once, on the slide
        self.assertIn("끝난 기간 중 가장 최근은 2024-25 10,024입니다", read)
        self.assertNotRegex(read, "960|가장 낮은")

    def test_a_note_about_another_time_does_not_become_the_date_of_the_data(self):
        wb = openpyxl.Workbook()
        wb.active.title = "자료"
        for year in range(2015, 2026):
            wb.active.append([f"{year}-{(year + 1) % 100:02d}", 100 + year])
        wb.create_sheet("Notes").append(["Rules as at 01-Jul-2012 apply."])
        path = os.path.join(self.dir, "다른날짜.xlsx")
        wb.save(path)
        source = core.ExcelSource(path)
        source.wb.properties.modified = dt.datetime(2025, 10, 3)
        self.assertIn("파일을 저장한 2025-10-03", source.text_value("해설:자료!A,B")[0])

    def test_a_format_that_hides_small_counts_shows_large_ones_with_commas(self):
        fmt = "[>=5]##,###,##0;[<=-5]\\-##,###,##0;\\<\\5"
        self.assertEqual(format_value(190000, fmt), ("190,000", True))
        self.assertEqual(format_value(-190000, fmt), ("-190,000", True))
        self.assertEqual(format_value(3, fmt), ("<5", True))
        self.assertEqual(format_value(-3, fmt), ("<5", True))
        self.assertEqual(format_value(100, "[<50]0;[>=50]0.0"), ("100.0", True))
        self.assertEqual(format_value(-5, "[<50]0;[>=50]0.0"), ("-5", True))

    def test_a_rates_reading_does_not_repeat_its_sentence(self):
        rows = [["구분", 2019, 2020, 2021, 2022, 2023], ["Total children", 83, 84, 85, 85.8, 87], [], ["Unit:"], ["Per cent"]]
        source = core.ExcelSource(self.book("비율.xlsx", 표=rows))
        section = suggest.plan_report(source).sections[0]
        note, reading = (next(item.result for item in section.sentences if item.code == code) for code in ("note", "reading"))
        self.assertIn("+1.2포인트", note)
        self.assertNotIn("+1.2포인트", reading)
        self.assertIn("가장 높은", reading)

    def test_one_english_file_says_its_period_in_the_reports_language(self):
        self.assertEqual(core._period_in("2023", "ko"), "2023년")
        self.assertEqual(core._period_in("Aug 2021", "ko"), "2021년 8월")
        self.assertEqual(core._period_in("Q3 2021", "ko"), "2021년 3분기")
        self.assertEqual(core._period_in("2021년 8월", "en"), "Aug 2021")
        self.assertEqual(core._period_in("2024-25", "ko"), "2024-25")

    def test_machine_names_and_footnote_marks_read_as_words(self):
        self.assertEqual(facts.readable_name("RETAINED_CATCH_KG"), "Retained catch (kg)")
        self.assertEqual(facts.readable_name("합계 RETAINED_CATCH_KG"), "합계 Retained catch (kg)")
        self.assertEqual(facts.readable_name("제주(A) 2025년 (월계)"), "제주 2025년 (월계)")
        self.assertEqual(facts.readable_name("# EAF Items (FY Subtotal)"), "EAF Items (FY Subtotal)")
        self.assertEqual(facts.readable_name("Plan (A)"), "Plan (A)")
        self.assertEqual(facts.readable_name("매출"), "매출")
        self.assertEqual(facts.readable_name("E_0000_WH"), "E_0000_WH")  # a code, not words

    def test_a_report_says_readable_headings_where_a_template_keeps_the_sheets(self):
        rows = [["YEAR", "RETAINED_CATCH_KG"]] + [[year, 1000 + year] for year in range(2015, 2025)]
        source = core.ExcelSource(self.book("어획.xlsx", 표=rows))
        self.assertTrue(source.text_value("해설:표!A,B")[0].startswith("2024년 Retained catch (kg) 3,024"))
        self.assertEqual(source.range_cells("최신3:표!A,B")[0][0][1][1], "RETAINED_CATCH_KG")
        source.thousands = True
        self.assertEqual(source.range_cells("최신3:표!A,B")[0][0][1][1], "Retained catch (kg)")

    def test_a_laid_out_table_does_not_call_its_dates_by_the_wrong_heading(self):
        across = [["Visa Category"] + [f"20{y}-{y + 1}" for y in range(15, 21)], ["Grand Total"] + [100 + y * 3 for y in range(6)]]
        source = core.ExcelSource(self.book("가로.xlsx", 표=across))
        source.thousands = True
        section = suggest.plan_report(source).sections[0]
        key = core.MARKER.fullmatch(section.table.typed).group(2)
        self.assertEqual(source.range_cells(key)[0][0][0][1], "기간")

    def test_a_blank_item_heading_is_filled_in_a_laid_out_table(self):
        rows = [[None, "매출"], ["가", 30], ["나", 20], ["다", 10], ["라", 5]]
        source = core.ExcelSource(self.book("빈칸.xlsx", 표=rows))
        source.thousands = True
        self.assertEqual(source.range_cells("구성비3:표!A,B")[0][0][0][1], "항목")

    def test_a_table_with_no_dates_says_the_period_in_its_title(self):
        rows = [["2026년 6월 건축 통계표"], [None], ["구역", "건수"]] + [[f"구{n}", 10 + n * n] for n in range(8)]
        section = suggest.plan_report(core.ExcelSource(self.book("건축.xlsx", 자료=rows))).sections[0]
        self.assertTrue(section.heading.startswith("2026년 6월"), section.heading)
        self.assertEqual(section.period, "2026년 6월")

    def test_a_span_in_a_title_is_not_one_period(self):
        rows = [["2020년~2024년 통계표"], [None], ["구역", "건수"]] + [[f"구{n}", 10 + n * n] for n in range(8)]
        self.assertEqual(suggest.plan_report(core.ExcelSource(self.book("기간.xlsx", 자료=rows))).sections[0].period, "")

    def test_two_years_beside_each_other_are_compared_under_one_kind(self):
        rows = [[None, None, "2025년", "2024년"]]
        for name, month, year in (("중국", 100, 900), ("대만", 61, 400), ("일본", 47, 300), ("태국", 23, 100), ("미국", 11, 90)):
            rows += [[name, "월계", month, month // 2], [None, "누계", year, year // 2]]
        section = suggest.plan_report(core.ExcelSource(self.book("두해.xlsx", 표=rows))).sections[0]
        change = next(item for item in section.sentences if item.code == "change").result
        self.assertIn("119 → 242", change)  # 월계 only: 누계 would have made it far larger
        self.assertIn("+103.4%", change)

    def test_a_kept_doubt_is_said_on_the_slide(self):
        rows = [["연월", "회사", "발량"]] + [[f"2025-{m:02d}", f"회사{c}", 42.4 + (m + c) % 3 * 0.1] for m in range(1, 7) for c in range(10)]
        plan = suggest.plan_report(core.ExcelSource(self.book("의심.xlsx", 표=rows)))
        self.assertTrue(plan.doubts)
        words = "※ 확인되지 않은 값입니다."
        sections = suggest.with_doubt_notes(plan.sections, [("", words)])
        self.assertEqual(sections[0].sentences[-1].typed, words)
        self.assertEqual(len(plan.sections[0].sentences) + 1, len(sections[0].sentences))
        out = os.path.join(self.dir, "표시.pptx")
        suggest.make_report(out, "제목", sections)
        self.assertIn(words, [shape.text_frame.text for shape in Presentation(out).slides[0].shapes if shape.has_text_frame])

    def test_a_document_says_what_matters_and_not_a_pointer_or_a_glossary_entry(self):
        import documents
        pointer = "The shaded cells in the Table 1 below show the years for which data is included for each fishery, noting that zero catch shows nothing at all."
        glossary = "Australian Fishing Zone: Waters adjacent to Australia which extend from defined baselines to 200 nm seawards, but not including coastal and excepted waters."
        about = "This dataset shows the annual landed catch in kilograms, by species and fishery, for fisheries managed by the authority in each calendar year."
        care = "Care should be taken when interpreting trends in this data because there are many factors that can affect catches in each and every year."
        document = documents.Document("x", "x", paragraphs=[about, pointer, care, glossary])
        self.assertEqual(documents.key_paragraphs(document), [about, care])

    def test_a_document_in_a_workbooks_folder_comes_just_before_it(self):
        folder = os.path.join(self.dir, "함께")
        os.makedirs(folder)
        apart = self.yearly("따로")
        with_it = os.path.join(folder, "같이.xlsx")
        shutil.copy(self.yearly("원본"), with_it)
        long = "이 문서는 자료의 범위와 집계 방식을 설명하는 긴 문단이며, 값을 해석할 때 유의할 점을 충분히 길게 적어 둔 문장입니다. 그래서 슬라이드에 그대로 옮길 수 있습니다."
        word = self.docx("설명.docx", [long])
        shutil.move(word, os.path.join(folder, "설명.docx"))
        plan = suggest.plan_files(core.Sources([apart, with_it, os.path.join(folder, "설명.docx")]))
        titles = [section.title for section in plan.sections]
        self.assertEqual(titles, ["", "따로", "같이 · 설명", "같이"])

    def test_documents_stay_in_front_when_a_folder_says_nothing(self):
        one, two = self.yearly("하나"), self.yearly("둘")
        word = self.docx("설명.docx", ["이 문서는 자료의 범위와 집계 방식을 설명하는 긴 문단이며, 값을 해석할 때 유의할 점을 충분히 길게 적어 둔 문장입니다. 그래서 슬라이드에 그대로 옮길 수 있습니다."])
        titles = [section.title for section in suggest.plan_files(core.Sources([one, two, word])).sections]
        self.assertEqual(titles, ["", "설명", "하나", "둘"])


class CompanyTemplates(Base):
    """What a company-style template made in real PowerPoint and Excel turned up, on 2026-11-05."""

    def table_book(self, rows, merged=()):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "표"
        for row in rows:
            ws.append(row)
        for cells in merged:
            ws.merge_cells(cells)
        path = os.path.join(self.dir, "표.xlsx")
        wb.save(path)
        return path

    def table_template(self, rows, columns, key, merges=(), top=1.0, height=0.3, below=None):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        frame = slide.shapes.add_table(rows, columns, Inches(0.5), Inches(top), Inches(8), Inches(height * rows))
        for row in frame.table.rows:
            row.height = Inches(height)
        frame._element.xpath(".//p:cNvPr")[0].set("descr", "{{표:" + key + "}}")
        for first, last in merges:
            frame.table.cell(*first).merge(frame.table.cell(*last))
        if below is not None:
            box = slide.shapes.add_textbox(Inches(0.5), Inches(below), Inches(8), Inches(0.5))
            box.name = "아래 글"
            box.text_frame.text = "설명"
        path = os.path.join(self.dir, "표.pptx")
        prs.save(path)
        return path

    DATA = [["지역", "지점", "당월", None], [None, None, "매출", "목표"]] + [[f"권역{n}", f"지점{n}", 100 + n, 200 + n] for n in range(5)]

    def filled(self, template, excel, key="error"):
        out = os.path.join(self.dir, "out.pptx")
        if os.path.exists(out):
            os.remove(out)
        report = core.run(template, excel, out)
        return report, Presentation(out)

    def test_a_two_row_merged_heading_can_grow_and_keeps_its_merges(self):
        excel = self.table_book(self.DATA, ("A1:A2", "B1:B2", "C1:D1"))
        template = self.table_template(4, 4, "표!A1:D7", (((0, 0), (1, 0)), ((0, 1), (1, 1)), ((0, 2), (0, 3))))
        report, prs = self.filled(template, excel)
        self.assertEqual(report.count("error"), 0, [(i.key, [x.code for x in i.issues]) for i in report.items])
        table = next(s for s in prs.slides[0].shapes if s.has_table).table
        self.assertEqual(len(table.rows), 7)
        self.assertEqual(table.cell(0, 0).text, "지역")
        self.assertTrue(table.cell(0, 0).is_merge_origin)
        self.assertEqual(table.cell(0, 0).span_height, 2)
        self.assertEqual(table.cell(0, 2).span_width, 2)
        self.assertEqual(table.cell(6, 3).text, "204")

    def test_a_two_row_merged_heading_can_shrink(self):
        excel = self.table_book(self.DATA[:3], ("A1:A2", "B1:B2", "C1:D1"))
        template = self.table_template(6, 4, "표!A1:D3", (((0, 0), (1, 0)), ((0, 1), (1, 1))))
        report, prs = self.filled(template, excel)
        self.assertEqual(report.count("error"), 0)
        self.assertEqual(len(next(s for s in prs.slides[0].shapes if s.has_table).table.rows), 3)

    def test_cells_merged_in_the_body_still_refuse_to_resize(self):
        excel = self.table_book(self.DATA)
        template = self.table_template(4, 4, "표!A1:D7", (((2, 0), (3, 0)),))
        report, _prs = self.filled(template, excel)
        self.assertEqual([x.code for i in report.items for x in i.issues], ["merged_resize"])

    def test_a_merge_cannot_be_cut_through_by_shrinking(self):
        excel = self.table_book(self.DATA[:1])
        template = self.table_template(4, 4, "표!A1:D1", (((0, 0), (1, 0)),))
        report, _prs = self.filled(template, excel)
        self.assertEqual([x.code for i in report.items for x in i.issues], ["merged_resize"])

    def test_a_range_starting_inside_a_merge_shows_the_word_a_person_sees(self):
        source = core.ExcelSource(self.table_book(self.DATA, ("A1:A2", "B1:B2", "C1:D1")))
        rows = source.range_cells("표!A2:D3")[0]
        self.assertEqual([cell[1] for cell in rows[0]], ["지역", "지점", "매출", "목표"])
        whole = source.range_cells("표!A1:D2")[0]
        self.assertEqual(whole[1][0][1], "")  # the word is already in the range, above it

    def test_the_latest_date_is_not_a_month_with_only_its_target_written(self):
        rows = [["월", "실적", "목표", "이익"]] + [[f"2026-{m:02d}", 10 + m, 20, 1 + m] for m in range(1, 11)] + [["2026-11", None, 21, None], ["2026-12", None, 22, None]]
        source = core.ExcelSource(self.table_book(rows))
        text, _where, issues = source.text_value("최신:표!A")
        self.assertEqual(text, "2026-10")
        self.assertEqual([issue.code for issue in issues], ["latest_older"])

    def test_the_latest_date_stays_when_rows_are_filled(self):
        rows = [["월", "실적", "목표"]] + [[f"2026-{m:02d}", 10 + m, 20] for m in range(1, 11)]
        text, _where, issues = core.ExcelSource(self.table_book(rows)).text_value("최신:표!A")
        self.assertEqual((text, issues), ("2026-10", []))

    def test_rows_shown_with_empty_columns_are_mentioned(self):
        rows = [["월", "실적", "목표"]] + [[f"2026-{m:02d}", 10 + m, 20] for m in range(1, 11)] + [["2026-11", None, 21]]
        source = core.ExcelSource(self.table_book(rows))
        codes = [issue.code for issue in source.range_cells("최신3:표!A,B,C")[2]]
        self.assertIn("latest_partial_rows", codes)
        self.assertNotIn("latest_partial_rows", [issue.code for issue in source.range_cells("최신3:표!A,C")[2]])

    def test_this_year_before_last_year_is_warned_about(self):
        rows = [["제품", "매출", "전년"], ["가", 120, 100], ["나", 80, 90], ["다", 50, 40], ["라", 30, 20]]
        codes = [issue.code for issue in core.ExcelSource(self.table_book(rows)).text_value("읽는법:표!A,B,C")[2]]
        self.assertIn("cross_backwards", codes)
        self.assertNotIn("cross_backwards", [issue.code for issue in core.ExcelSource(self.table_book(rows)).text_value("읽는법:표!A,C,B")[2]])

    def test_two_years_written_newest_first_are_warned_about(self):
        rows = [["제품", "2025년", "2024년"], ["가", 120, 100], ["나", 80, 90], ["다", 50, 40]]
        source = core.ExcelSource(self.table_book(rows))
        self.assertIn("cross_backwards", [issue.code for issue in source.text_value("읽는법:표!A,B,C")[2]])
        self.assertNotIn("cross_backwards", [issue.code for issue in source.text_value("읽는법:표!A,C,B")[2]])

    def test_a_table_that_grows_over_a_shape_below_is_warned_about(self):
        rows = [["지점", "매출"]] + [[f"지점{n}", 100 + n] for n in range(9)]
        excel = self.table_book(rows)
        report, _prs = self.filled(self.table_template(3, 2, "표!A1:B10", top=1.0, height=0.4, below=3.0), excel)
        found = [x for i in report.items for x in i.issues if x.code == "table_covers"]
        self.assertEqual(found[0].params["name"], "아래 글")
        report, _prs = self.filled(self.table_template(3, 2, "표!A1:B3", top=1.0, height=0.4, below=3.0), excel)
        self.assertEqual([x.code for i in report.items for x in i.issues if x.code == "table_covers"], [])

    def test_smartart_is_listed_once_and_says_which_slide(self):
        wb_path = self.table_book([["이름"], ["값"]])
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        second = prs.slides.add_slide(prs.slide_layouts[6])
        second.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "{{표!A2}}"
        plain = os.path.join(self.dir, "plain.pptx")
        prs.save(plain)
        patched = os.path.join(self.dir, "smart.pptx")
        with zipfile.ZipFile(plain) as src, zipfile.ZipFile(patched, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = data.replace(b"</Types>", (b'<Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>'
                                                      b'<Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/></Types>'))
                if info.filename == "ppt/slides/_rels/slide2.xml.rels":
                    data = data.replace(b"</Relationships>", (b'<Relationship Id="rId98" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>'
                                                              b'<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/></Relationships>'))
                dst.writestr(info, data)
            dst.writestr("ppt/diagrams/data1.xml", "<x>{{매출}}</x>")
            dst.writestr("ppt/diagrams/drawing1.xml", "<x>{{매출}}</x>")
        report = core.run(patched, wb_path)
        found = [item for item in report.items if item.kind == "unsupported"]
        self.assertEqual([item.key for item in found], ["/ppt/diagrams/data1.xml"])
        self.assertEqual(found[0].issues[0].params["slide"], 2)
        self.assertIn("슬라이드 2", i18n.text("ko", "issue_unsupported_location", **found[0].issues[0].params))


class WordDocuments(Base):
    def docx(self, name, lines, styles=None):
        """A small real .docx, written with no Word needed."""
        import zipfile
        from xml.sax.saxutils import escape
        styles = styles or {}
        body = "".join(f'<w:p>{f"""<w:pPr><w:pStyle w:val="{styles[n]}"/></w:pPr>""" if n in styles else ""}<w:r><w:t>{escape(line)}</w:t></w:r></w:p>' for n, line in enumerate(lines))
        xml = f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>'
        path = os.path.join(self.dir, name)
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("word/document.xml", xml)
        return path

    def workbook(self, name="노선.xlsx"):
        """Quarters across the top, routes down the side: 가람 falls by half after the middle, 나루 does not move."""
        path = os.path.join(self.dir, name)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "노선"
        quarters = [dt.datetime(2020 + q // 4, q % 4 * 3 + 3, 1) for q in range(16)]
        ws.append(["노선"] + quarters)
        ws.append(["가람"] + [1000] * 8 + [500] * 8)
        ws.append(["나루"] + [300] * 16)
        ws.append(["다온"] + [10] * 16)
        ws.append(["라온"] + [20] * 16)
        wb.save(path)
        return path

    LINES = ["항공 통계", "노선 관련 기록", "2021년 12월 A항공이 가람 노선에서 철수했습니다.", "2021년 12월 B항공이 나루 노선에 들어왔습니다.", "다른 설명 문단입니다.", "2022년 3월 C항공이 어느 곳에도 없는 노선에 들어왔습니다."]

    def test_dated_entries_are_found_in_either_language(self):
        import documents
        for text, year, month in (("May 2000Horizon withdraws", 2000, 5), ("January 2001: Yanda withdraws", 2001, 1), ("2021년 12월 A항공 철수", 2021, 12), ("2021.03 B항공 진입", 2021, 3), ("14 March 2013 Brindabella replaced", 2013, 3), ("June 30, 2015 Order gazetted", 2015, 6)):
            events = documents._events([(text, False)])
            self.assertEqual((events[0].date.year, events[0].date.month), (year, month), text)
        self.assertEqual(documents._events([("2021 was a good year", False), ("매출 2021년 12월 기준", False)]), [])

    def test_a_short_line_between_dated_ones_keeps_the_date(self):
        import documents
        events = documents._events([("May 2015 Order on Mudgee.", False), ("Order on Narrabri.", False), ("August 2015 Order on Cobar.", False)])
        self.assertEqual([(e.date.month, e.text) for e in events], [(5, "Order on Mudgee."), (5, "Order on Narrabri."), (8, "Order on Cobar.")])

    def test_a_name_must_be_a_whole_word(self):
        import documents
        self.assertTrue(documents.mentions("REX entered the Armidale-Sydney route", "Armidale"))
        self.assertFalse(documents.mentions("Rexburg route", "Rex"))
        self.assertTrue(documents.mentions("가람 노선에서", "가람"))

    def test_the_means_leave_out_the_period_it_happened_in(self):
        source = core.Sources([self.workbook(), self.docx("문서.docx", self.LINES)])
        _document, links = source.links()[0]
        top = links[0]
        self.assertEqual((top.item.name, top.before, top.after, top.window), ("가람", 1000, 500, 4))
        self.assertEqual(round(top.change, 1), -50.0)
        self.assertEqual(len(links), 2)  # 가람 and 나루; the route no one has is not forced onto anything

    def test_too_little_on_one_side_is_not_compared(self):
        import documents
        points = [type("P", (), {"date": dt.datetime(2020, m, 1), "value": m})() for m in (3, 6, 9, 12)]
        item = documents.Item("가람", "표", points)
        self.assertIsNone(documents.link(documents.Event(dt.datetime(2020, 6, 1), "가람"), item))

    def test_a_report_holds_the_document_slide_and_says_it_is_only_the_same_time(self):
        excel, word = self.workbook(), self.docx("문서.docx", self.LINES, {0: "Title", 1: "Heading1"})
        source = core.Sources([excel, word])
        plan = suggest.plan_files(source)
        path, out = os.path.join(self.dir, "t.pptx"), os.path.join(self.dir, "o.pptx")
        suggest.make_report(path, "제목", plan.sections)
        report = core.run(path, [excel, word], out, thousands=True)
        self.assertEqual(report.count("error"), 0, [(i.key, [x.code for x in i.issues]) for i in report.items])
        slides = Presentation(out).slides
        words = " ".join(shape.text_frame.text for shape in slides[1].shapes if shape.has_text_frame)
        self.assertIn("항공 통계", words)
        self.assertIn("4건 중 2건", words)  # a short line between dated ones counts as one, and names nothing here
        self.assertIn("같은 시기", words)
        self.assertIn("원인인지는", words)
        table = next(shape for shape in slides[1].shapes if shape.has_table).table
        self.assertEqual([table.cell(1, c).text for c in (0, 1, 3, 4, 5)], ["2021년 12월", "가람", "1,000", "500", "-50.0%"])
        self.assertIn("A항공", table.cell(1, 2).text)
        self.assertNotIn("{{", " ".join(table.cell(r, c).text for r in range(len(table.rows)) for c in range(6)))

    def test_the_model_is_given_the_line_up_and_told_not_to_claim_a_cause(self):
        excel, word = self.workbook(), self.docx("문서.docx", self.LINES)
        source = core.Sources([excel, word])
        plan = suggest.plan_files(source)
        request = suggest.ask_insights(source, plan.sections, "제목")[core.ALL_FILES][0]
        links = [fact for fact in request.facts if fact.id.startswith("doc1_link")]
        self.assertEqual(len(links), 2)
        self.assertIn("가람", links[0].text)
        self.assertIn("1,000", links[0].numbers)
        self.assertIn("500", links[0].numbers)
        self.assertIn("때문에", request.facts[0].text)

    def test_a_document_with_no_dates_gives_its_paragraphs_to_the_model(self):
        long = "이 문서는 분기별 노선 승객 수의 집계 방식과 범위를 설명하는 문단이며 숫자 1,234가 들어 있고 충분히 길게 쓴 문장입니다. 그래서 모델이 읽을 수 있는 설명이 됩니다."
        excel, word = self.workbook(), self.docx("설명.docx", ["설명 문서", long])
        source = core.Sources([excel, word])
        plan = suggest.plan_files(source)
        self.assertEqual(plan.sections[1].sentences[0].code, "doc_quote")
        self.assertEqual(plan.sections[1].sentences[0].result, long)  # in its own words
        self.assertIsNone(plan.sections[1].table)
        request = suggest.ask_insights(source, plan.sections, "제목")[core.ALL_FILES][0]
        self.assertTrue(any(fact.id == "doc1_text1" and "1,234" in fact.numbers for fact in request.facts))

    def test_the_documents_figures_are_looked_for_in_the_workbooks(self):
        import documents
        word = self.docx("수치.docx", ["올해 승객은 1,000명, 다른 노선은 7,777명이고 2021년 자료이며 3개 표가 있습니다."])
        source = core.Sources([self.workbook(), word])
        found, missing = documents.check_figures(source.documents[0], source._all_values())
        self.assertEqual((found, missing), (["1,000"], ["7,777"]))
        self.assertIn("찾지 못한 수치: 7,777", documents.figures_sentence(found, missing, "ko"))

    def test_small_numbers_do_not_count_as_found_and_little_overlap_is_said(self):
        import documents
        word = self.docx("겹침.docx", ["접수는 2023, 그 뒤 5,276건과 4,117건, 133건과 500건이었습니다."])
        source = core.Sources([self.workbook(), word])
        self.assertEqual(sorted(documents.figures_in(source.documents[0])), ["4,117", "5,276"])
        found, missing = documents.check_figures(source.documents[0], source._all_values())
        self.assertIn("서로 다른 범위", documents.figures_sentence(found, missing, "ko"))

    def test_highlights_keep_the_sharpest_falls_as_well_as_the_largest_moves(self):
        import documents
        event = documents.Event(dt.datetime(2020, 1, 1), "기록")
        item = lambda name: documents.Item(name, "표", [])  # noqa: E731
        links = [documents.Link(event, item("큰1"), 1000, 2000, 4), documents.Link(event, item("큰2"), 1000, 1900, 4), documents.Link(event, item("큰3"), 1000, 1800, 4),
                 documents.Link(event, item("큰4"), 1000, 1700, 4), documents.Link(event, item("급락"), 500, 20, 4), documents.Link(event, item("아주작음"), 5, 0, 4)]
        names = [found.item.name for found in documents.highlights(links, 4)]
        self.assertIn("큰1", names)
        self.assertIn("급락", names)  # small, and nearly gone
        self.assertNotIn("아주작음", names)  # a handful of people

    def test_looking_for_a_documents_figures_makes_no_new_cells(self):
        path = os.path.join(self.dir, "넓은.xlsx")
        wb = openpyxl.Workbook()
        wb.active["A1"] = 1500
        wb.active.cell(1, 16383).value = "멀리 있는 칸"
        wb.save(path)
        source = core.Sources([path, self.docx("문서.docx", ["숫자 1,500이 있습니다."])])
        ws = source.sources[0].wb.worksheets[0]
        before = len(ws._cells)
        self.assertIn(1500.0, source._all_values())
        self.assertEqual(len(ws._cells), before)  # or a sheet like this one takes gigabytes

    def test_a_word_file_alone_asks_for_a_workbook(self):
        with self.assertRaises(core.SlideFillError) as caught:
            core.Sources([self.docx("혼자.docx", self.LINES)])
        self.assertEqual(caught.exception.code, "doc_needs_workbook")

    def test_a_damaged_word_file_is_said_plainly(self):
        path = os.path.join(self.dir, "깨짐.docx")
        with open(path, "wb") as fh:
            fh.write(b"not a zip file at all")
        with self.assertRaises(core.SlideFillError) as caught:
            core.Sources([self.workbook(), path])
        self.assertEqual(caught.exception.code, "file_not_office")

    def test_an_outside_entity_in_a_word_file_is_not_fetched(self):
        import zipfile
        import documents
        path = os.path.join(self.dir, "엔티티.docx")
        evil = '<?xml version="1.0"?><!DOCTYPE d [<!ENTITY x SYSTEM "file:///c:/windows/win.ini">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>앞&x;뒤</w:t></w:r></w:p></w:body></w:document>'
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("word/document.xml", evil)
        document = documents.read(path, "엔티티")
        self.assertNotIn("fonts", " ".join(document.paragraphs).lower())


class CrossFileReadings(unittest.TestCase):
    def test_opposite_directions_are_said_with_their_figures(self):
        found = facts.cross_file_readings([("생산량", -10.2, "2026년 9월", "year"), ("공급량", 25.3, "2026년 9월", "year")], "ko")
        self.assertEqual(found[0].id, "cross_direction")
        self.assertEqual(found[0].text, "2026년 9월 전년 같은 기간보다 생산량 10.2% 감소, 공급량 25.3% 증가로 방향이 서로 다릅니다.")
        self.assertEqual(set(found[0].numbers), {"10.2", "25.3"})

    def test_the_same_direction_is_said_too(self):
        text = facts.cross_file_readings([("가", 1.0, "2026년 1월", "prev"), ("나", 2.0, "2026년 1월", "prev")], "ko")[0].text
        self.assertTrue(text.endswith("모두 늘었습니다."), text)
        self.assertIn("all fell", facts.cross_file_readings([("a", -1.0, "2026-01", "year"), ("b", -2.0, "2026-01", "year")], "en")[0].text)

    def test_different_times_are_not_compared(self):
        found = facts.cross_file_readings([("가", 1.0, "2026년 8월", "year"), ("나", -2.0, "2024년 1월", "year")], "ko")
        self.assertEqual(found[0].id, "cross_periods")
        self.assertIn("2024년 1월", found[0].text)
        self.assertNotIn("다릅니다로", found[0].text)
        self.assertNotIn("1.0", found[0].text)

    def test_different_bases_are_not_compared(self):
        self.assertEqual(facts.cross_file_readings([("가", 1.0, "2026년 1월", "year"), ("나", 2.0, "2026년 1월", "prev")], "ko")[0].id, "cross_basis")

    def test_a_file_with_no_percent_change_is_left_out_and_one_alone_says_nothing(self):
        self.assertEqual(facts.cross_file_readings([("가", 1.0, "2026년 1월", "year"), ("나", None, "2026년 1월", "year")], "ko"), [])

    def test_the_checks_catch_what_the_model_got_wrong(self):
        import insight
        streak = facts.Fact("streak", "2회 연속 증가", ("2",))
        rank = facts.Fact("rank", "최근 30개 값 중 17번째로 높음", ("30", "17"))
        year = facts.Fact("year", "전년 같은 달 대비 +36.6%", ("36.6",))
        pack = (streak, rank, year)
        self.assertTrue(insight.check({"statement": "전년 대비 36.6% 늘어 지속적인 증가세입니다.", "fact_ids": ["year"]}, pack, "ko"), "one change is not a trend")
        self.assertEqual(insight.check({"statement": "2회 연속 증가하는 추세입니다.", "fact_ids": ["streak"]}, pack, "ko"), [])
        wrong = {"이는 12개월 연속 증가 추세입니다.": ["streak"], "최근 30개 값 중 상위 17%에 해당합니다.": ["rank"], "최근 30개 값 중 상위권에 속합니다.": ["rank"], "12번 연속 높습니다.": ["streak"], "전년 대비 감소는 정책의 결과일 수 있습니다.": ["year"]}
        for statement, cited in wrong.items():
            self.assertTrue(insight.check({"statement": statement, "fact_ids": cited}, pack, "ko"), statement)
        fine = {"2회 연속 증가했습니다.": ["streak"], "최근 30개 값 중 17번째로 높습니다.": ["rank"], "2026년 8월 값은 전년 같은 달보다 36.6% 높습니다.": ["year"]}
        for statement, cited in fine.items():
            self.assertEqual(insight.check({"statement": statement, "fact_ids": cited}, pack, "ko"), [], statement)


class SentenceAccuracy(unittest.TestCase):
    """Each of these sentences was written by the real model on real files, passed the old checks, and was wrong."""

    PACK = (facts.Fact("subject", "이번 기간은 2026년 8월입니다.", (), "내국인", "2026년 8월"),
            facts.Fact("year", "전년 같은 달 2025년 8월 206 대비 -6.8%", ("2025", "8", "206", "6.8")),
            facts.Fact("run", "전년 같은 달보다 6번 연속 낮습니다.", ("6",)), facts.Fact("read_year_run", "전년 같은 달보다 6번 연속 낮습니다.", ("6",)),
            facts.Fact("rank", "최근 21개 값 중 19번째로 높음", ("21", "19")), facts.Fact("avg", "직전 7개 평균 대비 -27.6%", ("7", "27.6")),
            facts.Fact("high", "최근 21개 중 최고 309 (2024년 5월)", ("21", "309")))

    def problems(self, statement, *cited, pack=None):
        import insight
        return insight.check({"statement": statement, "fact_ids": list(cited)}, pack or self.PACK, "ko")

    def test_a_figure_that_is_part_of_a_file_must_be_named(self):
        self.assertIn("scope", self.problems("2026년 8월 매출액은 전년 같은 달보다 6.8% 낮습니다.", "year"))
        self.assertNotIn("scope", self.problems("2026년 8월 내국인 매출액은 전년 같은 달보다 6.8% 낮습니다.", "year"))

    def test_a_whole_needs_no_name(self):
        pack = (facts.Fact("subject", "", (), "합계", "2026년 8월"), facts.Fact("year", "대비 -6.8%", ("6.8",)))
        self.assertNotIn("scope", self.problems("2026년 8월은 전년보다 6.8% 낮습니다.", "year", pack=pack))

    def test_a_judgement_is_not_a_fact(self):
        self.assertIn("judged", self.problems("내국인은 21개 중 19번째로 여전히 높은 수준입니다.", "rank"))
        self.assertIn("judged", self.problems("내국인은 상대적으로 안정적인 수준입니다.", "rank"))

    def test_a_result_is_a_cause(self):
        self.assertIn("cause", self.problems("내국인은 6번 연속 낮은 결과입니다.", "run") + self.problems("내국인은 6.8% 감소한 결과입니다.", "year"))
        self.assertIn("cause", self.problems("두 지표의 상관관계는 확인되지 않습니다.", "year"))

    def test_the_applicants_are_not_a_cause(self):
        import insight
        self.assertIsNone(insight.CAUSE.search("2026년 지원인원은 554,174명입니다."))
        self.assertIsNotNone(insight.CAUSE.search("감소의 원인은 계절입니다."))

    def test_a_comparison_period_is_not_now(self):
        self.assertIn("period", self.problems("내국인은 2025년 8월에 206으로 가장 높았습니다.", "year"))
        self.assertNotIn("period", self.problems("내국인은 2025년 8월 206 대비 6.8% 낮습니다.", "year"))
        self.assertNotIn("period", self.problems("내국인의 최고는 309(2024년 5월)입니다.", "high"))

    def test_figures_from_two_facts_do_not_share_a_clause(self):
        pack = (facts.Fact("a", "paid dividends returning $64 million", ("64",)), facts.Fact("b", "8,019 estates were finalised", ("8,019",)))
        import insight
        self.assertIn("join", insight.check({"statement": "Dividends of $64 million across 8,019 estates.", "fact_ids": ["a", "b"]}, pack, "en"))
        self.assertNotIn("join", insight.check({"statement": "Dividends returned $64 million, while 8,019 estates were finalised.", "fact_ids": ["a", "b"]}, pack, "en"))

    def test_a_run_against_last_year_is_not_a_run_of_falls(self):
        self.assertIn("run", self.problems("내국인은 6번 연속 하락했습니다.", "read_year_run"))
        self.assertNotIn("run", self.problems("내국인은 전년 같은 달보다 6번 연속 낮습니다.", "read_year_run"))

    def test_a_right_figure_with_a_forgotten_citation_is_kept(self):
        self.assertEqual(self.problems("2026년 8월 내국인은 전년 같은 달보다 6.8% 낮고, 최고는 309입니다.", "year"), [])

    def test_a_figure_two_facts_hold_cannot_be_credited_to_either(self):
        pack = (facts.Fact("a", "평균 500", ("500",)), facts.Fact("b", "최고 500", ("500",)), facts.Fact("c", "최저 10", ("10",)))
        import insight
        self.assertIn("number", insight.check({"statement": "값은 500입니다.", "fact_ids": ["c"]}, pack, "ko"))

    def test_a_run_of_real_rises_can_be_called_one(self):
        pack = (facts.Fact("read_year_run", "2년 연속 증가했습니다.", ("2",)),)
        import insight
        self.assertEqual(insight.check({"statement": "2년 연속 증가했습니다.", "fact_ids": ["read_year_run"]}, pack, "ko"), [])

    def test_a_rank_over_a_stretch_nobody_calculated_fails(self):
        self.assertIn("window", self.problems("내국인은 최근 6개월 중 5번째입니다.", "run", "rank"))
        self.assertNotIn("window", self.problems("내국인은 최근 21개 값 중 19번째입니다.", "rank"))

    def test_last_years_value_is_not_this_years(self):
        import insight
        pack = (facts.Fact("subject", "", (), "합계", "Q2 2018"), facts.Fact("value", "Q2 2018 합계: 543,694", ("543,694",)),
                facts.Fact("read_year", "+5.5% against the same quarter last year, Q2 2017, 515,260.", ("5.5", "515,260")))
        wrong = "Q2 2018 was 515,260, which is +5.5% above a year earlier."
        self.assertIn("old_value", insight.check({"statement": wrong, "fact_ids": ["read_year"]}, pack, "en"))
        right = "Q2 2018 was 543,694, which is +5.5% above a year earlier."
        self.assertEqual(insight.check({"statement": right, "fact_ids": ["read_year", "value"]}, pack, "en"), [])
        compared = "Q2 2018 was 5.5% above the 515,260 of a year earlier."
        self.assertNotIn("old_value", insight.check({"statement": compared, "fact_ids": ["read_year"]}, pack, "en"))

    def test_the_wrong_way_round_fails(self):
        import insight
        up = (facts.Fact("read_year", "+6.8% against the same month last year, Aug 2025, 420,651.", ("6.8", "420,651")),)
        check = lambda statement: insight.check({"statement": statement, "fact_ids": ["read_year"]}, up, "en")  # noqa: E731
        self.assertIn("direction", check("It was +6.8% below the same month a year earlier."))
        self.assertIn("direction", check("It was 420,651% below the same month a year earlier."))
        self.assertNotIn("direction", check("It was 6.8% above the same month a year earlier."))

    def test_a_share_is_not_a_fall(self):
        import insight
        pack = (facts.Fact("read_top", "The largest is Short-term (2,236,830), 29.8% of the total.", ("2,236,830", "29.8")),)
        self.assertIn("direction", insight.check({"statement": "The total was 29.8% below a year earlier.", "fact_ids": ["read_top"]}, pack, "en"))
        self.assertNotIn("direction", insight.check({"statement": "Short-term is 29.8% of the total.", "fact_ids": ["read_top"]}, pack, "en"))

    def test_the_largest_is_not_a_multiple_of_the_smaller(self):
        import insight
        pack = (facts.Fact("item_1", "1위 보건복지부 137,494,919, 18.9%", ("137,494,919", "18.9")), facts.Fact("read_gap", "2위 교육부(106,360,749, 14.6%)의 1.3배입니다.", ("106,360,749", "14.6", "1.3")))
        self.assertIn("multiple", insight.check({"statement": "교육부의 본예산은 보건복지부의 1.3배입니다.", "fact_ids": ["read_gap"]}, pack, "ko"))
        self.assertNotIn("multiple", insight.check({"statement": "보건복지부의 본예산은 교육부의 1.3배입니다.", "fact_ids": ["read_gap"]}, pack, "ko"))

    def test_yearly_figures_have_no_same_month(self):
        pack = (facts.Fact("subject", "", (), "이용환자수", "2024년"), facts.Fact("read_prev_year", "직전 해 2023년 1,711,910 대비 +14.5%입니다.", ("1,711,910", "14.5")))
        self.assertIn("period_word", self.problems("2024년 이용환자수는 전년 같은 달보다 14.5% 증가했습니다.", "read_prev_year", pack=pack))
        self.assertNotIn("period_word", self.problems("2024년 이용환자수는 직전 해보다 14.5% 증가했습니다.", "read_prev_year", pack=pack))

    PARTIAL_KO = (facts.Fact("subject", "", (), "외래연인원", "2022년 10월까지"),
                  facts.Fact("read_partial_year", "2022년 10월까지 값은 10개월치로 보입니다. 12개월치인 2021년 879,753 대비 -20.1%는 기간이 달라 그대로 읽을 수 없고, 월평균으로 맞추면 -4.2%입니다.", ("10", "12", "879,753", "20.1", "4.2")),
                  facts.Fact("per_month", "월평균 70,254, 2021년 월평균 73,313 대비 -4.2%", ("70,254", "73,313", "4.2")))

    def test_a_comparison_the_facts_do_not_make_is_refused(self):
        for sentence in ("외래연인원 월평균 70,254는 전년 같은 기간 대비 4.2% 감소했습니다.", "외래연인원 월평균은 전년 같은 달보다 4.2% 감소했습니다."):
            self.assertIn("period_word", self.problems(sentence, "per_month", pack=self.PARTIAL_KO), sentence)
        self.assertNotIn("period_word", self.problems("2022년 10월까지 외래연인원 월평균 70,254는 2021년 월평균 73,313 대비 4.2% 감소했습니다.", "per_month", pack=self.PARTIAL_KO))

    def test_last_week_is_not_a_year_ago(self):
        import insight
        pack = (facts.Fact("subject", "", (), "Douglas WTP (ML)", "Nov 1 (Sun)"), facts.Fact("week", "-6.4% against the same weekday last week, Oct 25 (Sun), 120.79", ("6.4", "120.79")), facts.Fact("avg7", "-17.0% against the mean of the previous 7, 136.19", ("17.0", "136.19")))
        said = lambda text, *ids: insight.check({"statement": text, "fact_ids": list(ids)}, pack, "en")
        self.assertIn("period_word", said("Douglas WTP (ML) in Nov 1 was 6.4% below the same weekday a year earlier.", "week"))
        self.assertNotIn("period_word", said("Douglas WTP (ML) in Nov 1 was 6.4% below the same weekday last week.", "week"))
        self.assertIn("period_word", said("Douglas WTP (ML) in Nov 1 was 17.0% below last week.", "avg7"))
        self.assertNotIn("period_word", said("Douglas WTP (ML) in Nov 1 was 17.0% below the mean of the previous 7.", "avg7"))

    def test_the_month_before_is_not_last_year(self):
        pack = (facts.Fact("subject", "", (), "매출", "2026년 8월"), facts.Fact("prev", "직전 값 2026년 7월 100 대비 -3.0%", ("100", "3.0")))
        self.assertIn("period_word", self.problems("2026년 8월 매출은 전년보다 3.0% 낮습니다.", "prev", pack=pack))
        self.assertNotIn("period_word", self.problems("2026년 8월 매출은 직전 값보다 3.0% 낮습니다.", "prev", pack=pack))

    def test_a_year_ago_said_either_way_is_the_same_comparison(self):
        import insight
        pack = (facts.Fact("subject", "", (), "Sales", "Aug 2026"), facts.Fact("year", "-6.8% against the same month last year, Aug 2025, 100", ("6.8", "100")))
        self.assertNotIn("period_word", insight.check({"statement": "Sales in Aug 2026 was 6.8% below the same month a year earlier.", "fact_ids": ["year"]}, pack, "en"))

    def test_a_change_between_different_lengths_is_not_said_as_a_plain_fall(self):
        self.assertIn("period_word", self.problems("외래연인원은 2021년 879,753 대비 20.1% 감소했습니다.", "read_partial_year", pack=self.PARTIAL_KO))
        self.assertNotIn("period_word", self.problems("외래연인원 -20.1%는 10개월치와 12개월치의 비교라 기간이 달라 그대로 읽을 수 없습니다.", "read_partial_year", pack=self.PARTIAL_KO))
        self.assertNotIn("period_word", self.problems("외래연인원은 월평균으로 맞추면 4.2% 감소입니다.", "read_partial_year", pack=self.PARTIAL_KO))
        english = (facts.Fact("subject", "", (), "Visits", "Oct 2022 year to date"), facts.Fact("read_partial_year", "Oct 2022 year to date looks like 10 months, so -20.1% against the 12 months of 2021, 879,753, cannot be read as it stands. Per month it is -4.2%.", ("10", "12", "2021", "879,753", "20.1", "4.2")))
        self.assertIn("period_word", self.problems("Visits fell 20.1% against 2021.", "read_partial_year", pack=english))

    def test_part_of_a_year_is_not_called_one_month(self):
        self.assertIn("period_word", self.problems("2022년 10월 외래연인원 월평균은 70,254입니다.", "per_month", pack=self.PARTIAL_KO))
        self.assertNotIn("period_word", self.problems("2022년 10월까지 외래연인원 월평균은 70,254입니다.", "per_month", pack=self.PARTIAL_KO))

    def test_no_current_period_is_claimed_when_the_list_has_none(self):
        pack = (facts.Fact("subject", "", (), "직원수", ""), facts.Fact("item_1", "1위 가 100", ("100",)))
        self.assertIn("period_word", self.problems("이번 기간 직원수는 가가 100으로 가장 많습니다.", "item_1", pack=pack))
        self.assertNotIn("period_word", self.problems("직원수는 가가 100으로 가장 많습니다.", "item_1", pack=pack))

    ITEMS_KO = (facts.Fact("subject", "", (), "입주 당분기", ""),
                facts.Fact("read_top", "가장 큰 항목은 서울(15,610)이고, 전체 65,352의 23.9%입니다.", ("15,610", "65,352", "23.9")),
                facts.Fact("read_top_three", "항목 40개 중 상위 3개가 전체의 55.9%입니다.", ("40", "3", "55.9")),
                facts.Fact("item_1", "1위 서울 15,610, 23.9%", ("1", "15,610", "23.9")), facts.Fact("item_2", "2위 시화 11,344, 17.4%", ("2", "11,344", "17.4")))

    def test_words_that_judge_are_refused(self):
        for sentence in ("상위 3개의 입주 당분기가 전체의 55.9%를 차지하며 집중도가 높다.", "상위 3개의 입주 당분기가 전체의 55.9%를 차지하며, 집중도가 매우 높다.", "서울과 시화의 입주 당분기가 각각 23.9%, 17.4%로 큰 차이를 보이며, 서울의 비중이 두드러진다."):
            self.assertIn("judged", self.problems(sentence, "read_top_three", "item_1", "item_2", pack=self.ITEMS_KO), sentence)
        for sentence in ("서울과 시화의 입주 당분기 차이가 큽니다.", "상위 3개가 전체의 55.9%를 차지하여 주요 분야를 나타냅니다."):
            self.assertIn("judged", self.problems(sentence, "read_top_three", pack=self.ITEMS_KO), sentence)
        self.assertNotIn("judged", self.problems("서울과 시화의 입주 당분기 차이가 크지 않습니다.", "read_top_three", pack=self.ITEMS_KO))
        self.assertNotIn("judged", self.problems("상위 3개의 입주 당분기가 전체의 55.9%를 차지합니다.", "read_top_three", pack=self.ITEMS_KO))
        self.assertNotIn("judged", self.problems("2024년 입주 당분기는 최근 9개 값 중 가장 낮은 수준입니다.", "read_top_three", pack=self.ITEMS_KO))
        self.assertNotIn("judged", self.problems("입주 당분기는 최근 30개 값 중 9번째로 높은 수준입니다.", "read_top_three", pack=self.ITEMS_KO))
        self.assertIn("judged", self.problems("입주 당분기는 높은 수준입니다.", "read_top_three", pack=self.ITEMS_KO))
        import insight
        pack = (facts.Fact("subject", "", (), "Sales", ""), facts.Fact("read_top", "The largest is A (5), 50.0% of the items' sum, 10.", ("5", "50.0", "10")))
        self.assertIn("judged", insight.check({"statement": "Sales were highly concentrated in A, at 50.0% of the items' sum.", "fact_ids": ["read_top"]}, pack, "en"))
        self.assertIn("judged", insight.check({"statement": "Sales rose sharply.", "fact_ids": ["read_top"]}, pack, "en"))

    def test_a_judging_word_that_is_part_of_a_name_or_a_fact_is_allowed(self):
        pack = (facts.Fact("subject", "", (), "집중치료 환자수", "2026년"), facts.Fact("read_prev_year", "직전 해 2025년 100 대비 +3.0%입니다.", ("100", "3.0")))
        self.assertNotIn("judged", self.problems("2026년 집중치료 환자수는 직전 해 2025년 대비 3.0% 증가했습니다.", "read_prev_year", pack=pack))
        said = (facts.Fact("subject", "", (), "매출", ""), facts.Fact("read_gap", "2위와 차이가 크게 벌어졌습니다.", ()))
        self.assertNotIn("judged", self.problems("매출은 2위와 차이가 크게 벌어졌습니다.", "read_gap", pack=said))

    def test_the_sum_is_not_one_items_value(self):
        import insight
        pack = (facts.Fact("subject", "", (), "employees_covered", ""), facts.Fact("read_top", "The largest is Woolworths Group Limited (165,507), 3.8% of the items' sum, 4,393,656.", ("165,507", "3.8", "4,393,656")),
                facts.Fact("item_1", "number 1 Woolworths Group Limited 165,507, 3.8%", ("1", "165,507", "3.8")), facts.Fact("item_2", "number 2 Coles Supermarkets 117,552, 2.7%", ("2", "117,552", "2.7")))
        said = lambda text, *ids: insight.check({"statement": text, "fact_ids": list(ids)}, pack, "en")
        self.assertIn("item_value", said("employees_covered for Woolworths Group Limited was 4,393,656.", "read_top"))
        self.assertNotIn("item_value", said("employees_covered for Woolworths Group Limited was 165,507.", "read_top"))
        self.assertNotIn("item_value", said("employees_covered for Woolworths Group Limited was 3.8% of the items' sum, 4,393,656.", "read_top"))
        self.assertIn("item_value", said("employees_covered for Woolworths Group Limited was 117,552.", "read_top", "item_2"))
        self.assertNotIn("item_value", said("employees_covered for Coles Supermarkets was 117,552.", "item_2"))

    def test_the_sum_in_korean_and_two_names_in_one_clause(self):
        self.assertIn("item_value", self.problems("서울의 입주 당분기는 65,352입니다.", "read_top", pack=self.ITEMS_KO))
        self.assertNotIn("item_value", self.problems("서울의 입주 당분기는 전체 65,352의 23.9%입니다.", "read_top", pack=self.ITEMS_KO))
        self.assertNotIn("item_value", self.problems("서울의 입주 당분기 15,610은 시화 11,344보다 많습니다.", "item_1", "item_2", pack=self.ITEMS_KO))
        self.assertNotIn("item_value", self.problems("입주 당분기의 합은 65,352입니다.", "read_top", pack=self.ITEMS_KO))

    def test_the_second_is_not_a_multiple_of_the_second(self):
        pack = (facts.Fact("subject", "", (), "Staff", ""), facts.Fact("read_gap", "That is 1.2 times the second, VIC (612, 23.8%).", ("1.2", "612", "23.8")),
                facts.Fact("item_1", "number 1 NSW 683, 26.6%", ("1", "683", "26.6")), facts.Fact("item_2", "number 2 VIC 612, 23.8%", ("2", "612", "23.8")))
        said = lambda text: insight.check({"statement": text, "fact_ids": ["read_gap"]}, pack, "en")
        self.assertIn("multiple", said("Staff in VIC was 1.2 times the second largest value."))
        self.assertIn("multiple", said("Staff in VIC was 612, which is 1.2 times the second largest value."))
        self.assertNotIn("multiple", said("Staff in NSW was 1.2 times the second largest value."))
        self.assertNotIn("multiple", said("Staff in NSW was 1.2 times VIC."))
        self.assertNotIn("multiple", said("Staff was 1.2 times the second."))
        short = tuple(fact for fact in pack if fact.id != "item_2")  # with several files, only the first item is listed
        self.assertIn("multiple", insight.check({"statement": "Staff in VIC was 612, which is 1.2 times the second largest value.", "fact_ids": ["read_gap"]}, short, "en"))

    def test_a_monthly_mean_is_compared_with_a_monthly_mean(self):
        for sentence in ("2022년 10월까지 외래연인원은 2021년 월평균 대비 4.2% 감소했습니다.", "2022년 10월까지 외래연인원은 2021년 월평균 73,313보다 4.2% 낮습니다."):
            self.assertIn("period_word", self.problems(sentence, "per_month", pack=self.PARTIAL_KO), sentence)
        for sentence in ("2022년 10월까지 외래연인원 월평균 70,254는 2021년 월평균 73,313 대비 4.2% 감소했습니다.", "2022년 10월까지 외래연인원은 월평균으로 맞추면 4.2% 감소입니다."):
            self.assertNotIn("period_word", self.problems(sentence, "per_month", pack=self.PARTIAL_KO), sentence)

    def test_a_value_is_not_a_worth(self):
        self.assertIn("meaning", self.problems("상위 3개 항목이 전체 가치의 55.9%를 차지합니다.", "read_top_three", pack=self.ITEMS_KO))

    PEOPLE_KO = (facts.Fact("subject", "", (), "합계", ""), facts.Fact("sheet", "이 수치가 있는 시트의 이름: 1.시·군별 세대 및 인구. 시트 제목: 1. 시·군별 세대 및 인구", ()),
                 facts.Fact("read_top", "가장 큰 항목은 청주시(877,691)이고, 전체 1,658,715의 52.9%입니다.", ("877,691", "1,658,715", "52.9")), facts.Fact("item_1", "1위 청주시 877,691, 52.9%", ("1", "877,691", "52.9")), facts.Fact("item_2", "2위 충주시 213,090, 12.8%", ("2", "213,090", "12.8")))

    def test_a_name_taken_from_a_sheets_title_is_refused(self):
        for sentence in ("청주시의 세대 수는 전체 1,658,715의 52.9%를 차지합니다.", "상위 3개 항목의 세대 수는 전체의 52.9%를 차지합니다.", "청주시의 인구는 전체의 52.9%입니다."):
            self.assertIn("name", self.problems(sentence, "read_top", pack=self.PEOPLE_KO), sentence)
        self.assertNotIn("name", self.problems("청주시는 전체 1,658,715의 52.9%를 차지합니다.", "read_top", pack=self.PEOPLE_KO))
        pack = (facts.Fact("subject", "", (), "입주 당분기", ""), facts.Fact("sheet", "이 수치가 있는 시트의 이름: 표1. 시트 제목: 단지별 입주업체 현황", ()), facts.Fact("item_1", "1위 서울 15,610, 23.9%", ("1", "15,610", "23.9")))
        self.assertIn("name", self.problems("서울 지역의 입주업체 수는 전체의 23.9%입니다.", "item_1", pack=pack))
        self.assertNotIn("name", self.problems("서울의 입주 당분기는 전체의 23.9%입니다.", "item_1", pack=pack))

    def test_a_title_naming_one_thing_can_name_a_total(self):
        pack = (facts.Fact("subject", "", (), "합계", "12월"), facts.Fact("sheet", "이 수치가 있는 시트의 이름: 입장객 현황", ()), facts.Fact("value", "12월 합계: 2,242", ("12", "2,242")))
        self.assertNotIn("name", self.problems("12월 입장객 합계는 2,242입니다.", "value", pack=pack))

    def test_a_measure_no_fact_mentions_is_refused(self):
        for sentence in ("1위 청주시의 방문객 수는 877,691입니다.", "청주시의 생산량은 877,691입니다.", "청주시는 877,691명으로 전체의 52.9%입니다."):
            self.assertIn("name", self.problems(sentence, "read_top", pack=self.PEOPLE_KO), sentence)
        for sentence in ("항목 수는 2개입니다.", "청주시가 전체의 52.9%입니다."):
            self.assertNotIn("name", self.problems(sentence, "read_top", pack=self.PEOPLE_KO), sentence)
        pack = (facts.Fact("subject", "", (), "이용환자수(명)", "2024년"), facts.Fact("value", "2024년 이용환자수(명): 1,960,526", ("2024", "1,960,526")))
        self.assertNotIn("name", self.problems("2024년 이용환자수는 1,960,526명입니다.", "value", pack=pack))

    def test_an_english_name_taken_from_a_title_or_made_up_is_refused(self):
        pack = (facts.Fact("subject", "", (), "Female individuals", ""), facts.Fact("sheet", "The sheet these figures are on is named: data. Its title says: Average taxable income and households", ()),
                facts.Fact("item_1", "number 1 NSW 683, 26.6%", ("1", "683", "26.6")))
        said = lambda text: insight.check({"statement": text, "fact_ids": ["item_1"]}, pack, "en")
        self.assertIn("name", said("NSW had 683 households."))
        self.assertIn("name", said("The number of visitors in NSW was 683."))
        self.assertNotIn("name", said("Female individuals in NSW was 683."))
        self.assertNotIn("name", said("The total of NSW was 683."))
        self.assertNotIn("name", said("The number of items for NSW was 683."))

    def test_the_prompts_say_a_name_comes_from_the_subject_only(self):
        self.assertIn("시트 이름이나 제목에서", insight.PROMPT["ko"])
        self.assertIn("sheet's name or title", insight.PROMPT["en"])

    UNFINISHED_EN = (facts.Fact("subject", "", (), "Items", "2026"), facts.Fact("read_lead", "2026 had not ended when the file was saved on 2025-10-03, so the value may be incomplete.", ("2026", "2025", "10", "03")),
                     facts.Fact("read_prev_year", "-88.0% against the year before, 2025, 8,010.", ("88.0", "2025", "8,010")), facts.Fact("value", "2026 Items: 960", ("2026", "960")),
                     facts.Fact("read_extreme", "This is the lowest of the latest 8 values.", ("8",)), facts.Fact("high", "highest of the latest 8: 960 (2026)", ("8", "960", "2026")))

    def test_a_change_in_an_unfinished_period_needs_saying_so(self):
        said = lambda text, *ids: insight.check({"statement": text, "fact_ids": list(ids)}, self.UNFINISHED_EN, "en")
        self.assertIn("unfinished", said("Items in 2026 was 88.0% below the year before, 2025.", "read_prev_year"))
        self.assertIn("unfinished", said("Items in 2026 is the lowest of the latest 8 values.", "read_extreme"))
        self.assertNotIn("unfinished", said("Items in 2026 was 88.0% below the year before, but 2026 had not ended when the file was saved.", "read_prev_year", "read_lead"))
        self.assertNotIn("unfinished", said("Items in 2026 so far was 960.", "value"))
        self.assertNotIn("unfinished", said("Items in 2026 was 960.", "value"))

    def test_a_change_in_an_unfinished_period_in_korean(self):
        pack = (facts.Fact("subject", "", (), "건수", "2026년"), facts.Fact("read_lead", "2026년은 파일을 저장한 2026-05-10에 아직 끝나지 않은 기간이라 값이 덜 쌓였을 수 있습니다.", ("2026", "2026", "05", "10")),
                facts.Fact("read_prev_year", "직전 해 2025년 3,117 대비 -93.8%입니다.", ("2025", "3,117", "93.8")))
        self.assertIn("unfinished", self.problems("2026년 건수는 직전 해 2025년 대비 93.8% 감소했습니다.", "read_prev_year", pack=pack))
        self.assertNotIn("unfinished", self.problems("2026년은 아직 끝나지 않은 기간이라 값이 덜 쌓였을 수 있으며, 건수는 직전 해 2025년 대비 93.8% 낮습니다.", "read_prev_year", pack=pack))

    def test_other_warnings_ask_for_nothing_of_the_kind(self):
        pack = (facts.Fact("subject", "", (), "건수", "2026년"), facts.Fact("read_lead", "그 뒤 2개 기간은 '집계중'이라 2026년 값을 썼습니다.", ("2", "2026")), facts.Fact("read_prev_year", "직전 해 2025년 3,117 대비 -93.8%입니다.", ("2025", "3,117", "93.8")))
        self.assertNotIn("unfinished", self.problems("2026년 건수는 직전 해 2025년 대비 93.8% 감소했습니다.", "read_prev_year", pack=pack))

    def test_a_calculated_change_says_its_period_is_unfinished(self):
        made = [candidate.statement for candidate in insight.calculated_candidates(self.UNFINISHED_EN, "en")]
        change = next(text for text in made if "88.0" in text)
        self.assertIn("had not ended", change)
        lowest = next(text for text in made if "lowest" in text)
        self.assertIn("had not ended", lowest)

    def test_the_prompts_ask_for_the_warning(self):
        self.assertIn("덜 쌓였다고", insight.PROMPT["ko"])
        self.assertIn("may be incomplete", insight.PROMPT["en"])

    def test_a_sentence_about_named_items_can_leave_out_the_figures_name(self):
        for sentence, cited in (("서울 15,610은 전체 65,352의 23.9%를 차지합니다.", "read_top"), ("상위 3개 항목(서울, 시화, 반월)은 전체의 55.9%를 차지합니다.", "read_top_three")):
            self.assertNotIn("scope", self.problems(sentence, cited, pack=self.ITEMS_KO), sentence)

    def test_with_two_files_the_figures_name_is_still_needed(self):
        pack = tuple(facts.Fact("file1_" + fact.id, fact.text, fact.numbers, fact.figure, fact.period) for fact in self.ITEMS_KO) + (facts.Fact("file2_subject", "", (), "매출", ""),)
        self.assertIn("scope", self.problems("서울은 전체 65,352의 23.9%입니다.", "file1_read_top", pack=pack))

    NOW_KO = (facts.Fact("subject", "이번 기간은 2025년이고 이번 값은 739226입니다.", ("2025", "739226"), "직장보험료(억원)", "2025년"),
              facts.Fact("read_prev_year", "직전 해 2024년 714365 대비 +3.5%입니다.", ("2024", "714365", "3.5")), facts.Fact("value", "2025년 직장보험료(억원): 739226", ("2025", "739226")),
              facts.Fact("high", "최근 5개 중 최고 739226 (2025년)", ("5", "739226", "2025")), facts.Fact("per_month", "월평균 61,602, 2024년 월평균 59,530 대비 +3.5%", ("61,602", "2024", "59,530", "3.5")))

    def test_the_current_value_is_credited_however_many_facts_hold_it(self):
        for sentence in ("2025년 직장보험료(억원)는 739226으로, 직전 해 2024년 714365 대비 3.5% 증가했습니다.", "2025년 직장보험료(억원)는 739226으로 직전 해 2024년 714365 대비 3.5% 증가했습니다."):
            self.assertEqual(self.problems(sentence, "read_prev_year", pack=self.NOW_KO), [], sentence)

    def test_the_current_value_does_not_excuse_a_comparison_of_another_size(self):
        self.assertIn("join", self.problems("2025년 직장보험료(억원) 739226은 2024년 월평균 59,530 대비 3.5% 높습니다.", "per_month", pack=self.NOW_KO))

    def test_the_form_a_model_copies_says_this_figure_and_this_period(self):
        ask = insight.prompt_for(self.NOW_KO, "ko")
        self.assertIn("2025년 직장보험료(억원)은", ask)
        self.assertNotIn("[이번 기간]", ask)
        items = insight.prompt_for(self.ITEMS_KO, "ko")
        self.assertIn("입주 당분기", items)
        self.assertNotIn("이번 기간] ", items)
        english = insight.prompt_for((facts.Fact("subject", "", (), "employees_covered", ""), facts.Fact("item_1", "number 1 A 5", ("1", "5"))), "en")
        self.assertIn("employees_covered of [an item's name]", english)
        self.assertNotIn("[current period]", english)
        two = insight.prompt_for((facts.Fact("file1_subject", "", (), "가", "2025년"), facts.Fact("file2_subject", "", (), "나", "2024년")), "ko")
        self.assertIn("[수치 이름]", two)

    def test_prompts_do_not_teach_a_comparison_the_facts_may_lack(self):
        import insight
        for prompt in insight.PROMPT.values():
            self.assertIsNone(re.search(r"같은 달|same month|a year earlier", prompt))

    def test_a_good_sentence_still_passes(self):
        self.assertEqual(self.problems("2026년 8월 내국인은 전년 같은 달보다 6.8% 낮고, 전년 같은 달보다 6번 연속 낮습니다.", "year", "run"), [])

    def test_calculated_sentences_say_what_the_results_say(self):
        import insight
        pack = (facts.Fact("files", "지시", ()), facts.Fact("cross_direction", "모두 줄었습니다.", ()), facts.Fact("doc1_link1", "[문서] 2005년 7월에 적힌 일: 철수. Ballina 전 51,166, 후 70,112, +37.0%.", ("51,166",)),
                facts.Fact("file1_subject", "[가] 지시", (), "내국인", "2026년 8월"), facts.Fact("file1_read_year_run", "[가] 전년 같은 달보다 6번 연속 낮습니다.", ("6",)))
        made = insight.calculated_candidates(pack, "ko")
        self.assertEqual([c.statement for c in made], ["모두 줄었습니다.", "2005년 7월에 적힌 일: 철수. Ballina 전 51,166, 후 70,112, +37.0%.", "2026년 8월 내국인: 전년 같은 달보다 6번 연속 낮습니다."])
        self.assertTrue(all(c.calculated for c in made))


class InsightProgress(unittest.TestCase):
    def test_each_answer_is_reported_with_what_has_passed_so_far(self):
        import insight
        pack = (facts.Fact("value", "2026년 8월 매출 192", ("192",)),)
        replies = iter(['{"insights": [{"statement": "매출은 192입니다.", "fact_ids": ["value"], "hypothesis": ""}]}', "못 읽는 답", '{"insights": [{"statement": "192로 나타납니다.", "fact_ids": ["value"], "hypothesis": ""}]}'])
        real, insight.post = insight.post, lambda path, body, timeout: {"response": next(replies)}
        seen = []
        try:
            made = insight.make_candidates(pack, "ko", on_run=lambda done, runs, found: seen.append((done, runs, len(found.candidates))))
        finally:
            insight.post = real
        self.assertEqual(seen, [(1, 3, 1), (2, 3, 1), (3, 3, 2)])  # a reply that cannot be read still counts as a round
        self.assertEqual(len(made.candidates), 2)


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

    def test_a_machines_name_for_a_sheet_says_nothing_about_language(self):
        self.assertEqual(facts.detect_language(["tourStatConfirmManageBiz_2021-0", "입장객"]), "ko")
        self.assertEqual(facts.detect_language(["RETAINED_CATCH_KG", "어획량"]), "ko")

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
        self.assertTrue(self.problems(statement("직전 7개 평균 대비 8.6% 감소했습니다.", "avg7")))  # 8.6 is another fact's

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
        self.assertEqual(self.problems(statement("The value fell further today.", "week")), ["language"])
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
        self.assert_fails("wrong_type_excel", excel=self.write("x.doc", core.OLE_MAGIC))
        self.assert_fails("file_unreadable", excel=self.write("x.xls", core.OLE_MAGIC))  # the name is right, and the content cannot be read
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
