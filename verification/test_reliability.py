"""Source-to-final-file regressions for the client-report reliability contract."""
import datetime as dt
import copy
from fractions import Fraction
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from pptx import Presentation
from pptx.util import Inches
import core
import facts
import suggest
import documents
import openpyxl
from . import automatic, test_automatic as fixture, test_formulas as formula_fixture


class ExactReadings(unittest.TestCase):
    def test_integer_text_is_not_rounded_through_binary_float(self):
        for number in (2**53+1,10**25+3,-(2**53+1)):
            self.assertEqual(core.as_number(str(number)),number)
            self.assertEqual(core.as_number(f'{number:,}'),number)
        self.assertEqual(core.figure_number('(9,007,199,254,740,993)'),-(2**53+1))

    def test_cancellation_and_percentage_point_ties_follow_declared_math(self):
        self.assertEqual(facts.exact_total([1e16,1,-1e16]),1)
        self.assertEqual(facts.exact_total([0.1,0.2]),Fraction(3,10))
        self.assertEqual(facts.ratio(0.3,0.2),0.5)
        self.assertEqual(facts.points_text(0.25,'en'),'+0.3 points')
        self.assertEqual(facts.points_text(-0.25,'en'),'-0.3 points')
        self.assertEqual(facts.points_text(Fraction(1,40),'ko',100),'+2.5포인트')

    def test_nonfinite_values_are_not_figures(self):
        for number in (float('nan'),float('inf'),-float('inf')):
            self.assertIsNone(core.as_number(number))

    def test_movement_words_agree_with_percentage_rounding_at_half_ties(self):
        for lang in ('en','ko'):
            text=facts._movement(Fraction(1,2000),'previous',lang)
            self.assertIn('0.1%',text)
            self.assertNotIn('unchanged' if lang=='en' else '변동 없음',text)

    def test_fraction_formatting_keeps_decimal_half_ties_and_optional_zeros(self):
        from numfmt import format_value
        cases=[(Fraction(3,10),'0.##','0.3'),(Fraction(1,3),'0.000','0.333'),
               (Fraction(535,200),'0.00','2.68'),(Fraction(-535,200),'0.00','-2.68'),
               (Fraction(20000000000000001,20),'0.00','1000000000000000.05')]
        for value,fmt,wanted in cases:self.assertEqual(format_value(value,fmt),(wanted,True))

    def test_document_means_are_exact_and_nonpositive_bases_are_not_relabelled_zero(self):
        points=[facts.Point(dt.datetime(2020,1,1)+dt.timedelta(days=90*i),value)
                for i,value in enumerate([1e16,1,-1e16,3,999,2,2,2,2,2])]
        item=documents.Item('Route A','Data',points)
        found=documents.link(documents.Event(points[4].date,'Route A changed'),item)
        self.assertEqual((found.before,found.after),(1,2))
        negative=documents.Link(found.event,item,-100,-50,4)
        self.assertEqual(documents.change_text(negative),'-100 → -50')
        self.assertEqual(documents.mean_text(Fraction(5,2)),'3')

    def test_mixed_document_windows_are_shown_per_row(self):
        item=documents.Item('Route A','Data',[])
        links=[documents.Link(documents.Event(dt.datetime(2024,1,1),'A'),item,100,200,4),
               documents.Link(documents.Event(dt.datetime(2025,1,1),'B'),item,50,75,12)]
        rows=documents.table_rows(links,6,'en')
        self.assertEqual(rows[0][3][1],'Mean before (periods)')
        self.assertEqual(rows[1][3][1],'100 (4)')
        self.assertEqual(rows[2][3][1],'50 (12)')


class ReportReliability(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app
    formula=formula_fixture.FormulaApplication.formula
    blocked_by=fixture.AutomaticGuard.blocked_by

    def text(self,key='Data!B2'):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1),Inches(1),Inches(8),Inches(1)).text='Value: {{'+key+'}}'
        prs.save(self.template)

    def table(self,key):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6])
        shape=slide.shapes.add_table(2,2,Inches(1),Inches(1),Inches(8),Inches(3))
        shape._element.xpath('.//p:cNvPr')[0].set('descr','{{table:'+key+'}}')
        prs.save(self.template)

    def test_numeric_calculation_text_is_independently_recomputed(self):
        self.write([['Date','Amount']]+[[dt.datetime(2025,1,1)+dt.timedelta(days=i),10*(i+1)] for i in range(8)])
        original=core.ExcelSource.text_value
        for name in ('prev','week','avg7','peak','low','전일대비','평균대비7'):
            with self.subTest(name=name):
                self.text(name+':Data!A,B')
                clean=self.run_app()
                self.assertEqual(clean.verification['state'],'passed',clean.verification)
                def bad(source,key):
                    text,where,issues=original(source,key)
                    return '999',where,issues
                with patch.object(core.ExcelSource,'text_value',bad):
                    preview=self.run_app()
                    self.assertTrue(preview.blocked,preview.verification)
                    with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_aggregate_table_preserves_small_amount_next_to_large_amount(self):
        day=dt.datetime(2025,1,1)
        self.write([['Date','Amount'],[day,10**15],[day,0.05]])
        wb=openpyxl.load_workbook(self.book)
        for row in (2,3):wb['Data'].cell(row,2).number_format='0.00'
        wb.save(self.book)
        key='latest:Data!A,합계(B)'
        self.table(key);report=self.run_app(True)
        self.assertEqual(report.verification['state'],'passed',report.verification)
        prs=Presentation(self.out)
        self.assertEqual(prs.slides[0].shapes[0].table.cell(1,1).text,'1000000000000000.05')
        self.out.unlink();self.make(key)
        preview=self.run_app()
        self.assertTrue(preview.blocked,preview.verification)
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_chart_cannot_round_an_exact_integer_to_a_nearby_integer(self):
        self.write([['Item','Amount'],['A',str(2**53+1)],['B',34]])
        self.make()
        preview=self.run_app()
        self.assertTrue(preview.blocked,preview.verification)
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_aggregate_chart_keeps_the_declared_decimal_format(self):
        self.write([['Date','Amount'],[dt.datetime(2025,1,1),0.1],[dt.datetime(2025,1,1),0.2]])
        wb=openpyxl.load_workbook(self.book)
        for row in (2,3):wb['Data'].cell(row,2).number_format='0.00'
        wb.save(self.book);self.make('latest:Data!A,sum(B)')
        report=self.run_app(True)
        self.assertEqual(report.verification['state'],'passed',report.verification)
        with ZipFile(self.out) as z:
            root=ET.fromstring(z.read('ppt/charts/chart1.xml'))
            self.assertEqual(root.findtext('.//'+fixture.C+'numCache/'+fixture.C+'formatCode'),'0.00')
        self.out.unlink()
        self.blocked_by(lambda blob:fixture.alter(blob,'ppt/charts/chart1.xml',lambda data:data.replace(b'>0.00<',b'>General<')))

    def test_text_only_value_is_source_checked_and_saved_text_is_guarded(self):
        self.text();report=self.run_app(True)
        self.assertEqual(report.verification['state'],'passed',report.verification)
        self.assertEqual(report.verification['checked'],0)
        self.assertEqual(report.verification['text_checked'],1)
        self.assertEqual(report.verification['text_source_checked'],1)
        self.out.unlink()
        self.blocked_by(lambda blob:fixture.alter(blob,'ppt/slides/slide1.xml',lambda data:data.replace(b'Value: 12',b'Value: 999')))

    def test_bad_resolver_text_is_not_accepted_as_source_truth(self):
        self.text();original=core.ExcelSource.text_value
        def bad(source,key):
            text,where,issues=original(source,key)
            return '999',where,issues
        with patch.object(core.ExcelSource,'text_value',bad):
            self.assertTrue(self.run_app().blocked)
            with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_text_only_stale_formula_cache_is_blocked(self):
        self.formula('B3*2',999);self.text()
        self.assertTrue(self.run_app().blocked)
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_stale_unselected_date_cannot_hide_the_actual_latest_period(self):
        self.write([['Date','Amount'],[dt.datetime(2025,1,1),12],[dt.datetime(2026,1,1),34]])
        self.formula('A2+365',1,cell_ref='A3')
        self.text('latest:Data!A,B')
        self.assertTrue(self.run_app().blocked)
        with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_incomplete_formula_coverage_cannot_be_saved(self):
        self.formula('UNKNOWN(B3)',12);self.text()
        self.assertTrue(self.run_app().incomplete)
        with self.assertRaises(core.SlideFillError) as error:self.run_app(True)
        self.assertEqual(error.exception.code,'verification_incomplete')
        self.assertFalse(self.out.exists())

    def test_static_template_does_not_overwrite_existing_output(self):
        prs=Presentation();prs.slides.add_slide(prs.slide_layouts[6]);prs.save(self.template)
        self.out.write_bytes(b'previous output')
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertEqual(self.out.read_bytes(),b'previous output')

    def test_document_change_after_preview_blocks_even_workbook_only_fill(self):
        doc=self.root/'context.docx'
        with ZipFile(doc,'w') as z:z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>')
        source=core.Sources([str(self.book),str(doc)]);self.text()
        preview=core.run(str(self.template),[str(self.book),str(doc)],source=source)
        self.assertEqual(preview.verification['state'],'passed')
        with doc.open('ab') as stream:stream.write(b'changed')
        with self.assertRaises(core.SlideFillError):self.run_app(True,source)
        self.assertFalse(self.out.exists())

    def test_rate_mean_requires_explicit_scope_review(self):
        rows=[['Date','Completion rate','Completed','Eligible']]
        for month in range(1,7):rows += [[dt.datetime(2025,month,1),1,100,100],[dt.datetime(2025,month,1),0,0,900]]
        self.write(rows)
        plan=suggest.plan_report(core.ExcelSource(str(self.book)),'en')
        self.assertIn('rate_mean',[code for code,_params in plan.doubts])

    def test_excess_source_decimal_precision_is_not_silently_passed(self):
        def change(data):
            root=ET.fromstring(data)
            cell=next(n for n in root.findall('.//'+fixture.S+'c') if n.get('r')=='B2')
            cell.find(fixture.S+'v').text='0.123456789012345678901'
            return ET.tostring(root,encoding='utf-8')
        self.book.write_bytes(fixture.alter(self.book.read_bytes(),'xl/worksheets/sheet1.xml',change))
        self.make();report=self.run_app()
        self.assertTrue(report.incomplete,report.verification)
        with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_duplicate_source_cells_block_instead_of_accepting_the_last_value(self):
        def change(data):
            root=ET.fromstring(data)
            row=next(n for n in root.findall('.//'+fixture.S+'row') if n.get('r')=='2')
            row.append(copy.deepcopy(row[-1]))
            return ET.tostring(root,encoding='utf-8')
        self.book.write_bytes(fixture.alter(self.book.read_bytes(),'xl/worksheets/sheet1.xml',change))
        self.make();self.assertTrue(self.run_app().blocked)
        with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_document_mean_corruption_is_independently_detected(self):
        rows=[['Date','Route A']]+[[dt.datetime(2020+i//4,(i%4)*3+1,1),100 if i<10 else 200] for i in range(20)]
        self.write(rows);doc=self.root/'context.docx'
        with ZipFile(doc,'w') as z:
            z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>2022-07-01 Route A changed.</w:t></w:r></w:p></w:body></w:document>')
        source=core.Sources([str(self.book),str(doc)])
        plan=suggest.plan_files(source,'en');suggest.make_report(str(self.template),'Report',plan.sections)
        clean=core.run(str(self.template),[str(self.book),str(doc)],source=source)
        self.assertEqual(clean.verification['state'],'passed',clean.verification)
        original=documents.table_rows
        def damaged(*args):
            values=original(*args)
            if len(values)>1:values[1][3]=('999','999','General')
            return values
        with patch.object(documents,'table_rows',side_effect=damaged):
            preview=core.run(str(self.template),[str(self.book),str(doc)],source=source)
            self.assertTrue(preview.blocked,preview.verification)
            with self.assertRaises(core.SlideFillError):
                core.run(str(self.template),[str(self.book),str(doc)],str(self.out),source=source)


if __name__=='__main__':unittest.main()
