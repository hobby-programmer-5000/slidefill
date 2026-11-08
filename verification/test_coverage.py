"""New input/chart/proof coverage via the actual automatic application boundary."""
from fractions import Fraction
from pathlib import Path
import io
import shutil
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile,ZIP_DEFLATED
import openpyxl
from pptx import Presentation
from pptx.chart.data import BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
import core,layouts,suggest
from . import automatic,kernel,oracle,sources
from . import test_automatic as fixture
alter,C=fixture.alter,fixture.C
from .test_verification import rates,staff


def ods(path,extra=''):
    text='''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Data">
<table:table-row><table:table-cell office:value-type="string"><text:p>Item</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Amount</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="12"/></table:table-row>
<table:table-row><table:table-cell office:value-type="string"><text:p>B</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="34"/></table:table-row>'''+extra+'''</table:table></office:spreadsheet></office:body></office:document-content>'''
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('mimetype','application/vnd.oasis.opendocument.spreadsheet')
        z.writestr('content.xml',text)
        z.writestr('META-INF/manifest.xml','<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>')


class Coverage(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app
    blocked_by=fixture.AutomaticGuard.blocked_by

    def test_delimited_encodings_quotes_codes_and_source_corruption(self):
        for ext,codec,sep in [('.csv','utf-8-sig',','),('.tsv','utf-16','\t'),('.txt','cp949',';')]:
            with self.subTest(ext=ext):
                self.book=self.root/('Data'+ext)
                self.book.write_text(f'Item{sep}Amount\n007{sep}12\n서울{sep}34\n',encoding=codec)
                self.make();report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
                self.assertEqual(sources.read(self.book)['Data'].get('A2'),'007')
                original=core.ExcelSource.range_cells
                def broken(source,key):
                    grid,where,issues=original(source,key);grid[2][1]=(999,'999','General');return grid,where,issues
                with patch.object(core.ExcelSource,'range_cells',broken):self.assertTrue(self.run_app().blocked)
        self.book=self.root/'Data.csv';self.book.write_text('Item,Amount\n"a,b","1,234"\nB,34\n',encoding='utf-8')
        self.make();self.assertEqual(self.run_app().verification['state'],'passed')

    def test_csv_numeric_filter_normalization_keeps_exact_arithmetic_source(self):
        self.book=self.root/'Data.csv'
        self.book.write_text('Date,Group,Amount\n2026-01-01,1.00,2\n2026-01-02,1.00,3\n',encoding='utf-8')
        self.make('latest2:Data!A,sum(C, B=1.0)')
        self.assertEqual(self.run_app().verification['state'],'passed')
        sheet=sources.read(self.book)['Data']
        self.assertEqual(sheet.get('B2'),'1.00')
        self.assertEqual(sheet.condition_values['B2'],'1.0')

    def test_renamed_xlsx_and_ods_are_independently_checked(self):
        new=self.root/'renamed.xls';shutil.copy2(self.book,new);self.book=new;self.make()
        self.assertEqual(self.run_app().verification['state'],'passed')
        self.book=self.root/'data.ods';ods(self.book)
        self.assertEqual(self.run_app(True).verification['state'],'passed')
        self.assertTrue(self.out.exists())
        original=core.ExcelSource.range_cells
        def broken(source,key):
            grid,where,issues=original(source,key);grid[1][1]=(100,'100','General');return grid,where,issues
        with patch.object(core.ExcelSource,'range_cells',broken):self.assertTrue(self.run_app().blocked)

    def test_ods_repeats_merges_dates_boolean_and_bounded_expansion(self):
        self.book=self.root/'data.ods'
        ods(self.book,'<table:table-row><table:table-cell office:value-type="date" office:date-value="2026-01-01"/><table:table-cell office:value-type="boolean" office:boolean-value="true"/></table:table-row><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string" table:number-columns-spanned="2"><text:p>merged</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row table:number-rows-repeated="1000000"/>')
        sheet=sources.read(self.book)['Data'];self.assertEqual(sheet.get('B4'),True);self.assertEqual(sheet.types['A4'],'d')
        self.assertEqual(sheet.anchor(6,2),'A6');self.assertLess(len(sheet.cells),20)
        ods(self.book,'<table:table-row table:number-rows-repeated="3000000"><table:table-cell office:value-type="float" office:value="1"/></table:table-row>')
        with self.assertRaises(sources.Unsupported):sources.read(self.book)

    def test_unsupported_ods_value_is_partial_and_known_output_damage_still_blocks(self):
        self.book=self.root/'data.ods';ods(self.book,'<table:table-row><table:table-cell office:value-type="time" office:time-value="P1Y"/></table:table-row>')
        self.make();report=self.run_app();self.assertEqual(report.verification['state'],'partial',report.verification)
        def corrupt(raw):
            node=ET.fromstring(raw);node.find('.//'+C+'val/'+C+'numRef/'+C+'numCache/'+C+'pt/'+C+'v').text='999';return ET.tostring(node)
        self.blocked_by(lambda blob:alter(blob,'ppt/charts/chart1.xml',corrupt))

    def bubble(self,chart_type=XL_CHART_TYPE.BUBBLE):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);data=BubbleChartData();data.add_series('old').add_data_point(1,2,3)
        shape=slide.shapes.add_chart(chart_type, Inches(1),Inches(1),Inches(8),Inches(4),data)
        shape._element.xpath('.//p:cNvPr')[0].set('descr','{{chart:Data!A1:E4}}');prs.save(self.template)

    def test_bubble_multiseries_and_gaps_save_with_source_linkage(self):
        self.write([['X','First','Size','Second','Size'],[1,12,4,30,9],[2,None,7,40,16],[3,20,0,50,25]])
        for kind in (XL_CHART_TYPE.BUBBLE,XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT):
            self.bubble(kind);self.assertEqual(self.run_app().verification['state'],'passed')
        result=self.run_app(True);self.assertEqual(result.verification['source_checked'],1)
        prs=Presentation(self.out);self.assertEqual(len(prs.slides[0].shapes[0].chart.series[0].values),2)

    def test_bubble_size_cache_and_embedded_damage_block(self):
        self.write([['X','First','Size','Second','Size'],[1,12,4,30,9],[2,15,7,40,16],[3,20,0,50,25]]);self.bubble()
        def corrupt(raw):
            node=ET.fromstring(raw);node.find('.//'+C+'bubbleSize/'+C+'numRef/'+C+'numCache/'+C+'pt/'+C+'v').text='999';return ET.tostring(node)
        self.blocked_by(lambda blob:alter(blob,'ppt/charts/chart1.xml',corrupt))
        def embedded(blob):
            def changed(raw):
                wb=openpyxl.load_workbook(io.BytesIO(raw));wb.active['C2']=999;stream=io.BytesIO();wb.save(stream);return stream.getvalue()
            with ZipFile(io.BytesIO(blob)) as z:member=next(n for n in z.namelist() if n.startswith('ppt/embeddings/') and n.endswith('.xlsx'))
            return alter(blob,member,changed)
        self.blocked_by(embedded)

    def test_bubble_refuses_ambiguous_columns_and_negative_sizes_without_changing_chart(self):
        self.write([['X','Y','Size','Y2','Size2'],[1,2,-3,4,5],[2,3,4,5,6],[3,4,5,6,7]]);self.bubble()
        report=self.run_app();self.assertEqual(report.items[0].issues[0].code,'chart_bubble_size')
        prs=Presentation(self.template);shape=prs.slides[0].shapes[0];shape._element.xpath('.//p:cNvPr')[0].set('descr','{{chart:Data!A1:B4}}');prs.save(self.template)
        self.assertEqual(self.run_app().items[0].issues[0].code,'chart_bubble_columns')

    def test_empty_bubble_series_is_an_explicit_input_error(self):
        self.write([['X','First','Size','Second','Size'],[1,None,4,30,9],[2,None,7,40,16],[3,None,0,50,25]])
        self.bubble();report=self.run_app()
        self.assertEqual(report.items[0].issues[0].code,'chart_bubble_empty')
        self.assertEqual(report.count('error'),1)

    def test_structured_views_beyond_old_row_column_limits_have_complete_coverage(self):
        for make in (rates,staff):
            make(self.book,dr=350,dc=90);source=core.ExcelSource(str(self.book));sheet=source.wb.active
            views=layouts.discover(sheet);self.assertTrue(views)
            raw=sources.read(self.book)[sheet.title];expected=oracle.expected(raw)
            self.assertEqual(len(views[0].rows),len(expected[views[0].id]))
            plan=suggest.plan_report(source,'en');self.assertIsNotNone(plan)
            suggest.make_report(str(self.template),'Shifted',plan.sections)
            self.assertEqual(self.run_app(source=source).verification['state'],'passed')

    def test_tall_rate_blocks_are_not_truncated_at_row300(self):
        rates(self.book);wb=openpyxl.load_workbook(self.book);ws=wb.active
        for r in range(5,406):
            for c,value in enumerate([(r-5)*10+1,(r-4)*10,2,3,5],1):ws.cell(r,c,value)
        wb.save(self.book);source=core.ExcelSource(str(self.book))
        views=layouts.discover(source.wb.active);self.assertEqual(len(views[0].rows),401)
        expected=oracle.expected(sources.read(self.book)['Rates']);self.assertEqual(len(expected[views[0].id]),401)
        plan=suggest.plan_report(source,'en');self.assertEqual(len(plan.sections),51)
        self.assertIn('~400~401:',plan.sections[-1].table.key)

    def test_rational_cross_product_is_actual_generated_path(self):
        with patch.object(kernel.verified,'RationalWithin',wraps=kernel.verified.RationalWithin) as method:
            self.assertTrue(kernel.equal_rationals(Fraction(-1,3),Fraction(-2,6)))
            method.assert_called_once_with(-1,3,-1,3,0,1)
        for sign in (-1,1):
            a=Fraction(sign*10**100+1,10**50+3);b=a+Fraction(1,101)
            self.assertTrue(kernel.equal_rationals(a,b,Fraction(1,101)))
            self.assertFalse(kernel.equal_rationals(a,b,Fraction(1,102)))
        with self.assertRaises(ValueError):kernel.equal_rationals(0,1,-1)


if __name__=='__main__':unittest.main(verbosity=2)
