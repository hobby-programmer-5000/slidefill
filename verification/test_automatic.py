"""Automatic app-path regressions. Tests never require a user to run a checker."""
import datetime as dt
from decimal import Decimal
from fractions import Fraction
import io
import itertools
import json
from pathlib import Path
import random
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile,ZIP_DEFLATED
import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData,XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
import core
import suggest
from . import automatic,kernel
from .test_verification import rates

C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def alter(blob,member,transform):
    target=io.BytesIO()
    with ZipFile(io.BytesIO(blob)) as src,ZipFile(target,'w',ZIP_DEFLATED) as dst:
        for name in src.namelist():dst.writestr(name,transform(src.read(name)) if name==member else src.read(name))
    return target.getvalue()


class AutomaticGuard(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.book=self.root/'data.xlsx';self.template=self.root/'template.pptx';self.out=self.root/'output.pptx'
        self.write([['Item','Amount'],['A',12],['B',34]])

    def tearDown(self):self.temp.cleanup()

    def write(self,rows):
        wb=openpyxl.Workbook();wb.active.title='Data'
        for row in rows:wb.active.append(row)
        wb.save(self.book)

    def make(self,key='Data!A1:B3',xy=False):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6])
        data=XyChartData() if xy else CategoryChartData()
        if xy:data.add_series('Old').add_data_point(1,2)
        else:data.categories=['Old'];data.add_series('Old',[1])
        shape=slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER if xy else XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(8), Inches(4),data)
        shape._element.xpath('.//p:cNvPr')[0].set('descr','{{chart:'+key+'}}')
        prs.save(self.template)

    def run_app(self,output=False,source=None):
        return core.run(str(self.template),str(self.book),str(self.out) if output else None,source=source)

    def blocked_by(self,transform):
        original=automatic.serialize
        with patch.object(automatic,'serialize',side_effect=lambda prs:transform(original(prs))):
            preview=self.run_app();self.assertTrue(preview.blocked,preview.verification)
            with self.assertRaises(core.SlideFillError) as caught:self.run_app(True)
            self.assertEqual(caught.exception.code,'verification_blocked')
        self.assertFalse(self.out.exists())

    def test_app_automatically_verifies_preview_and_saved_chart(self):
        self.make();before=self.book.read_bytes()
        for output in (False,True):
            report=self.run_app(output)
            self.assertEqual(report.verification['state'],'passed',report.verification)
            self.assertEqual(report.verification['source_checked'],1)
        self.assertEqual(before,self.book.read_bytes());self.assertTrue(self.out.exists())

    def test_source_to_resolved_grid_is_independent(self):
        self.make();original=core.ExcelSource.range_cells
        def bad(source,key):
            rows,where,issues=original(source,key);rows[1][1]=(999,'999','General');return rows,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):
            preview=self.run_app();self.assertTrue(preview.blocked)
            self.assertTrue(any('source' in e for _,e in preview.verification['failures']))

    def test_partial_expression_never_hides_a_definite_mismatch(self):
        self.write([['Item','Amount'],['A','12 apples'],['B',34]])
        self.make();original=core.ExcelSource.range_cells
        def bad(source,key):
            rows,where,issues=original(source,key);rows[2][1]=(999,'999','General');return rows,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):self.assertTrue(self.run_app().blocked)

    def test_changed_source_category_or_measure_name_is_blocked(self):
        self.make();original=core.ExcelSource.range_cells
        for row,col in [(0,1),(1,0)]:
            def bad(source,key):
                rows,where,issues=original(source,key);rows[row][col]=('Wrong','Wrong','General');return rows,where,issues
            with patch.object(core.ExcelSource,'range_cells',bad):self.assertTrue(self.run_app().blocked)

    def test_ranked_category_name_stays_linked_to_source(self):
        self.make('top2:Data!A,B');original=core.ExcelSource.range_cells
        def bad(source,key):
            rows,where,issues=original(source,key);rows[1][0]=('Wrong','Wrong','General');return rows,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):self.assertTrue(self.run_app().blocked)

    def test_repeated_institution_labels_use_the_actual_source_campus(self):
        self.write([['Institution','Campus','Amount'],['University','North',10],['University','South',30],['Other','Main',25]])
        self.make('top3:Data!A,C')
        report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
        original=core.ExcelSource.range_cells
        def bad(source,key):
            rows,where,issues=original(source,key);rows[1][0]=('University (Wrong)','University (Wrong)','General');return rows,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):self.assertTrue(self.run_app().blocked)

    def test_structured_column_omission_cannot_pass_as_complete(self):
        rates(self.book);source=core.ExcelSource(str(self.book));plan=suggest.plan_report(source,'en')
        for section in plan.sections:section.table.columns-=1
        suggest.make_report(str(self.template),'Rates',plan.sections);original=core.ExcelSource.range_cells
        def bad(source,key):
            rows,where,issues=original(source,key);return [line[:-1] for line in rows],where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):self.assertTrue(self.run_app(source=source).blocked)

    def test_cache_point_count_values_and_series_names_are_guarded(self):
        self.make()
        def mutation(field):
            def apply(raw):
                root=ET.fromstring(raw)
                if field=='value':root.find('.//'+C+'val/.//'+C+'pt/'+C+'v').text='999'
                elif field=='count':root.find('.//'+C+'val/.//'+C+'ptCount').set('val','99')
                elif field=='missing':
                    cache=root.find('.//'+C+'val/.//'+C+'numCache');cache.remove(cache.findall(C+'pt')[0])
                elif field=='name':root.find('.//'+C+'tx/.//'+C+'pt/'+C+'v').text='Wrong measure'
                return ET.tostring(root)
            return lambda blob:alter(blob,'ppt/charts/chart1.xml',apply)
        for field in ['value','count','missing','name']:
            with self.subTest(field=field):self.blocked_by(mutation(field))

    def test_large_decimal_error_is_not_hidden_by_relative_tolerance(self):
        self.write([['Item','Amount'],['A',1000000000000.25],['B',0.000000000000000001]])
        self.make()
        def bad(raw):
            root=ET.fromstring(raw);root.find('.//'+C+'val/.//'+C+'pt/'+C+'v').text='1000000000000.5';return ET.tostring(root)
        self.blocked_by(lambda blob:alter(blob,'ppt/charts/chart1.xml',bad))

    def test_embedded_workbook_mutation_is_guarded(self):
        self.make()
        def transform(blob):
            with ZipFile(io.BytesIO(blob)) as z:member=next(n for n in z.namelist() if n.startswith('ppt/embeddings/'))
            def embedded(raw):
                def sheet(xml):
                    root=ET.fromstring(xml);root.find('.//'+S+'c[@r="B2"]/'+S+'v').text='999';return ET.tostring(root)
                return alter(raw,'xl/worksheets/sheet1.xml',sheet)
            return alter(blob,member,embedded)
        self.blocked_by(transform)

    def test_chart_binding_relationship_cannot_be_switched(self):
        self.make()
        def rels(raw):return raw.replace(b'../charts/chart1.xml',b'../charts/missing.xml')
        self.blocked_by(lambda blob:alter(blob,'ppt/slides/_rels/slide1.xml.rels',rels))

    def test_source_change_after_preview_or_during_save_blocks(self):
        self.make();source=core.ExcelSource(str(self.book))
        self.write([['Item','Amount'],['A',99],['B',34]])
        self.assertTrue(self.run_app(source=source).blocked)
        with self.assertRaises(core.SlideFillError):self.run_app(True,source)
        self.assertFalse(self.out.exists())

    def test_final_candidate_tampering_is_not_published(self):
        self.make();original=automatic.Session._verify_candidate
        def bad(session,path):
            original(session,path)
            with open(path,'ab') as f:f.write(b'changed after check')
        with patch.object(automatic.Session,'_verify_candidate',bad):
            with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists());self.assertEqual(list(self.root.glob('.SlideFill-checked-*')),[])

    def test_source_change_after_candidate_check_is_not_published(self):
        self.make();original=automatic.Session._verify_candidate
        def bad(session,path):
            original(session,path);self.write([['Item','Amount'],['A',99],['B',34]])
        with patch.object(automatic.Session,'_verify_candidate',bad):
            with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_existing_output_is_preserved(self):
        self.make();self.out.write_bytes(b'KEEP')
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertEqual(self.out.read_bytes(),b'KEEP')

    def test_numeric_text_negative_missing_and_scatter_points(self):
        self.write([['Item','Amount'],['A','1,234'],['B','(500)'],['C',None],['D','미집계']]);self.make('Data!A1:B5')
        report=self.run_app(True);self.assertEqual(report.verification['state'],'passed',report.verification)
        self.out.unlink()
        self.write([['X','Y'],[1,2],[2,None],[3,'1,234']]);self.make('Data!A1:B4',xy=True)
        report=self.run_app(True);self.assertEqual(report.verification['state'],'passed',report.verification)

    def test_latest_aggregates_and_ranked_groups_are_source_checked(self):
        rows=[['Day','Group','Amount'],['2024-01-01','A',10],['2024-01-01','B',20],['2024-01-02','A',5],['2024-01-02','B',7]]
        for key in ['latest2:Data!A,sum(C)','latest2:Data!A,mean(C)','top2:Data!B,sum(C)']:
            self.write(rows);self.make(key)
            report=self.run_app()
            self.assertEqual(report.verification['state'],'passed',(key,report.verification))

    def test_native_datetime_and_1904_epoch_are_source_checked(self):
        self.write([['Day','Amount'],[dt.datetime(2024,1,1),10],[dt.datetime(2024,1,2),20]])
        self.make('latest2:Data!A,B')
        report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)

    def test_split_dates_and_share_denominators_are_source_checked(self):
        self.write([['Year','Month','Amount'],[2024,1,10],[2024,2,20]])
        self.make('latest2:Data!A+B,C')
        report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
        for rows in [[['Item','Amount'],['A',10],['B',30]], [['Item','Amount'],['A',10],['B',30],['Total',40]]]:
            self.write(rows);self.make('share2:Data!A,B')
            report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)

    def test_date_axis_and_multiple_series_are_checked(self):
        self.write([['Day','First','Second'],[dt.datetime(2024,1,1),10,20],[dt.datetime(2024,1,2),30,40]])
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);data=CategoryChartData()
        data.categories=[dt.date(2020,1,1),dt.date(2020,1,2)];data.add_series('Old',[1,2])
        shape=slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(1), Inches(1), Inches(8), Inches(4),data)
        self.assertTrue(shape.chart._chartSpace.xpath('.//c:dateAx'))
        shape._element.xpath('.//p:cNvPr')[0].set('descr','{{chart:Data!A1:C3}}');prs.save(self.template)
        report=self.run_app(True);self.assertEqual(report.verification['state'],'passed',report.verification)
        wb=openpyxl.load_workbook(self.book);wb.epoch=dt.datetime(1904,1,1);wb.save(self.book)
        report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)

    def test_structured_special_layout_verifies_without_cli(self):
        rates(self.book,broken=True);source=core.ExcelSource(str(self.book));plan=suggest.plan_report(source,'ko')
        suggest.make_report(str(self.template),'Rates',plan.sections)
        report=self.run_app(True,source)
        self.assertEqual(report.verification['state'],'passed',report.verification)
        self.assertEqual(report.verification['source_checked'],len(plan.sections))

    def test_csv_source_is_independently_checked(self):
        self.book=self.root/'Data.csv';self.book.write_text('Item,Amount\nA,12\nB,34\n',encoding='utf-8')
        self.make();report=self.run_app(True)
        self.assertEqual(report.verification['state'],'passed')
        self.assertEqual(report.verification['source_checked'],1)

    def test_text_only_report_does_not_run_extra_serialization(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);slide.shapes.add_textbox(0,0,Inches(5),Inches(1)).text='{{Data!B2}}';prs.save(self.template)
        with patch.object(automatic,'serialize',side_effect=AssertionError('unnecessary guard serialization')):
            report=self.run_app()
        self.assertEqual(report.verification,{})


class VerifiedBridge(unittest.TestCase):
    def test_proof_receipt_matches_actual_source_and_generated_bytes(self):
        import hashlib
        root=Path(__file__).parent
        receipt=json.loads((root/'proofs/receipt.json').read_text(encoding='utf-8'))
        self.assertEqual(receipt['source_sha256'],hashlib.sha256((root/'proofs/Kernel.dfy').read_bytes()).hexdigest())
        self.assertEqual(receipt['generated_class_sha256'],hashlib.sha256((root/'verified_kernel.py').read_bytes()).hexdigest())
        self.assertEqual((receipt['verified'],receipt['errors']),(9,0))
    def test_generated_sum_and_equality_bridge(self):
        rng=random.Random(42)
        for _ in range(100):
            a=[rng.randint(-10**30,10**30) for _ in range(rng.randint(0,50))]
            self.assertEqual(kernel.verified.Sum(a),sum(a))
            self.assertTrue(kernel.same_ints(a,list(a)))
            self.assertFalse(kernel.same_ints(a,a+[1]))
        for a,b in [('0.1','0.2'),('-1.12','2.1'),('123456789012345678901234567890.01','.09')]:
            self.assertEqual(Fraction(kernel.sum_decimals(a,b)),Fraction(a)+Fraction(b))
        with self.assertRaises(TypeError):kernel.same_ints([True],[1])

    def test_signed_rational_tolerance_matches_exact_definition(self):
        for a,b,eps in itertools.product(range(-3,4),range(-3,4),range(3)):
            x,y,t=Fraction(a,7),Fraction(b,11),Fraction(eps,100)
            self.assertEqual(kernel.equal_rationals(x,y,t),abs(x-y)<=t)


if __name__=='__main__':unittest.main(verbosity=2)
