"""Finite exhaustive mathematics, schema regressions and adversarial output tests.

Run: python -m verification selftest. No model, network or new dependencies.
"""
import copy
from fractions import Fraction
import io
import itertools
import json
import math
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
import core
import facts
import layouts
import suggest
from . import arithmetic as ref
from .check import check
from .oracle import coverage
from .output import chart, A
from .pipeline import certificate, generate
from .rawxlsx import workbook, digest


def rates(path, dr=0, dc=0, broken=False):
    wb=openpyxl.Workbook();ws=wb.active;ws.title='Rates'
    def put(r,c,v):ws.cell(r+dr,c+dc,v)
    put(1,1,'2024 school rates')
    put(3,1,'Population');put(3,3,'Primary')
    put(4,1,'min');put(4,2,'max')
    for c,v in enumerate(['food','operations','total'],3):put(4,c,v)
    for r,values in enumerate([[1,200,2700,300,3000],[201,500,2400,200,2600],[501,None,2200,200,9999 if broken else 2400],['Named School',None,2900,200,3100]],5):
        for c,v in enumerate(values,1):
            if v is not None:put(r,c,v)
    wb.save(path)


def staff(path, dr=0, dc=0, metric='현원'):
    wb=openpyxl.Workbook();ws=wb.active;ws.title='Staff'
    def put(r,c,v):ws.cell(r+dr,c+dc,v)
    put(1,1,'Staff 2024-07-31')
    put(3,1,'Department')
    for start,name in [(2,'Parent total'),(5,'Department A'),(8,'Department B')]:
        put(3,start,name)
        ws.merge_cells(start_row=3+dr,start_column=start+dc,end_row=3+dr,end_column=start+dc+2)
        for offset,value in enumerate(['합\n계','일\n반\n직','별\n정\n직']):put(4,start+offset,value)
    put(5,1,metric)
    for start,values in [(2,[25,23,2]),(5,[12,11,1]),(8,[13,12,1])]:
        for offset,value in enumerate(values):put(5,start+offset,value)
    wb.save(path)


def rewrite_zip(source,target,member,transform):
    with ZipFile(source) as src,ZipFile(target,'w',ZIP_DEFLATED) as dest:
        for name in src.namelist():
            content=src.read(name)
            dest.writestr(name,transform(content) if name==member else content)


class ExactMathematics(unittest.TestCase):
    def test_growth_matches_production_over_finite_signed_domain(self):
        count=0
        for a,b in itertools.product(range(-25,26),repeat=2):
            if b==0:
                self.assertIsNone(facts.ratio(a,b))
                with self.assertRaises(ValueError):ref.growth(a,b)
            else:
                self.assertTrue(math.isclose(facts.ratio(a,b),float(ref.growth(a,b)),rel_tol=1e-12,abs_tol=1e-12))
            count+=1
        self.assertEqual(count,2601)

    def test_all_small_closed_intervals_and_boundaries(self):
        for a in range(8):
            for b in range(a,8):
                for population in range(10):
                    if a<=population<=b:self.assertEqual(ref.lookup([(a,b,'selected')],population),'selected')
                    else:
                        with self.assertRaises(ValueError):ref.lookup([(a,b,'selected')],population)
        self.assertEqual(ref.lookup([(1,200,'a'),(201,500,'b')],200),'a')
        self.assertEqual(ref.lookup([(1,200,'a'),(201,500,'b')],201),'b')
        for bad in [[(1,None,'a')],[(1,200,'a'),(200,500,'b')],[(5,4,'a')]]:
            with self.assertRaises(ValueError):ref.lookup(bad,100)

    def test_every_assignment_of_four_atoms_to_three_groups(self):
        atoms=range(4);values=[2,3,5,7]
        for assignment in itertools.product(range(3),repeat=4):
            groups=[]
            for group in range(3):
                members=[i for i in atoms if assignment[i]==group]
                groups.append({'kind':'count','measure':'headcount','unit':'people','period':'2024','atoms':members,'value':sum(values[i] for i in members)})
            self.assertEqual(ref.partition_sum(groups,atoms),17)
            groups.append(copy.deepcopy(groups[assignment[0]]))
            with self.assertRaises(ValueError):ref.partition_sum(groups,atoms)

    def test_nonadditive_units_missing_scope_and_mixed_periods_rejected(self):
        a={'kind':'amount','measure':'cost','unit':'KRW','period':'2024','atoms':['a'],'value':'0.1'}
        b={**a,'atoms':['b'],'value':'0.2'}
        self.assertEqual(ref.partition_sum([a,b],['a','b']),Fraction(3,10))
        for bad in [{**b,'kind':'rate'},{**b,'period':'2023'},{**b,'unit':'USD'},{**b,'atoms':['a']}]:
            with self.assertRaises(ValueError):ref.partition_sum([a,bad],['a','b'])
        with self.assertRaises(ValueError):ref.partition_sum([a],['a','b'])

    def test_rounding_and_missing_values_are_explicit(self):
        self.assertEqual(str(ref.rounded(Fraction('2.675'),2)),'2.68')
        self.assertEqual(str(ref.rounded(Fraction('-2.675'),2)),'-2.68')
        self.assertEqual(ref.exact('1,234.50'),Fraction(2469,2))
        for value in [None,True,'-','', '1,2',float('nan')]:
            with self.assertRaises(ValueError):ref.exact(value)
        large='123456789012345678901234567890.01'
        self.assertEqual(Fraction(layouts._cost_sum(large,'0.09')),ref.exact(large)+ref.exact('0.09'))


class SourceAndOutput(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)

    def tearDown(self):self.temp.cleanup()

    def make(self,kind='rates',lang='en',**kwargs):
        xlsx=self.folder/f'{kind}-{len(list(self.folder.iterdir()))}.xlsx'
        (rates if kind=='rates' else staff)(xlsx,**kwargs)
        out=self.folder/(xlsx.stem+'-output')
        result=generate(xlsx,out,lang)
        self.assertEqual(result['result'],'PASS',result)
        return xlsx,out,json.loads((out/'certificate.json').read_text(encoding='utf-8'))

    def test_both_layouts_and_languages_keep_original_values_and_notes(self):
        for kind,lang in itertools.product(['rates','staff'],['ko','en']):
            xlsx,out,cert=self.make(kind,lang,**({'broken':True} if kind=='rates' else {}))
            checked=check(xlsx,out/'report.pptx',cert)
            if kind=='rates':self.assertEqual({f['kind'] for f in checked['source_findings']},{'unknown_upper_bound','source_total_mismatch'})
            else:self.assertEqual(checked['source_findings'],[])
            self.assertTrue(cert['tables'])
            self.assertEqual(digest(xlsx),cert['workbook_sha256'])

    def test_block_translation_preserves_value_sequence(self):
        for kind in ['rates','staff']:
            seen=[]
            for dr,dc in [(0,0),(3,2),(7,5)]:
                xlsx,out,cert=self.make(kind,dr=dr,dc=dc)
                plan=suggest.plan_report(core.ExcelSource(str(xlsx)))
                self.assertTrue(all('2024' in section.heading for section in plan.sections))
                book=workbook(xlsx);sheet=next(iter(book.values()))
                seen.append([sheet.get(t['refs'][0]) for p in cert['tables'] for row in p['rows'] for t in row if t['op']=='cell'])
            self.assertEqual(seen[0],seen[1]);self.assertEqual(seen[0],seen[2])

    def test_organization_totals_and_children_stay_separate(self):
        xlsx,out,cert=self.make('staff')
        rows=cert['tables'][0]['rows']
        self.assertEqual([r[1]['refs'][0] for r in rows],['B5','E5','H5'])
        self.assertEqual(len(rows),3)  # no invented 50, no dropped children based only on sums
        plan=suggest.plan_report(core.ExcelSource(str(xlsx)))
        self.assertEqual(plan.doubts,[('structured_layout',{})])
        self.assertTrue(all(s.table.code=='table_layout' for s in plan.sections))
        self.assertTrue(all(s.chart is None for s in plan.sections))

    def test_multi_file_report_routes_structured_references(self):
        first=self.folder/'rates.xlsx';second=self.folder/'staff.xlsx'
        rates(first);staff(second)
        sources=core.Sources([str(first),str(second)])
        plan=suggest.plan_files(sources,'en')
        template=self.folder/'multi-template.pptx';out=self.folder/'multi.pptx'
        suggest.make_report(str(template),'Multiple sources',plan.sections)
        report=core.run(str(template),[str(first),str(second)],str(out),source=sources)
        self.assertEqual(report.count('error'),0)
        self.assertEqual(sum(s.has_table for slide in Presentation(out).slides for s in slide.shapes),2)

    def test_quota_and_headcount_headers_remain_distinct(self):
        xlsx,out,cert=self.make('staff',metric='정원')
        self.assertEqual(cert['tables'][0]['headers'][1],'Quota (source)')

    def test_structured_table_uses_full_width_and_fits_slide(self):
        for kind in ['rates','staff']:
            xlsx,out,cert=self.make(kind)
            prs=Presentation(out/'report.pptx')
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_table:
                        self.assertGreater(shape.width,Inches(11))
                        self.assertLessEqual(shape.top+shape.height,prs.slide_height)

    def test_no_declared_schema_does_not_fabricate_report(self):
        wb=openpyxl.Workbook();ws=wb.active;ws.append(['현원',10]);path=self.folder/'no.xlsx';wb.save(path)
        self.assertEqual(layouts.discover(core.ExcelSource(str(path)).wb.active),[])

    def test_certificate_mutations_are_detected(self):
        xlsx,out,cert=self.make()
        mutations=[]
        c=copy.deepcopy(cert);c['tables'][0]['rows'].pop();c['tables'][0]['stop']-=1;mutations.append(c)
        c=copy.deepcopy(cert);c['tables'][0]['rows'][0][1]['refs']=['A5'];mutations.append(c)
        c=copy.deepcopy(cert);c['tables'][0]['rows'][0][1],c['tables'][0]['rows'][0][2]=c['tables'][0]['rows'][0][2],c['tables'][0]['rows'][0][1];mutations.append(c)
        c=copy.deepcopy(cert);c['tables']*=2;mutations.append(c)
        c=copy.deepcopy(cert);c['tables'][0]['headers'][1]='Spending';mutations.append(c)
        c=copy.deepcopy(cert);c['workbook_sha256']='0'*64;mutations.append(c)
        for altered in mutations:self.assertEqual(check(xlsx,out/'report.pptx',altered)['result'],'FAIL')

    def test_saved_value_row_and_warning_mutations_are_detected(self):
        xlsx,out,cert=self.make(broken=True)
        def bad_value(raw):
            root=ET.fromstring(raw)
            table=root.find('.//'+A+'tbl')
            table.findall(A+'tr')[1].findall(A+'tc')[1].find('.//'+A+'t').text='999999'
            return ET.tostring(root)
        def missing_row(raw):
            root=ET.fromstring(raw);tbl=root.find('.//'+A+'tbl');tbl.remove(tbl.findall(A+'tr')[-1]);return ET.tostring(root)
        def missing_warning(raw):
            root=ET.fromstring(raw)
            for node in root.findall('.//'+A+'t'):
                if 'source total' in (node.text or ''):node.text='Source data checked.'
            return ET.tostring(root)
        def wrong_period(raw):return raw.replace(b'2024',b'2023')
        def wrong_group(raw):return raw.replace(b'Primary',b'Other')
        for i,mutation in enumerate([bad_value,missing_row,missing_warning,wrong_period,wrong_group]):
            bad=self.folder/f'bad-{i}.pptx';rewrite_zip(out/'report.pptx',bad,'ppt/slides/slide1.xml',mutation)
            self.assertEqual(check(xlsx,bad,cert)['result'],'FAIL')

    def test_already_existing_output_is_not_overwritten(self):
        xlsx,out,cert=self.make()
        old=digest(out/'report.pptx')
        with self.assertRaises(ValueError):generate(xlsx,out)
        self.assertEqual(old,digest(out/'report.pptx'))

    def test_corrupt_zip_and_unverified_chart_fail_closed(self):
        xlsx,out,cert=self.make()
        bad=self.folder/'corrupt.pptx';bad.write_bytes(b'not a zip')
        self.assertEqual(check(xlsx,bad,cert)['result'],'FAIL')
        for invalid in [[],{},{'tables':None},{'tables':[],'charts':None},{'tables':[None]}]:
            self.assertEqual(check(xlsx,out/'report.pptx',invalid)['result'],'FAIL')
        prs=Presentation(out/'report.pptx');data=CategoryChartData();data.categories=['A'];data.add_series('Wrong',[999])
        prs.slides[0].shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(2), Inches(2),data)
        bad=self.folder/'extra-chart.pptx';prs.save(bad)
        self.assertEqual(check(xlsx,bad,cert)['result'],'FAIL')

    def test_decimal_source_inconsistency_is_preserved_and_disclosed(self):
        path=self.folder/'decimal.xlsx';rates(path)
        wb=openpyxl.load_workbook(path);ws=wb.active
        for c,value in zip(['C5','D5','E5'],[0.1,0.2,0.4]):ws[c]=value
        wb.save(path)
        checked=generate(path,self.folder/'decimal-output')
        self.assertEqual(checked['result'],'PASS',checked)
        self.assertIn('source_total_mismatch',{f['kind'] for f in checked['source_findings']})

    def test_read_only_check_has_no_application_imports(self):
        # Static dependency guard complements runtime numerical tests.
        for filename in ['arithmetic.py','rawxlsx.py','oracle.py','output.py','check.py']:
            text=(Path(__file__).parent/filename).read_text(encoding='utf-8')
            self.assertNotIn('import core',text);self.assertNotIn('import layouts',text);self.assertNotIn('import suggest',text)

    def test_chart_cache_and_embedded_workbook_mutations(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6])
        data=CategoryChartData();data.categories=['A','B'];data.add_series('Count',[12,34])
        slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(5), Inches(4), data)
        original=self.folder/'chart.pptx';prs.save(original)
        args=('ppt/charts/chart1.xml',['A','B'],[('Count',['12','34'])])
        self.assertEqual(chart(original,*args),[])
        def bad_cache(raw):return raw.replace(b'<c:v>12</c:v>',b'<c:v>99</c:v>')
        bad=self.folder/'chart-cache.pptx';rewrite_zip(original,bad,args[0],bad_cache)
        self.assertTrue(chart(bad,*args))
        with ZipFile(original) as z:embedded=next(n for n in z.namelist() if n.startswith('ppt/embeddings/'))
        def bad_embedded(raw):
            src=io.BytesIO(raw);dest=io.BytesIO()
            with ZipFile(src) as zz,ZipFile(dest,'w',ZIP_DEFLATED) as out:
                for name in zz.namelist():
                    content=zz.read(name)
                    if name.startswith('xl/worksheets/'):content=content.replace(b'<v>12</v>',b'<v>99</v>')
                    out.writestr(name,content)
            return dest.getvalue()
        bad=self.folder/'chart-embedded.pptx';rewrite_zip(original,bad,embedded,bad_embedded)
        self.assertTrue(chart(bad,*args))


if __name__=='__main__':unittest.main(verbosity=2)
