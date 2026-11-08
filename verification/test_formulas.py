"""Formula cache regressions and independently specified arithmetic cases."""
from fractions import Fraction
import itertools
import math
import unittest
import xml.etree.ElementTree as ET

from . import formulas, rawxlsx
from . import test_automatic as fixture

alter, S = fixture.alter, fixture.S


def book(cells, expressions, kinds=None):
    return {'Data': rawxlsx.Sheet(cells, [], {}, kinds or {ref:'n' for ref in cells},
                                 formulas={ref:{'text':text,'kind':'normal'} for ref,text in expressions.items()})}


class FormulaArithmetic(unittest.TestCase):
    def test_numeric_cache_exponents_are_bounded_before_fraction_allocation(self):
        for literal in ('1e1000000000','9'*9000):
            source=book({'A1':literal,'B1':'0'},{'B1':'A1+1'})
            result=formulas.inspect(source,{('Data','B1')})
            self.assertTrue(result['partial'],result)
            self.assertEqual(result['checked'],0)

    def test_finite_exhaustive_integer_operations_against_fraction_spec(self):
        for a,b in itertools.product(range(-5,6),repeat=2):
            for op,expected in [('+',Fraction(a+b)),('-',Fraction(a-b)),('*',Fraction(a*b))] + ([('/',Fraction(a,b))] if b else []):
                source=book({'A1':str(a),'A2':str(b),'B1':str(float(expected))},{'B1':f'A1{op}A2'})
                result=formulas.inspect(source,{('Data','B1')})
                self.assertEqual((result['failures'],result['partial']),([],[]),(a,b,op,result))
                self.assertEqual(result['checked'],1)

    def test_excel_operator_precedence_and_function_results(self):
        cases={'-2^2':4,'2^3^2':64,'(2+3)*4':20,'50%*20':10,
               'SUM(1,2,3)':6,'AVERAGE(1,2,6)':3,'MIN(3,-7,8)':-7,
               'MAX(3,-7,8)':8,'COUNT(1,2,3)':3,'ABS(-3)':3,'INT(-1.2)':-2,
               'ROUND(2.675,2)':'2.68','ROUND(-2.675,2)':'-2.68',
               'ROUNDUP(-2.671,2)':'-2.68','ROUNDDOWN(-2.679,2)':'-2.67',
               'ROUND(125,-1)':130,'IF(1<2,12,999)':12,'IF(FALSE,9^1001,7)':7}
        for expression,expected in cases.items():
            result=formulas.inspect(book({'A1':str(expected)},{'A1':expression}),{('Data','A1')})
            self.assertFalse(result['partial'] or result['failures'],(expression,result))

    def test_range_ignores_text_and_blank_but_direct_numeric_text_is_coerced(self):
        source=book({'A1':'1','A2':'note','A3':None,'A4':'2','B1':'3'}, {'B1':'SUM(A1:A4)'},
                    {'A1':'n','A2':'s','A3':'s','A4':'n','B1':'n'})
        result=formulas.inspect(source,{('Data','B1')})
        self.assertEqual((result['failures'],result['partial']),([],[]))

    def test_source_decimal_arithmetic_and_bounded_float_cache_bridge(self):
        source=book({'A1':'0.1','A2':'0.2','B1':'0.30000000000000004'}, {'B1':'A1+A2'})
        self.assertFalse(formulas.inspect(source,{('Data','B1')})['failures'])
        source['Data'].cells['B1']='0.301'
        self.assertTrue(formulas.inspect(source,{('Data','B1')})['failures'])

    def test_native_excel_rounding_uses_the_bounded_decoded_operand_contract(self):
        source=book({'A1':'2.6749999999999998','B1':'2.68'},{'B1':'ROUND(A1,2)'})
        result=formulas.inspect(source,{('Data','B1')})
        self.assertEqual((result['failures'],result['partial']),([],[]))
        source['Data'].cells['A1']='2.6749999999999998123'
        self.assertTrue(formulas.inspect(source,{('Data','B1')})['partial'])

    def test_stale_formula_and_stale_dependency_cannot_pass(self):
        source=book({'A1':'10','B1':'999','C1':'1000'}, {'B1':'A1*2','C1':'B1+1'})
        result=formulas.inspect(source,{('Data','C1')})
        self.assertTrue(result['failures']);self.assertEqual(result['checked'],0)

    def test_shared_formula_relative_absolute_and_quoted_sheet_references(self):
        source=book({'A1':'3','A2':'4','B1':'9','B2':'12'}, {})
        source['Other sheet']=rawxlsx.Sheet({'A1':'3'},[],{}, {'A1':'n'})
        source['Data'].formulas={
            'B1':{'text':"A1*'Other sheet'!$A$1",'kind':'shared','range':'B1:B2'},
            'B2':{'text':"A1*'Other sheet'!$A$1",'kind':'shared','origin':'B1','range':'B1:B2'}}
        result=formulas.inspect(source,{('Data','B2')})
        self.assertEqual((result['failures'],result['partial']),([],[]));self.assertEqual(result['checked'],1)

    def test_unsupported_cycles_external_links_and_budgets_are_incomplete(self):
        cases=['UNKNOWN(2)','SUM([other.xlsx]Data!A1)','A1+1','9^1001','('+ '('*70+'1'+')'*70+')']
        for expression in cases:
            result=formulas.inspect(book({'A1':'12'},{'A1':expression}),{('Data','A1')})
            self.assertTrue(result['partial'],(expression,result));self.assertEqual(result['checked'],0)

    def test_array_region_result_without_own_formula_is_not_certified(self):
        source=book({'A1':'12','A2':'34'}, {})
        source['Data'].formula_regions=['A1:A2']
        result=formulas.inspect(source,{('Data','A2')})
        self.assertTrue(result['partial']);self.assertEqual(result['checked'],0)

    def test_unknown_does_not_hide_known_bad_cache(self):
        result=formulas.inspect(book({'A1':'12','B1':'999'}, {'A1':'UNKNOWN(2)','B1':'2*3'}),{('Data','A1'),('Data','B1')})
        self.assertTrue(result['partial']);self.assertTrue(result['failures'])

    def test_text_boolean_comparisons_and_unknown_branches_do_not_claim_coverage(self):
        for expression in ('IF(TRUE=1,1,2)','IF("10">10,1,2)','IF(FALSE,UNKNOWN(2),7)','IF("1",1,2)'):
            result=formulas.inspect(book({'A1':'2'},{'A1':expression}),{('Data','A1')})
            self.assertTrue(result['partial'],(expression,result))
            self.assertEqual(result['checked'],0)

    def test_iso_dates_are_not_misclassified_as_ignored_sum_text(self):
        source=book({'A1':'2026-01-01','B1':'0'},{'B1':'SUM(A1)'},{'A1':'d','B1':'n'})
        result=formulas.inspect(source,{('Data','B1')})
        self.assertTrue(result['partial']);self.assertEqual(result['checked'],0)

    def test_filter_domains_and_structured_refs_are_collected(self):
        refs=formulas.references({'sheet':'Data','rows':[{'refs':['A1','A2']}],
                                  'filter':{'sheet':'Other','cell':'B1','domain':['B2','B3']}})
        self.assertEqual(refs,{('Data','A1'),('Data','A2'),('Other','B1'),('Other','B2'),('Other','B3')})


class FormulaApplication(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app

    def formula(self, expression, cached, cell_ref='B2', **attributes):
        def change(data):
            root=ET.fromstring(data)
            cell=next(n for n in root.findall('.//'+S+'c') if n.get('r')==cell_ref)
            for old in list(cell):cell.remove(old)
            cell.set('t','n')
            f=ET.SubElement(cell,S+'f',attributes);f.text=expression
            if cached is not None:ET.SubElement(cell,S+'v').text=str(cached)
            return ET.tostring(root,encoding='utf-8')
        self.book.write_bytes(alter(self.book.read_bytes(),'xl/worksheets/sheet1.xml',change))

    def test_correct_formula_cache_is_checked_without_modifying_source(self):
        self.formula('B3*2',68);self.make();before=self.book.read_bytes()
        report=self.run_app(True)
        self.assertEqual(report.verification['state'],'passed',report.verification)
        self.assertEqual(report.verification['formula_checked'],1)
        self.assertEqual(self.book.read_bytes(),before)

    def test_stale_cache_blocks_preview_and_export(self):
        import core
        self.formula('B3*2',999);self.make()
        preview=self.run_app();self.assertTrue(preview.blocked,preview.verification)
        with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_unknown_formula_never_gets_passed(self):
        self.formula('UNKNOWN(B3)',12);self.make()
        report=self.run_app()
        self.assertEqual(report.verification['state'],'partial',report.verification)
        self.assertTrue(any('formula' in reason for reason in report.verification['reasons']))

    def test_shared_formula_metadata_is_resolved_by_independent_reader(self):
        self.formula('B3*2',68,t='shared',si='0',ref='B2:B3')
        # ref in this helper selects the edited cell; set the region explicitly.
        raw=rawxlsx.workbook(self.book)['Data']
        self.assertEqual(raw.formulas['B2']['kind'],'shared')

    def test_shared_formula_followers_are_checked_against_shifted_references(self):
        self.write([['Input','Result'],[3,0],[4,0]])
        self.formula('A2*2',6,t='shared',si='0',ref='B2:B3')
        self.formula('',8,cell_ref='B3',t='shared',si='0')
        self.make();report=self.run_app()
        self.assertEqual(report.verification['state'],'passed',report.verification)
        self.assertEqual(report.verification['formula_checked'],2)
        self.formula('',999,cell_ref='B3',t='shared',si='0')
        self.assertTrue(self.run_app().blocked)


if __name__=='__main__':unittest.main()
