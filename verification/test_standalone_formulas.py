"""A matching source hash and report must not certify a stale formula cache."""
import copy
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET

from . import test_verification as fixture
from .check import check
from .rawxlsx import S, digest, workbook


class StandaloneFormulaCaches(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)
        self.original=self.folder/'rates.xlsx'
        fixture.rates(self.original)
        self.output=self.folder/'output'
        result=fixture.generate(self.original,self.output,'en')
        self.assertEqual(result['result'],'PASS',result)
        import json
        self.certificate=json.loads((self.output/'certificate.json').read_text(encoding='utf-8'))

    def tearDown(self):self.temp.cleanup()

    def with_formulas(self, entries):
        target=self.folder/'formulas.xlsx'
        def inject(raw):
            root=ET.fromstring(raw)
            for address,text,cached,attributes in entries:
                cell=root.find('.//'+S+f'c[@r="{address}"]')
                if cell is None:
                    row=root.find('.//'+S+f'row[@r="{int("".join(c for c in address if c.isdigit()))}"]')
                    cell=ET.SubElement(row,S+'c',{'r':address,'t':'n'})
                    ET.SubElement(cell,S+'v').text=str(cached)
                self.assertEqual(cell.find(S+'v').text,str(cached))
                formula=ET.Element(S+'f',attributes);formula.text=text
                cell.insert(0,formula)
            return ET.tostring(root)
        fixture.rewrite_zip(self.original,target,'xl/worksheets/sheet1.xml',inject)
        certificate=copy.deepcopy(self.certificate)
        certificate['workbook_sha256']=digest(target)
        return target,certificate

    def checked(self,entries):
        source,certificate=self.with_formulas(entries)
        before=digest(source)
        result=check(source,self.output/'report.pptx',certificate)
        self.assertEqual(digest(source),before)
        return result

    def test_matching_supported_formula_and_dependency_pass(self):
        result=self.checked([('C5','3*900',2700,{}),('E5','C5+D5',3000,{})])
        self.assertEqual(result['result'],'PASS',result)
        self.assertEqual(result['formula_checked'],2)

    def test_unchanged_cached_values_and_refreshed_hash_do_not_hide_wrong_math(self):
        source,certificate=self.with_formulas([('E5','C5+D5+1',3000,{})])
        self.assertEqual(workbook(source)['Rates'].get('E5'),'3000')
        self.assertEqual(certificate['workbook_sha256'],digest(source))
        result=check(source,self.output/'report.pptx',certificate)
        self.assertEqual(result['result'],'FAIL',result)
        self.assertTrue(any('stale or inconsistent formula result at Rates!E5' in error for error in result['errors']),result)

    def test_unknown_formula_in_declared_value_is_refused(self):
        result=self.checked([('E5','FORECAST(1,2,3)',3000,{})])
        self.assertEqual(result['result'],'FAIL',result)
        self.assertTrue(any('source formula cannot be independently checked' in error for error in result['errors']),result)

    def test_formula_outside_table_rows_is_audited_for_full_sheet_schema_coverage(self):
        result=self.checked([('G1','UNSUPPORTED(42)',42,{})])
        self.assertEqual(result['result'],'FAIL',result)
        self.assertTrue(any('Rates!G1' in error for error in result['errors']),result)

    def test_array_region_is_not_certified_from_its_saved_values(self):
        result=self.checked([('E5','C5+D5',3000,{'t':'array','ref':'E5:E6'})])
        self.assertEqual(result['result'],'FAIL',result)
        self.assertTrue(any('array/data-table' in error for error in result['errors']),result)
