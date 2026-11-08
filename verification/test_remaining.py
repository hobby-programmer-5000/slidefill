"""Regression contracts for short cells, duration literals and exact aggregates."""
from fractions import Fraction
import io
import math
import random
import struct
import unittest
from unittest.mock import patch
from zipfile import ZipFile,ZIP_DEFLATED
import core,readers
from pptx import Presentation
from pptx.util import Inches
from . import binary,kernel,recipes,sources
from . import test_automatic as fixture
from .test_binary import xlsb_fixture,record
from .test_coverage import ods


def shorts(path,all_cells=False):
    """Original synthetic writer transformation, not a vendor test workbook."""
    result=io.BytesIO()
    with ZipFile(path) as src,ZipFile(result,'w',ZIP_DEFLATED) as dst:
        for name in src.namelist():
            raw=src.read(name)
            if name.endswith('sheet1.bin'):
                raw=b''.join(record(kind+11,body[4:]) if 1<=kind<=7 and (all_cells or struct.unpack_from('<I',body)[0]==1) else record(kind,body) for kind,body in binary.xlsb_records(raw))
            dst.writestr(name,raw)
    path.write_bytes(result.getvalue())


class Remaining(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app
    blocked_by=fixture.AutomaticGuard.blocked_by

    def test_short_strings_numbers_row_reset_preview_save_and_source_mutation(self):
        for all_cells in (False,):
            with self.subTest(all_cells=all_cells):
                self.book=self.root/'data.xlsb';xlsb_fixture(self.book);shorts(self.book,all_cells);self.make()
                before=self.book.read_bytes()
                self.assertEqual(self.run_app().verification['state'],'passed')
                self.assertEqual(self.run_app(True).verification['state'],'passed')
                self.out.unlink();self.assertEqual(before,self.book.read_bytes())
                self.assertEqual(sources.read(self.book)['Data'].get('B3'),'-2.5')
                original=core.ExcelSource.range_cells
                def bad(source,key):
                    grid,where,issues=original(source,key);grid[2][1]=(-3,'-3','General');return grid,where,issues
                with patch.object(core.ExcelSource,'range_cells',bad):
                    self.assertTrue(self.run_app().blocked)
                    with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_all_seven_short_kinds_and_blank_column_advance(self):
        self.book=self.root/'data.xlsb';xlsb_fixture(self.book)
        wide=lambda text:struct.pack('<I',len(text))+text.encode('utf-16-le')
        # A full cell anchors the row; full gap at C then short blank at D.
        extra=record(0,struct.pack('<I',4))+record(6,bytes(8)+wide('X'))+record(5,struct.pack('<II',2,0)+struct.pack('<d',5))
        extra+=record(12,bytes(4))+record(13,bytes(4)+struct.pack('<I',((-7<<2)|2)&0xffffffff))
        extra+=record(14,bytes(4)+bytes([7]))+record(15,bytes(4)+bytes([1]))+record(16,bytes(4)+struct.pack('<d',1.25))+record(18,bytes(4)+struct.pack('<I',0))
        extra+=record(17,bytes(4)+wide('inline'))
        buffer=io.BytesIO()
        with ZipFile(self.book) as src,ZipFile(buffer,'w',ZIP_DEFLATED) as dst:
            for name in src.namelist():
                raw=src.read(name)
                if name.endswith('sheet1.bin'):raw=raw[:-len(record(146)+record(130))]+extra+record(146)+record(130)
                dst.writestr(name,raw)
            dst.writestr('xl/sharedStrings.bin',record(159,struct.pack('<II',1,1))+record(19,b'\0'+wide('shared'))+record(160))
        self.book.write_bytes(buffer.getvalue())
        independent=sources.read(self.book)['Data'];production=readers.load(str(self.book)).active
        for ref,value in [('A5','X'),('C5',5),('D5',None),('E5',-7),('F5','#DIV/0!'),('G5',True),('H5',1.25),('I5','shared'),('J5','inline')]:
            self.assertEqual(production[ref].value,value,ref)
            expected=independent.get(ref)
            if isinstance(value,(int,float)) and not isinstance(value,bool):self.assertEqual(Fraction(expected),Fraction(str(value)),ref)
            else:self.assertEqual(expected,value,ref)

    def test_unanchored_short_row_is_rejected_after_native_excel_failure(self):
        self.book=self.root/'data.xlsb';xlsb_fixture(self.book);shorts(self.book,True)
        with self.assertRaises(sources.Unsupported):sources.read(self.book)
        with self.assertRaises(ValueError):readers.load(str(self.book))

    def test_short_truncation_overflow_and_order_are_explicit(self):
        self.book=self.root/'data.xlsb';xlsb_fixture(self.book)
        for body in [bytes(3),bytes(4),bytes(9)]:
            original=self.book.read_bytes();buffer=io.BytesIO()
            with ZipFile(io.BytesIO(original)) as src,ZipFile(buffer,'w',ZIP_DEFLATED) as dst:
                for name in src.namelist():dst.writestr(name,src.read(name)+record(16,body) if name.endswith('sheet1.bin') else src.read(name))
            self.book.write_bytes(buffer.getvalue())
            with self.assertRaises(sources.Unsupported):sources.read(self.book)
            self.book.write_bytes(original)
        for extra in [record(5,struct.pack('<II',1,0)+struct.pack('<d',1)),
                      record(0,struct.pack('<I',6))+record(5,struct.pack('<II',16383,0)+struct.pack('<d',1))+record(13,bytes(4)+struct.pack('<I',6))]:
            buffer=io.BytesIO()
            with ZipFile(io.BytesIO(original)) as src,ZipFile(buffer,'w',ZIP_DEFLATED) as dst:
                for name in src.namelist():dst.writestr(name,src.read(name)+extra if name.endswith('sheet1.bin') else src.read(name))
            self.book.write_bytes(buffer.getvalue())
            with self.assertRaises(sources.Unsupported):sources.read(self.book)
            with self.assertRaises(ValueError):readers.load(str(self.book))
            self.book.write_bytes(original)

    def test_duration_zero_day_sign_precision_and_repeated_positions(self):
        for literal,expected in [('PT0S','00:00:00'),('PT1H2M3.125S','01:02:03.125000'),('P1D','P1D'),('PT25H','PT25H'),('-PT1H','-PT1H'),('PT0.0000001S','PT0.0000001S'),('P2DT3H4M5.125S','P2DT3H4M5.125S')]:
            with self.subTest(literal=literal):
                self.book=self.root/'data.ods'
                ods(self.book,'<table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="time" office:time-value="'+literal+'" table:number-columns-repeated="2"/></table:table-row>')
                prod=readers.load(str(self.book)).active;raw=sources.read(self.book)['Data']
                for ref in ('A4','B4','A5','B5'):self.assertEqual(prod[ref].value,expected);self.assertEqual(raw.get(ref),expected)
                # Table and numeric report paths also exercise automatic source coverage.
                prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);shape=slide.shapes.add_table(5,2,0,0,Inches(9),Inches(4))
                shape._element.xpath('.//p:cNvPr')[0].set('descr','{{table:Data!A1:B5}}');prs.save(self.template)
                result=self.run_app();self.assertEqual(result.verification['state'],'passed',result.verification)

    def test_duration_literal_mutation_blocks_save(self):
        self.book=self.root/'data.ods';ods(self.book,'<table:table-row><table:table-cell office:value-type="time" office:time-value="-PT25H"/></table:table-row>')
        self.make('Data!A1:B4');original=core.ExcelSource.range_cells
        def bad(source,key):
            grid,where,issues=original(source,key);grid[3][0]=('PT25H','PT25H','General');return grid,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):
            self.assertTrue(self.run_app().blocked)
            with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_html_optional_cell_row_closures_and_row_groups(self):
        self.book=self.root/'Data.xls'
        self.book.write_text('<html><table><thead><tr><th>Item<th>Amount</thead><tbody><tr><td>A &amp; B<td>12<tr><td><b>서울</b><td>34</tbody></table><table><tr><td>Second<td>99</table></html>',encoding='utf-8')
        self.make();self.assertEqual(self.run_app(True).verification['state'],'passed')
        raw=sources.read(self.book);self.assertEqual(raw['Data'].get('A2'),'A & B');self.assertEqual(raw['Data'].get('B3'),'34')
        self.assertEqual(raw['Data 2'].get('B1'),'99')

    def test_exact_cancellation_sum_mean_product_and_ranked_groups(self):
        for kind in ('sum(B)','mean(B)','sum(B*C)'):
            self.write([['Day','Amount','Factor'],['2026-01-01',1e16,1],['2026-01-01',1.0,1],['2026-01-01',-1e16,1],['2026-01-02',0.1,1],['2026-01-02',0.2,1]])
            key='latest2:Data!A,'+kind;self.make(key)
            source=core.ExcelSource(str(self.book));grid,_,_=source.range_cells(key)
            expected=Fraction(1,3) if kind.startswith('mean') else 1
            self.assertEqual(grid[1][1][0],expected,(kind,grid))
            self.assertEqual(self.run_app(True,source).verification['state'],'passed');self.out.unlink()
        self.write([['Item','Amount'],['A',1e16],['A',1.0],['A',-1e16],['B',0.1],['B',0.2]])
        self.make('top2:Data!A,sum(B)');self.assertEqual(self.run_app().verification['state'],'passed')

    def test_one_ulp_aggregate_corruption_and_integer_off_by_one_block(self):
        self.write([['Day','Amount'],['2026-01-01',0.1],['2026-01-01',0.2],['2026-01-02',2]])
        self.make('latest2:Data!A,sum(B)');original=core.ExcelSource.range_cells
        for row in (1,2):
            def bad(source,key):
                grid,where,issues=original(source,key);value=grid[row][1][0]
                value=math.nextafter(value,math.inf) if row==1 else value+1
                grid[row][1]=(value,str(value),'General');return grid,where,issues
            with patch.object(core.ExcelSource,'range_cells',bad):
                self.assertTrue(self.run_app().blocked)
                with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_exact_integer_sum_bridge_and_independent_randomized_recipes(self):
        rng=random.Random(9181)
        for _ in range(250):
            values=[Fraction(rng.randrange(-100000,100001),rng.choice((1,10,100,1000,3,7))) for _ in range(rng.randrange(0,25))]
            self.assertEqual(kernel.sum_rationals(values),sum(values,Fraction()))
        with patch.object(kernel.verified,'Sum',wraps=kernel.verified.Sum) as checked:
            self.assertEqual(kernel.sum_rationals([Fraction(1,10),Fraction(2,10)]),Fraction(3,10));self.assertTrue(checked.called)
        with patch.object(kernel.verified,'SumBound',wraps=kernel.verified.SumBound) as checked:
            for differences in [[],[Fraction(1,10),Fraction(-2,10)],[Fraction(1,3),Fraction(-1,7)]]:
                total,bound=kernel.sum_error_bounds(differences)
                self.assertEqual(total,sum(differences,Fraction()));self.assertEqual(bound,sum(map(abs,differences),Fraction()));self.assertLessEqual(abs(total),bound)
            self.assertTrue(checked.called)

    def test_randomized_group_sums_products_means_and_filtered_values(self):
        rng=random.Random(1318);rows=[['Day','Amount','Factor','Flag']];groups={}
        for day in range(1,9):
            key=f'2026-01-{day:02d}';groups[key]=[]
            values=[1e16,0.1,-1e16]+[rng.randrange(-100,101)/100 for _ in range(20)]
            for index,value in enumerate(values):
                factor=rng.choice([0.1,0.2,3]);factor=1 if index<3 else factor;flag='keep' if index<3 or index%3 else 'skip'
                rows.append([key,value,factor,flag]);groups[key].append((Fraction(str(value)),Fraction(str(factor)),flag))
        self.write(rows)
        for expression,product,mean,filtered in [('sum(B)',False,False,False),('sum(B*C)',True,False,False),('mean(B)',False,True,False),('sum(B, D=keep)',False,False,True)]:
            key='latest8:Data!A,'+expression;source=core.ExcelSource(str(self.book));grid,_,_=source.range_cells(key)
            for line,(day,items) in zip(grid[1:],groups.items()):
                terms=[a*b if product else a for a,b,flag in items if not filtered or flag=='keep']
                exact=sum(terms,Fraction())/(len(terms) if mean else 1)
                expected=exact
                self.assertEqual(line[1][0],expected,(expression,day))
            self.make(key);result=self.run_app(source=source);self.assertEqual(result.verification['state'],'passed',result.verification)

    def test_higher_precision_operand_is_partial_without_masking_damage(self):
        self.book=self.root/'Data.csv'
        self.book.write_text('Day,Amount\n2026-01-01,0.1234567890123456789\n2026-01-02,2\n',encoding='utf-8')
        self.make('latest2:Data!A,sum(B)')
        result=self.run_app();self.assertEqual(result.verification['state'],'partial',result.verification)
        self.assertTrue(any('precision' in reason for reason in result.verification['reasons']))
        original=core.ExcelSource.range_cells
        def bad(source,key):
            grid,where,issues=original(source,key);grid[2][1]=(999,'999','General');return grid,where,issues
        with patch.object(core.ExcelSource,'range_cells',bad):
            self.assertTrue(self.run_app().blocked)
            with self.assertRaises(core.SlideFillError):self.run_app(True)


if __name__=='__main__':unittest.main(verbosity=2)
