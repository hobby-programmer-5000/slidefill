"""Original legacy-BIFF and nested-table fixtures at the automatic app boundary."""
import io
from pathlib import Path
import struct
import unittest
from unittest.mock import patch
import core,readers,numfmt
from pptx import Presentation
from pptx.util import Inches
from . import sources,binary,test_automatic as fixture
from .test_binary import biff,compound_file,unicode_text,xls_fixture


def legacy_fixture(path,version=5,codepage=1252):
    codec='cp'+str(codepage)
    def text(value,width=2):
        raw=value.encode(codec);return len(raw).to_bytes(width,'little')+raw
    if version==5:
        bof=lambda kind:biff(0x809,struct.pack('<HHHH',0x500,kind,0x0dbb,1993))
        prefix=bof(5)+biff(0x42,struct.pack('<H',codepage))
        prefix+=biff(0x31,struct.pack('<HHHHHBBBBB',200,0,0x7fff,400,0,0,0,0,0,5)+b'Arial')
        prefix+=biff(0xe0,bytes(16))+biff(0xe0,struct.pack('<HH',0,164)+bytes(12))
        prefix+=biff(0xe0,struct.pack('<HH',0,14)+bytes(12))
        prefix+=biff(0x41e,struct.pack('<H',164)+text('0.00',1))
        bound=lambda offset:biff(0x85,struct.pack('<IBB',offset,0,0)+text('Data',1))
        globals_=prefix+bound(0)+biff(10);globals_=prefix+bound(len(globals_))+biff(10)
        cell=lambda r,c,xf=0:struct.pack('<HHH',r,c,xf)
        rows=bof(0x10)+biff(0x200,struct.pack('<HHHHH',0,4,0,4,0))
        for r,c,value in [(0,0,'Item'),(0,1,'Amount'),(0,2,'Cache'),(1,0,'Café' if codepage==1252 else '서울')]:rows+=biff(0x204,cell(r,c)+text(value))
        rows+=biff(0xd6,cell(2,0)+text('Rich text')+b'\x01\x05\x00')
        rows+=biff(0x27e,cell(1,1,1)+struct.pack('<I',(12<<2)|2))+biff(0x203,cell(2,1,1)+struct.pack('<d',-2.5))
        rows+=biff(6,cell(1,2)+struct.pack('<d',6.25)+bytes(8))
        rows+=biff(6,cell(2,2)+bytes(6)+b'\xff\xff'+bytes(8))+biff(0x207,text('cached text'))+biff(10)
        rows=rows[:-4]+biff(0x203,cell(1,3,2)+struct.pack('<d',46023))+biff(10)
        path.write_bytes(compound_file(globals_+rows));return
    ident={2:9,3:0x209,4:0x409}[version]
    data=biff(ident,struct.pack('<HH',version*0x100,0x10)+(bytes(2) if version>2 else b''))+biff(0x42,struct.pack('<H',codepage))
    if version==2:
        data+=biff(0x31,struct.pack('<HH',200,0)+text('Arial',1))+biff(0x1e,text('General',1))
        data+=biff(0x43,bytes(4));cell=lambda r,c:struct.pack('<HH',r,c)+bytes(3)
    else:
        data+=biff(0x231,struct.pack('<HHH',200,0,0x7fff)+text('Arial',1))
        data+=biff(0x1e if version==3 else 0x41e,(bytes(2) if version==4 else b'')+text('General',1))
        data+=biff(0x243 if version==3 else 0x443,bytes(12));cell=lambda r,c:struct.pack('<HHH',r,c,0)
    for r,c,value in [(0,0,'Item'),(0,1,'Amount'),(0,2,'Cache'),(1,0,'Café'),(2,0,'South')]:
        data+=biff(4 if version==2 else 0x204,cell(r,c)+text(value,1 if version==2 else 2))
    data+=biff(2 if version==2 else 0x203,cell(1,1)+(struct.pack('<H',12) if version==2 else struct.pack('<d',12)))
    data+=biff(3 if version==2 else 0x203,cell(2,1)+struct.pack('<d',-2.5))
    data+=biff(6 if version==2 else 0x206 if version==3 else 0x406,cell(1,2)+struct.pack('<d',6.25)+(bytes(2) if version==2 else bytes(4)))
    data+=biff(5 if version==2 else 0x205,cell(2,2)+bytes([1,0]))+biff(10)
    path.write_bytes(data)


NESTED='<table><tr><th>Outer item</th><th>Amount</th></tr><tr><td>Before <table><tr><th>Inner item</th><th>Amount</th></tr><tr><td>A</td><td>12</td></tr><tr><td>B</td><td>34</td></tr></table> After</td><td>46</td></tr><tr><td>Tail</td><td>10</td></tr></table>'


class LegacyHTML(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app
    blocked_by=fixture.AutomaticGuard.blocked_by

    def table(self,key,columns=3):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);shape=slide.shapes.add_table(3,columns,0,0,Inches(9),Inches(4))
        shape._element.xpath('.//p:cNvPr')[0].set('descr','{{table:'+key+'}}');prs.save(self.template)

    def corruption_blocks(self,key,row=1,col=1):
        original=core.ExcelSource.range_cells
        def changed(source,ref):
            grid,where,issues=original(source,ref);grid[row][col]=(999,'999','General');return grid,where,issues
        with patch.object(core.ExcelSource,'range_cells',changed):
            self.assertTrue(self.run_app().blocked)
            with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.assertFalse(self.out.exists())

    def test_biff2_3_4_5_values_preview_save_and_source_mutation(self):
        for version in (2,3,4,5):
            with self.subTest(version=version):
                self.book=self.root/('biff'+str(version)+'.xls');legacy_fixture(self.book,version)
                sheet='Data' if version==5 else 'Sheet 1';key=sheet+'!A1:C3';raw=sources.read(self.book)[sheet]
                self.assertEqual(raw.get('A2'),'Café');self.assertEqual(raw.get('B2'),'12');self.assertEqual(raw.get('B3'),'-2.5');self.assertEqual(raw.get('C2'),'6.25')
                original=self.book.read_bytes();self.table(key);result=self.run_app();self.assertEqual(result.verification['state'],'passed',result.verification)
                self.assertEqual(self.run_app(True).verification['state'],'passed');self.assertEqual(original,self.book.read_bytes());self.out.unlink()
                self.corruption_blocks(key)

    def test_biff5_korean_codepage_rich_text_and_custom_format(self):
        self.book=self.root/'korean.xls';legacy_fixture(self.book,5,949);sheet=sources.read(self.book)['Data']
        self.assertEqual(sheet.get('A2'),'서울');self.assertEqual(sheet.get('A3'),'Rich text');self.assertEqual(sheet.formats['B2'],'0.00');self.assertEqual(sheet.get('C3'),'cached text')
        self.table('Data!A1:C3');result=self.run_app(True);self.assertEqual(result.verification['state'],'passed',result.verification)

    def test_biff8_rstring_unformatted_unicode_and_runs(self):
        self.book=self.root/'rich8.xls';xls_fixture(self.book)
        stream=binary.workbook_stream(self.book);rows=list(binary.biff_records(stream));offset=next(struct.unpack_from('<I',b)[0] for k,b in rows if k==0x85)
        sheet=b''.join(biff(0xd6,b[:6]+unicode_text('서울')+struct.pack('<HHH',1,0,0)) if k==0xfd and struct.unpack_from('<HH',b)==(2,0) else biff(k,b) for k,b in binary.biff_records(stream,offset))
        self.book.write_bytes(compound_file(stream[:offset]+sheet));self.assertEqual(sources.read(self.book)['Data'].get('A3'),'서울')
        self.make();result=self.run_app(True);self.assertEqual(result.verification['state'],'passed',result.verification)

    def test_builtin_legacy_date_uses_windows_short_date_instead_of_us_alias(self):
        self.book=self.root/'date.xls';legacy_fixture(self.book,5)
        locale=numfmt.Locale(date_separator='-',short_date='yyyy-mm-dd',korean=True)
        with patch.object(numfmt,'LOCALE',locale):
            source=core.ExcelSource(str(self.book));raw,text,fmt,issues=source.read(source.wb['Data'],2,4)
            self.assertEqual(text,'2026-01-01');self.assertFalse(issues)

    def test_legacy_truncation_duplicate_cells_and_unknown_codepage_do_not_pass(self):
        for version in (2,3,4):
            self.book=self.root/'bad.xls';legacy_fixture(self.book,version);raw=self.book.read_bytes()
            self.book.write_bytes(raw[:-1])
            with self.assertRaises(sources.Unsupported):sources.read(self.book)
            records=list(binary.biff_records(raw));cell=next((k,b) for k,b in records if k in (4,0x204))
            self.book.write_bytes(raw[:-4]+biff(*cell)+biff(10))
            with self.assertRaises(sources.Unsupported):sources.read(self.book)
        self.book=self.root/'unknown.xls';legacy_fixture(self.book,5)
        raw=binary.workbook_stream(self.book).replace(biff(0x42,struct.pack('<H',1252)),biff(0x42,struct.pack('<H',65535)),1)
        self.book.write_bytes(compound_file(raw))
        with self.assertRaises(sources.Unsupported):sources.read(self.book)
        with self.assertRaises(sources.Unsupported):binary.rich_text(bytes(6)+unicode_text('abc')+struct.pack('<H',2),0x600,'ascii')
        legacy_fixture(self.book,5);stream=binary.workbook_stream(self.book)
        offset=next(struct.unpack_from('<I',b)[0] for k,b in binary.biff_records(stream) if k==0x85)
        records=list(binary.biff_records(stream,offset));formula=next(b for k,b in records if k==6 and struct.unpack_from('<HH',b)==(2,2));cached=next(b for k,b in records if k==0x207)
        sheet=b''.join(biff(k,b) for k,b in records[:-1])+biff(6,formula)+biff(0x207,cached)+biff(10)
        self.book.write_bytes(compound_file(stream[:offset]+sheet))
        with self.assertRaises(sources.Unsupported):sources.read(self.book)

    def test_nested_parent_child_tail_and_markers_are_all_source_linked(self):
        self.book=self.root/'Nested.xls';self.book.write_text(NESTED,encoding='utf-8');raw=sources.read(self.book)
        self.assertEqual(list(raw),['Nested','Nested 2']);self.assertEqual(raw['Nested'].get('A2'),'Before [Nested 2] After')
        self.assertEqual(raw['Nested'].get('B2'),'46');self.assertEqual(raw['Nested'].get('B3'),'10');self.assertEqual(raw['Nested 2'].get('B3'),'34')
        before=self.book.read_bytes()
        for name in raw:
            key=name+'!A1:B3';self.table(key,2);result=self.run_app(True);self.assertEqual(result.verification['state'],'passed',result.verification)
            self.out.unlink();self.corruption_blocks(key)
        self.assertEqual(before,self.book.read_bytes())

    def test_multilevel_siblings_spans_and_optional_closures(self):
        self.book=self.root/'Tree.xls'
        page='<table><tr><th>Item<th>Amount<tr><td rowspan="2">Root<table><tr><td>Child<table><tr><td>Leaf<td>7</table><td>12</table> Done<td>46<tr><td>10</table><table><tr><td>Sibling<td>99</table>'
        self.book.write_text(page,encoding='utf-8');raw=sources.read(self.book);prod=readers.load(str(self.book))
        self.assertEqual(list(raw),['Tree','Tree 2','Tree 3','Tree 4']);self.assertEqual(raw['Tree'].get('A2'),'Root [Tree 2] Done')
        self.assertEqual(raw['Tree'].anchor(3,1),'A2');self.assertEqual(raw['Tree'].get('B3'),'10')
        self.assertEqual(raw['Tree 2'].get('A1'),'Child [Tree 3]');self.assertEqual(raw['Tree 3'].get('B1'),'7');self.assertEqual(raw['Tree 4'].get('B1'),'99')
        for name,sheet in raw.items():
            for ref,value in sheet.cells.items():self.assertEqual(str(prod[name][ref].value),str(value),(name,ref))

    def test_nested_marker_and_preserved_tail_mutations_block_save(self):
        self.book=self.root/'Nested.xls';self.book.write_text(NESTED,encoding='utf-8');self.table('Nested!A1:B3',2)
        original=core.ExcelSource.range_cells
        def changed(source,key):
            grid,where,issues=original(source,key);grid[1][0]=('Before After','Before After','General');return grid,where,issues
        with patch.object(core.ExcelSource,'range_cells',changed):
            self.assertTrue(self.run_app().blocked)
            with self.assertRaises(core.SlideFillError):self.run_app(True)
        self.corruption_blocks('Nested!A1:B3',2,1)

    def test_saved_nested_value_and_marker_corruption_cannot_bypass_guard(self):
        self.book=self.root/'Nested.xls';self.book.write_text(NESTED,encoding='utf-8');self.table('Nested!A1:B3',2)
        for old,new in [(b'>46<',b'>999<'),(b'[Nested 2]',b'')]:
            def corrupt(raw):
                self.assertIn(old,raw);return raw.replace(old,new)
            self.blocked_by(lambda blob:fixture.alter(blob,'ppt/slides/slide1.xml',corrupt))

    def test_invalid_nested_placement_unclosed_tables_and_depth_are_explicit(self):
        for page in ['<table><table><tr><td>1</td></tr></table></table>','<table><tr><td>missing end',('<table><tr><td>'*33)+('</td></tr></table>'*33)]:
            self.book=self.root/'invalid.xls';self.book.write_text(page,encoding='utf-8')
            with self.assertRaises(sources.Unsupported):sources.read(self.book)
            with self.assertRaises(ValueError):readers.load(str(self.book))


if __name__=='__main__':unittest.main(verbosity=2)
