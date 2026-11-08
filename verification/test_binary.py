"""Synthetic binary goldens and real app guards, without requiring Excel."""
from decimal import Decimal,ROUND_HALF_UP,localcontext
from fractions import Fraction
import io
import datetime as dt
import json
from pathlib import Path
import struct
import unittest
from unittest.mock import patch
from zipfile import ZipFile,ZIP_DEFLATED
import openpyxl
import core,numfmt,layouts,suggest
from pptx import Presentation
from pptx.util import Inches
from . import binary,compound,kernel,recipes,sources,test_automatic as fixture
from .test_coverage import ods


def biff(ident,body=b''):return struct.pack('<HH',ident,len(body))+body
def unicode_text(text):
    encoded=text.encode('utf-16-le');return struct.pack('<HB',len(encoded)//2,1)+encoded


def compound_file(stream,mini=False):
    """Tiny independent golden writer, no copied compound/parser code."""
    end,free=0xfffffffe,0xffffffff
    header=bytearray(512);header[:8]=bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH',header,24,62,3,0xfffe,9,6)
    def entry(name,typ,start,length,child=free):
        raw=bytearray(128);encoded=(name+'\0').encode('utf-16-le');raw[:len(encoded)]=encoded
        struct.pack_into('<HBBIII',raw,64,len(encoded),typ,1,free,free,child)
        struct.pack_into('<IQ',raw,116,start,length);return raw
    if mini:
        assert len(stream)<=512
        n=(len(stream)+63)//64
        directory=entry('Root Entry',5,2,n*64,1)+entry('Workbook',2,0,len(stream))+bytes(256)
        sectors=[directory,struct.pack('<128I',*[i+1 if i<n-1 else end if i==n-1 else free for i in range(128)]),stream.ljust(512,b'\0')]
        fat_index=3;directory_index=0;mini_index=1
        fat=[end,end,end,0xfffffffd]+[free]*124
        struct.pack_into('<IIIIII',header,56,4096,mini_index,1,end,0,fat_index)
    else:
        stream=stream.ljust(max(4096,((len(stream)+511)//512)*512),b'\0');n=len(stream)//512
        sectors=[stream[i:i+512] for i in range(0,len(stream),512)]
        directory_index=n;fat_index=n+1
        sectors.append(entry('Root Entry',5,end,0,1)+entry('Workbook',2,0,len(stream))+bytes(256))
        fat=[i+1 for i in range(n-1)]+[end,end,0xfffffffd]+[free]*(128-n-2)
        struct.pack_into('<IIIIII',header,56,4096,end,0,end,0,fat_index)
    struct.pack_into('<III',header,40,0,1,directory_index)
    struct.pack_into('<108I',header,80,*([free]*108))
    sectors.append(struct.pack('<128I',*fat))
    return bytes(header)+b''.join(sectors)


def xls_fixture(path):
    strings=['Item','Amount','A','서울'];sst=struct.pack('<II',len(strings),len(strings))+b''.join(unicode_text(s) for s in strings)
    bof=lambda kind:biff(0x0809,struct.pack('<HHHHII',0x0600,kind,0x0dbb,1996,0x41,6))
    font=struct.pack('<HHHHHBBBBBB',200,0,0x7fff,400,0,0,0,0,0,5,0)+b'Arial'
    prefix=bof(5)+biff(0x0042,struct.pack('<H',1200))+biff(0x0031,font)+biff(0x00e0,bytes(20))+biff(0x0022,b'\0\0')+biff(0x00fc,sst)
    bound=lambda offset:biff(0x0085,struct.pack('<IBBBB',offset,0,0,4,0)+b'Data')
    globals_=prefix+bound(0)+biff(10)
    globals_=prefix+bound(len(globals_))+biff(10)
    cell=lambda r,c:struct.pack('<HHH',r,c,0)
    body=bof(0x10)+biff(0x0200,struct.pack('<IIHHH',0,3,0,2,0))
    for r,c,index in [(0,0,0),(0,1,1),(1,0,2),(2,0,3)]:body+=biff(0x00fd,cell(r,c)+struct.pack('<I',index))
    body+=biff(0x027e,cell(1,1)+struct.pack('<I',(12<<2)|2))
    body+=biff(0x0203,cell(2,1)+struct.pack('<d',-2.5))+biff(10)
    path.write_bytes(compound_file(globals_+body))


def record(ident,body=b''):
    prefix=bytes([ident]) if ident<128 else bytes([(ident&127)|128,ident>>7])
    length=len(body);size=[]
    while True:
        size.append((length&127)|(128 if length>=128 else 0));length>>=7
        if not length:break
    return prefix+bytes(size)+body


def xlsb_fixture(path):
    wide=lambda s:struct.pack('<I',len(s.encode('utf-16-le'))//2)+s.encode('utf-16-le')
    head=lambda c:struct.pack('<II',c,0)
    workbook=record(131)+record(153,bytes(12))+record(143)+record(156,struct.pack('<II',0,1)+wide('rId1')+wide('Data'))+record(144)+record(132)
    sheet=record(129)+record(148,struct.pack('<IIII',0,2,0,1))+record(145)
    for row,items in [(0,[(0,'Item'),(1,'Amount')]),(1,[(0,'A'),(1,12)]),(2,[(0,'서울'),(1,-2.5)])]:
        sheet+=record(0,struct.pack('<IIHBBBIII',row,0,300,0,0,0,1,0,1))
        for col,value in items:
            sheet+=record(6,head(col)+wide(value)) if isinstance(value,str) else record(5,head(col)+struct.pack('<d',value))
    sheet+=record(146)+record(130)
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/></Types>')
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>')
        z.writestr('xl/workbook.bin',workbook)
        z.writestr('xl/_rels/workbook.bin.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.bin',sheet)


class BinaryCoverage(unittest.TestCase):
    setUp=fixture.AutomaticGuard.setUp
    tearDown=fixture.AutomaticGuard.tearDown
    write=fixture.AutomaticGuard.write
    make=fixture.AutomaticGuard.make
    run_app=fixture.AutomaticGuard.run_app

    def test_binary_sources_pass_app_and_independent_source_mutation_blocks(self):
        for ext,writer in [('.xls',xls_fixture),('.xlsb',xlsb_fixture)]:
            with self.subTest(ext=ext):
                self.book=self.root/('input'+ext);writer(self.book);self.make()
                report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
                self.assertEqual(sources.read(self.book)['Data'].get('B3'),'-2.5')
                original=core.ExcelSource.range_cells
                def wrong(source,key):
                    grid,where,issues=original(source,key);grid[1][1]=(999,'999','General');return grid,where,issues
                with patch.object(core.ExcelSource,'range_cells',wrong):
                    self.assertTrue(self.run_app().blocked)
                    with self.assertRaises(core.SlideFillError):self.run_app(True)

    def test_html_and_delimited_exports_named_xls_are_independently_checked(self):
        self.book=self.root/'Data.xls'
        self.book.write_text('<table><tr><th>Item</th><th>Amount</th></tr><tr><td rowspan="2">서울</td><td>12</td></tr><tr><td>34</td></tr></table>',encoding='utf-8')
        self.make();report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
        raw=sources.read(self.book)['Data'];self.assertEqual(raw.anchor(3,1),'A2');self.assertEqual(raw.get('B3'),'34')
        self.book.write_text('Item\tAmount\nA\t12\nB\t34\n',encoding='utf-16')
        self.assertEqual(self.run_app().verification['state'],'passed')

    def test_malformed_short_binary_cells_remain_unsupported(self):
        self.book=self.root/'input.xlsb';xlsb_fixture(self.book)
        stream=io.BytesIO()
        with ZipFile(self.book) as src,ZipFile(stream,'w',ZIP_DEFLATED) as dst:
            for name in src.namelist():
                raw=src.read(name)
                if name.endswith('sheet1.bin'):raw+=record(13,struct.pack('<I',0))
                dst.writestr(name,raw)
        self.book.write_bytes(stream.getvalue())
        with self.assertRaises(sources.Unsupported):sources.read(self.book)

    def test_compound_minifat_and_cyclic_chain(self):
        small=b'golden stream';self.book.write_bytes(compound_file(small,mini=True))
        self.assertEqual(compound.workbook_stream(self.book),small)
        raw=bytearray(compound_file(small,mini=True));struct.pack_into('<I',raw,512*4,0);self.book.write_bytes(raw)
        with self.assertRaises(sources.Unsupported):compound.workbook_stream(self.book)

    def test_biff_unicode_continuation_and_rk_sign_and_scale(self):
        first=struct.pack('<HB',4,1)+'가'.encode('utf-16-le')
        second=b'\0AB';third=b'\1'+'나'.encode('utf-16-le')
        self.assertEqual(binary.Strings([first,second,third]).text(),'가AB나')
        self.assertEqual(binary.rk(struct.pack('<I',((-7<<2)|2)&0xffffffff)),'-7')
        self.assertEqual(binary.rk(struct.pack('<I',(125<<2)|3)),'1.25')
        with self.assertRaises(sources.Unsupported):list(binary.xlsb_records(b'\x85'))

    def test_ods_time_date_and_cached_calculation_values(self):
        self.book=self.root/'Data.ods'
        ods(self.book,'<table:table-row><table:table-cell office:value-type="time" office:time-value="PT1H2M3.125S"/><table:table-cell office:value-type="float" office:value="46" table:formula="of:=12+34"/></table:table-row>')
        self.make('Data!A1:B4');report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)
        self.assertEqual(sources.read(self.book)['Data'].get('A4'),'01:02:03.125000')
        self.assertEqual(sources.read(self.book)['Data'].get('B4'),'46')

    def test_month_quarter_year_numeric_and_grouped_numeric_text(self):
        for labels in [['2026-01','2026-02'],['2026년 1월','2026년 2월'],['Jan 2026','Feb 2026'],['2026 Q1','2026 Q2'],[2025,2026],[20260101,20260102]]:
            with self.subTest(labels=labels):
                self.book=self.root/'data.xlsx';self.write([['Date','Amount'],[labels[0],'39 925'],[labels[1],12]])
                self.make('latest2:Data!A,B');report=self.run_app();self.assertEqual(report.verification['state'],'passed',report.verification)

    def test_explicit_row_outline_is_preserved_without_hierarchy_sums(self):
        wb=openpyxl.Workbook();ws=wb.active;ws.title='Groups'
        for row in [['Item','Amount'],['Parent',50],['Child A',12],['Child B',13]]:ws.append(row)
        ws.row_dimensions[3].outlineLevel=1;ws.row_dimensions[4].outlineLevel=1;wb.save(self.book)
        source=core.ExcelSource(str(self.book));views=layouts.discover(source.wb.active)
        self.assertEqual([v.kind for v in views],['outline']);self.assertEqual(len(views[0].rows),3)
        plan=suggest.plan_report(source,'en');self.assertIsNotNone(plan)
        suggest.make_report(str(self.template),'Groups',plan.sections)
        report=self.run_app(True,source);self.assertEqual(report.verification['state'],'passed',report.verification)
        prs=Presentation(self.out);tables=[s.table for slide in prs.slides for s in slide.shapes if s.has_table]
        self.assertEqual([[c.text for c in row.cells] for row in tables[0].rows][1:],[['Parent','0','50'],['Child A','1','12'],['Child B','1','13']])
        original=fixture.automatic.serialize
        def omit_note(prs):
            blob=original(prs)
            def changed(raw):return raw.replace(b'Grouping levels are source metadata, not verified business membership.',b'')
            return fixture.alter(blob,'ppt/slides/slide1.xml',changed)
        with patch.object(fixture.automatic,'serialize',side_effect=omit_note):self.assertTrue(self.run_app(source=source).blocked)

    def test_outline_certificate_requires_grouping_caveat(self):
        from .pipeline import generate
        from .check import check
        self.write([['Item','Amount'],['Parent',50],['Child A',12],['Child B',13]])
        wb=openpyxl.load_workbook(self.book);wb.active.row_dimensions[3].outlineLevel=1;wb.active.row_dimensions[4].outlineLevel=1;wb.save(self.book)
        folder=self.root/'certificate';self.assertEqual(generate(self.book,folder)['result'],'PASS')
        report=folder/'report.pptx';proof=json.loads((folder/'certificate.json').read_text('utf-8'))
        blob=fixture.alter(report.read_bytes(),'ppt/slides/slide1.xml',lambda raw:raw.replace(b'Grouping levels are source metadata, not verified business membership.',b''))
        report.write_bytes(blob);self.assertEqual(check(self.book,report,proof)['result'],'FAIL')

    def test_plain_labels_do_not_invent_row_grouping(self):
        self.write([['Item','Amount'],['Parent',50],['Child A',12],['Child B',13]])
        source=core.ExcelSource(str(self.book));self.assertFalse(any(v.kind=='outline' for v in layouts.discover(source.wb.active)))

    def test_date_typed_filter_is_reconstructed_from_source(self):
        self.write([['Date','Amount','GroupDate'],['2026-01-01',12,dt.datetime(2026,1,1)],['2026-01-02',34,dt.datetime(2026,1,2)]])
        self.make('latest2:Data!A,sum(B, C=2026)')
        self.assertEqual(self.run_app().verification['state'],'passed')

    def test_proved_rounding_matches_decimal_signed_ties_and_large_coefficients(self):
        for coefficient in range(-125,126):
            value=Decimal(coefficient)/100
            for places in (0,1,2,6):
                expected=value.quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_UP)
                self.assertEqual(kernel.round_decimal(value,places),expected)
        for text in ['-1.005','2.675','123456789012345678901234567890.555']:
            with localcontext() as ctx:
                ctx.prec=80
                self.assertEqual(kernel.round_decimal(text,2),Decimal(text).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
        with patch.object(kernel.verified,'RoundScaled',wraps=kernel.verified.RoundScaled) as method:
            self.assertEqual(numfmt.format_value(2.675,'0.00')[0],'2.68');self.assertTrue(method.called)


if __name__=='__main__':unittest.main(verbosity=2)
