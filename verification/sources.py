"""Independent source adapters; no production readers or spreadsheet libraries."""
import csv
import datetime as dt
from fractions import Fraction
import io
from pathlib import Path
import re
import struct
import xml.etree.ElementTree as ET
from zipfile import ZipFile, is_zipfile
from .rawxlsx import Sheet, address, workbook
from .recipes import Unsupported

T='{urn:oasis:names:tc:opendocument:xmlns:table:1.0}'
O='{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
P='{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
MAX_CELLS=2_000_000


def _title(name,taken):
    base=re.sub(r'[\\/*?:\[\]]',' ',name).strip()[:31] or 'Sheet'
    result=base;n=2
    while result in taken:result=f'{base[:28]}_{n}';n+=1
    return result


def _decode(raw):
    text=None
    for codec in (('utf-16',) if raw[:2] in (b'\xff\xfe',b'\xfe\xff') else ('utf-8-sig','cp949','utf-16')):
        try:text=raw.decode(codec);break
        except UnicodeError:continue
    if text is None:raise Unsupported('unrecognized delimited text encoding')
    return text


def _text_table(rows,name,merges):
    cells={};types={};formats={};conditions={};seen=0
    for r,row in enumerate(rows,1):
        for c,value in enumerate(row,1):
            seen+=1
            if seen>MAX_CELLS:raise Unsupported('delimited source exceeds independent cell budget')
            value=value.strip()
            if not value:continue
            ref=address(r,c);number=bool(re.fullmatch(r'-?(?:0|[1-9]\d*)(?:\.\d+)?|-?[1-9]\d{0,2}(?:,\d{3})+(?:\.\d+)?',value))
            cells[ref]=value.replace(',','') if number else value
            types[ref]='n' if number else 's'
            if number:conditions[ref]=str(float(value.replace(',',''))) if '.' in value else str(int(value.replace(',','')))
            formats[ref]=('#,##0.##' if '.' in value else '#,##0') if number and ',' in value else 'General'
    return {_title(name,{}):Sheet(cells,merges,formats,types,condition_values=conditions)}


def _text(path):
    text=_decode(Path(path).read_bytes())
    head=text[:4000].splitlines()[:5]
    marks=['\t',',',';','|'];counts={m:sum(line.count(m) for line in head) for m in marks}
    delimiter='\t' if Path(path).suffix.lower()=='.tsv' and counts['\t'] else max(marks,key=lambda m:(counts[m],m==','))
    return _text_table(csv.reader(io.StringIO(text),delimiter=delimiter),Path(path).stem,[])


def _paragraph(node):
    def words(element):
        value=element.text or ''
        for child in element:
            if child.tag==P+'s':value+=' '*int(child.get(P+'c','1'))
            elif child.tag==P+'tab':value+='\t'
            elif child.tag==P+'line-break':value+='\n'
            else:value+=words(child)
            value+=child.tail or ''
        return value
    return '\n'.join(words(p) for p in node.findall(P+'p'))


def _time(text):
    match=re.fullmatch(r'(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?',text or '')
    if not match or not any(match.groups()):raise Unsupported('unsupported ODS duration notation')
    negative,*parts=match.groups()
    if not any(value is not None for value in parts) or text.endswith('T'):raise Unsupported('unsupported ODS duration notation')
    days,hours,minutes,seconds=(Fraction(value or '0') for value in parts)
    micros=(days*86400+hours*3600+minutes*60+seconds)*1000000
    # Out-of-clock durations retain their source notation; never wrap a day,
    # discard the sign, or round finer precision into a different duration.
    if negative or micros.denominator!=1 or not 0<=micros<86400*1000000:return text
    micros=int(micros);whole,fraction=divmod(micros,1000000);h,rest=divmod(whole,3600);m,s=divmod(rest,60)
    return dt.time(h,m,s,fraction).isoformat()


def _ods(path):
    with ZipFile(path) as z:root=ET.fromstring(z.read('content.xml'))
    result={};budget=0
    for table in root.findall('.//'+O+'spreadsheet/'+T+'table'):
        cells={};types={};formats={};merges=[];formulas={};r=1
        # Descendant rows also cover explicit row/header groups.
        for row in table.iter(T+'table-row'):
            repeat=int(row.get(T+'number-rows-repeated','1'));c=1;items=[]
            if repeat<1:raise Unsupported('invalid ODS row repetition')
            for cell in row:
                if cell.tag not in (T+'table-cell',T+'covered-table-cell'):continue
                count=int(cell.get(T+'number-columns-repeated','1'))
                if count<1:raise Unsupported('invalid ODS column repetition')
                typ=cell.get(O+'value-type');value=None;kind='s'
                if typ in ('float','currency','percentage'):value=cell.get(O+'value');kind='n'
                elif typ=='boolean':
                    literal=cell.get(O+'boolean-value')
                    if literal not in ('true','false','1','0'):raise Unsupported('unsupported ODS boolean notation')
                    value=literal in ('true','1');kind='b'
                elif typ=='date':value=cell.get(O+'date-value');kind='d'
                elif typ=='time':value=_time(cell.get(O+'time-value'));kind='s'
                elif typ in (None,'string'):value=cell.get(O+'string-value',_paragraph(cell)) or None
                else:raise Unsupported('ODS value type outside independent reader contract')
                wide=int(cell.get(T+'number-columns-spanned','1'));high=int(cell.get(T+'number-rows-spanned','1'))
                if wide<1 or high<1:raise Unsupported('invalid ODS merge')
                formula=cell.get(T+'formula')
                if value is not None or formula is not None or wide>1 or high>1:
                    budget+=count*repeat
                    if budget>MAX_CELLS or c+count+wide-2>16384 or r+repeat+high-2>1048576:raise Unsupported('ODS source exceeds independent cell budget')
                    items.append((c,count,value,kind,wide,high,formula))
                c+=count
            # No materialization of huge trailing empty repeats.
            for rr in range(r,r+repeat) if items else ():
                for col,count,value,kind,wide,high,formula in items:
                    for cc in range(col,col+count):
                        ref=address(rr,cc)
                        if formula is not None:
                            formulas[ref]={'text':formula,'kind':'ods' if repeat==count==1 else 'ods-repeated'}
                        if value is not None:cells[ref]=value;types[ref]=kind;formats[ref]='General'
                        if wide>1 or high>1:merges.append((rr,cc,rr+high-1,cc+wide-1))
            r+=repeat
        result[_title(table.get(T+'name','Sheet'),result)]=Sheet(cells,merges,formats,types,formulas=formulas)
    if not result:raise Unsupported('ODS spreadsheet tables missing')
    return result


def read(path):
    ext=Path(path).suffix.lower()
    if ext in ('.csv','.tsv','.txt'):return _text(path)
    if ext=='.ods':return _ods(path)
    if ext=='.xlsb':
        from .binary import xlsb
        try:return xlsb(path)
        except (struct.error,IndexError,UnicodeError) as exc:raise Unsupported('binary source record could not be reconstructed') from exc
    if ext=='.xls':
        with open(path,'rb') as stream:signature=stream.read(8)
        if signature==bytes.fromhex('d0cf11e0a1b11ae1') or signature[:2] in (b'\x09\x00',b'\x09\x02',b'\x09\x04',b'\x09\x08'):
            from .binary import xls
            try:return xls(path)
            except (struct.error,IndexError,UnicodeError) as exc:raise Unsupported('binary source record could not be reconstructed') from exc
        if not is_zipfile(path):
            raw=Path(path).read_bytes()
            if re.search(rb'<\s*(?:table|html)',raw[:20000],re.I):
                from .htmlsource import read_html
                return read_html(path,_decode(raw))
            return _text(path)
    if ext in ('.xlsx','.xlsm','.xltx','.xltm') or ext=='.xls' and is_zipfile(path):return workbook(path)
    raise Unsupported('this input format has no independent source reader')
