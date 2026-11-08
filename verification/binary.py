"""Independent bounded BIFF5/8 and XLSB evidence. No xlrd/calamine/app imports.

Format facts: Microsoft MS-XLS, MS-XLSB, MS-CFB. Formula tokens are not executed.
"""
import math
import codecs
from pathlib import Path
import posixpath
import struct
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from .compound import workbook_stream
from .rawxlsx import Sheet,address
from .recipes import Unsupported

MAX_CELLS=2_000_000
ERRORS={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'}
BUILTINS={0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%'}


def numeric(value):
    if not math.isfinite(value):raise Unsupported('nonfinite binary numeric value')
    return str(int(value)) if value.is_integer() else repr(value)


def rk(raw):
    value=struct.unpack('<i',raw)[0]>>2 if raw[0]&2 else struct.unpack('<d',b'\x00'*4+bytes([raw[0]&0xfc])+raw[1:])[0]
    return numeric(float(value)/100 if raw[0]&1 else float(value))


class Strings:
    """BIFF Unicode with CONTINUE encoding flags only inside character data."""
    def __init__(self,parts):self.parts=parts;self.part=0;self.at=0
    def take(self,n):
        out=bytearray()
        while n:
            if self.part>=len(self.parts):raise Unsupported('truncated BIFF string')
            available=len(self.parts[self.part])-self.at
            if not available:self.part+=1;self.at=0;continue
            size=min(n,available);out.extend(self.parts[self.part][self.at:self.at+size]);self.at+=size;n-=size
        return bytes(out)
    def text(self):
        count=struct.unpack('<H',self.take(2))[0];flags=self.take(1)[0]
        rich=struct.unpack('<H',self.take(2))[0] if flags&8 else 0
        extra=struct.unpack('<I',self.take(4))[0] if flags&4 else 0
        wide=bool(flags&1);out=bytearray()
        while count:
            if self.part>=len(self.parts):raise Unsupported('truncated BIFF characters')
            if self.at==len(self.parts[self.part]):
                self.part+=1;self.at=0;wide=bool(self.take(1)[0]&1)
            width=2 if wide else 1
            available=(len(self.parts[self.part])-self.at)//width
            if not available:raise Unsupported('split BIFF Unicode code unit')
            n=min(count,available);raw=self.take(n*width)
            out.extend(raw if wide else b''.join(bytes((v,0)) for v in raw));count-=n
        self.take(rich*4+extra)
        return out.decode('utf-16-le')


def biff_records(data,start=0):
    at=start
    while at+4<=len(data):
        ident,size=struct.unpack_from('<HH',data,at);end=at+4+size
        if end>len(data):raise Unsupported('truncated BIFF record')
        yield ident,data[at+4:end];at=end
        if ident==0x000a:return
    raise Unsupported('BIFF EOF record missing')


def byte_text(body,at,codec,width=2):
    count=int.from_bytes(body[at:at+width],'little');end=at+width+count
    if at+width>len(body) or end>len(body):raise Unsupported('truncated legacy BIFF byte string')
    try:return body[at+width:end].decode(codec),end
    except UnicodeError as exc:raise Unsupported('legacy BIFF text does not match its codepage') from exc


def rich_text(body,version,codec):
    if version==0x0500:
        text,end=byte_text(body,6,codec);width=1;run_width=2
    else:
        if len(body)<9 or body[8]&~1:raise Unsupported('unsupported BIFF8 RSTRING flags')
        cursor=Strings([body[6:]]);text=cursor.text();end=6+cursor.at;width=2;run_width=4
    if end+width>len(body):raise Unsupported('truncated BIFF rich-text run count')
    count=int.from_bytes(body[end:end+width],'little');at=end+width
    if at+count*run_width!=len(body):raise Unsupported('truncated or trailing BIFF rich-text runs')
    last=-1;limit=len(text.encode('utf-16-le'))//2 if version==0x0600 else len(text)
    for i in range(count):
        index=int.from_bytes(body[at+i*run_width:at+i*run_width+width],'little')
        if not last<index<=limit:raise Unsupported('invalid BIFF rich-text run ordering')
        last=index
    return text


def legacy_codec(records):
    codepages=[struct.unpack_from('<H',body)[0] for ident,body in records if ident==0x0042]
    if len(set(codepages))>1:raise Unsupported('conflicting BIFF codepages')
    if not codepages:return 'ascii'  # no guessed locale for non-ASCII bytes
    codec={10000:'mac_roman',32768:'mac_roman',32769:'cp1252'}.get(codepages[0],'cp'+str(codepages[0]))
    try:codecs.lookup(codec)
    except LookupError as exc:raise Unsupported('unsupported BIFF codepage') from exc
    return codec


def early_sheet(records,version):
    """Single raw BIFF2/3/4 worksheet, not BIFF4W/workspace/macro containers."""
    if struct.unpack_from('<H',records[0][1],2)[0]!=0x10:raise Unsupported('legacy BIFF workbook/workspace is outside single-sheet contract')
    codec=legacy_codec(records);codes={};styles=[];epoch=False
    for ident,body in records:
        if ident==0x002f:raise Unsupported('encrypted legacy BIFF worksheet')
        if ident==0x0022:epoch=bool(struct.unpack_from('<H',body)[0])
        elif ident in (0x001e,0x041e):codes[len(codes)]=byte_text(body,2 if ident==0x041e else 0,codec,1)[0]
        elif ident in (0x0043,0x0243,0x0443):styles.append(body[1]&63 if version==2 else body[1])
    cells={};types={};formats={};formulas={};pending=None
    label,number,boolean,formula,string=(4,3,5,6,7) if version==2 else (0x204,0x203,0x205,0x206 if version==3 else 0x406,0x207)
    for ident,body in records:
        if ident==string:
            if pending is None:raise Unsupported('orphan legacy BIFF formula string')
            cells[pending]=byte_text(body,0,codec,1 if version==2 else 2)[0];types[pending]='s';pending=None
            if len(cells)>MAX_CELLS:raise Unsupported('legacy BIFF cell budget exceeded')
            continue
        if ident not in (label,number,boolean,formula,2 if version==2 else 0x27e):continue
        if pending is not None:raise Unsupported('legacy BIFF formula string missing')
        r,c=struct.unpack_from('<HH',body);at=7 if version==2 else 6
        if r>=16384 or c>=256:raise Unsupported('legacy BIFF cell outside grid')
        if len(body)<at:raise Unsupported('truncated legacy BIFF cell header')
        fmt=body[5]&63 if version==2 else styles[struct.unpack_from('<H',body,4)[0]] if struct.unpack_from('<H',body,4)[0]<len(styles) else 0
        ref=address(r+1,c+1)
        if ident==formula:formulas[ref]={'kind':'binary'}
        if ref in cells:raise Unsupported('duplicate legacy BIFF value cell')
        formats[ref]=codes.get(fmt,BUILTINS.get(fmt,'UNSUPPORTED'))
        if ident==label:value=byte_text(body,at,codec,1 if version==2 else 2)[0];kind='s'
        elif ident==boolean:value=ERRORS.get(body[at],'#ERROR') if body[at+1] else bool(body[at]);kind='e' if body[at+1] else 'b'
        elif ident==2:value=str(struct.unpack_from('<H',body,at)[0]);kind='n'
        elif ident==0x27e:value=rk(body[at:at+4]);kind='n'
        elif ident==number:value=numeric(struct.unpack_from('<d',body,at)[0]);kind='n'
        else:
            cached=body[at:at+8]
            if cached[6:8]!=b'\xff\xff':value=numeric(struct.unpack('<d',cached)[0]);kind='n'
            elif version==2:raise Unsupported('legacy BIFF2 special formula cache is unsupported')
            elif cached[0]==0:pending=ref;continue
            elif cached[0]==1:value=bool(cached[2]);kind='b'
            elif cached[0]==2:value=ERRORS.get(cached[2],'#ERROR');kind='e'
            elif cached[0]==3:value='';kind='s'
            else:raise Unsupported('unknown legacy BIFF formula cache')
        cells[ref]=value;types[ref]=kind
        if len(cells)>MAX_CELLS:raise Unsupported('legacy BIFF cell budget exceeded')
    if pending is not None:raise Unsupported('legacy BIFF formula string missing')
    return {'Sheet 1':Sheet(cells,[],formats,types,epoch,formulas=formulas)}


def xls(path):
    if Path(path).stat().st_size>128*1024*1024:raise Unsupported('BIFF input exceeds byte budget')
    with open(path,'rb') as stream:signature=stream.read(8)
    data=workbook_stream(path) if signature==bytes.fromhex('d0cf11e0a1b11ae1') else Path(path).read_bytes()
    records=list(biff_records(data));strings=[];tabs=[];styles=[];codes={};epoch=False
    if records and records[0][0] in (9,0x209,0x409):return early_sheet(records,{9:2,0x209:3,0x409:4}[records[0][0]])
    if not records or records[0][0]!=0x0809 or len(records[0][1])<4:raise Unsupported('unsupported BIFF workbook header')
    version=struct.unpack_from('<H',records[0][1])[0]
    if version not in (0x0500,0x0600):raise Unsupported('BIFF workbook version outside supported contract')
    codec=legacy_codec(records) if version==0x0500 else 'ascii'
    index=0
    while index<len(records):
        ident,body=records[index];index+=1
        if ident==0x002f:raise Unsupported('encrypted BIFF workbook')
        if ident==0x0085:
            offset=struct.unpack_from('<I',body)[0]
            if body[5]!=0:continue  # chart/macro sheets are not value worksheets
            if version==0x0500:name,_end=byte_text(body,6,codec,1)
            else:
                count,flags=body[6:8];name=body[8:8+count*(2 if flags&1 else 1)].decode('utf-16-le' if flags&1 else 'latin1')
            tabs.append((name,offset))
        elif ident==0x0022:epoch=bool(struct.unpack_from('<H',body)[0])
        elif ident==0x00e0:styles.append(struct.unpack_from('<H',body,2)[0])
        elif ident==0x041e:codes[struct.unpack_from('<H',body)[0]]=byte_text(body,2,codec,1)[0] if version==0x0500 else Strings([body[2:]]).text()
        elif ident==0x00fc:
            parts=[body]
            while index<len(records) and records[index][0]==0x003c:parts.append(records[index][1]);index+=1
            cursor=Strings(parts);_total,unique=struct.unpack('<II',cursor.take(8))
            if unique>MAX_CELLS:raise Unsupported('BIFF string count exceeds budget')
            strings=[cursor.text() for _ in range(unique)]
    result={};total=0
    for name,offset in tabs:
        if name in result:raise Unsupported('duplicate BIFF worksheet name')
        cells={};types={};formats={};merges=[];formulas={};pending=None
        for ident,body in biff_records(data,offset):
            if ident==0x00e5:
                count=struct.unpack_from('<H',body)[0]
                if len(body)!=2+8*count:raise Unsupported('invalid BIFF merge table')
                for i in range(count):
                    top,bottom,left,right=struct.unpack_from('<HHHH',body,2+i*8);merges.append((top+1,left+1,bottom+1,right+1))
                continue
            if ident==0x0207:
                if pending is None:raise Unsupported('orphan BIFF formula string')
                cells[pending]=byte_text(body,0,codec)[0] if version==0x0500 else Strings([body]).text();types[pending]='s';pending=None;total+=1
                if total>MAX_CELLS:raise Unsupported('BIFF cell count exceeds budget')
                continue
            if ident not in (0x0203,0x027e,0x00bd,0x00fd,0x0204,0x0205,0x0006,0x00d6):
                continue
            if pending is not None:raise Unsupported('BIFF formula string result missing')
            r,c,xf=struct.unpack_from('<HHH',body)
            if r>=(16384 if version==0x0500 else 65536) or c>=256:raise Unsupported('BIFF value outside version grid')
            if address(r+1,c+1) in cells:raise Unsupported('duplicate BIFF value cell')
            values=[]
            if ident==0x00bd:
                r,c=struct.unpack_from('<HH',body);last=struct.unpack_from('<H',body,len(body)-2)[0]
                if len(body)!=6+6*(last-c+1):raise Unsupported('invalid BIFF packed numeric row')
                for i in range(last-c+1):
                    style=struct.unpack_from('<H',body,4+i*6)[0];values.append((c+i,style,rk(body[6+i*6:10+i*6]),'n'))
            elif ident==0x0203:values=[(c,xf,numeric(struct.unpack_from('<d',body,6)[0]),'n')]
            elif ident==0x027e:values=[(c,xf,rk(body[6:10]),'n')]
            elif ident==0x00fd:values=[(c,xf,strings[struct.unpack_from('<I',body,6)[0]],'s')]
            elif ident==0x0204:values=[(c,xf,byte_text(body,6,codec)[0] if version==0x0500 else Strings([body[6:]]).text(),'s')]
            elif ident==0x00d6:values=[(c,xf,rich_text(body,version,codec),'s')]
            elif ident==0x0205:values=[(c,xf,ERRORS.get(body[6],'#ERROR') if body[7] else bool(body[6]),'e' if body[7] else 'b')]
            else:
                formulas[address(r+1,c+1)]={'kind':'binary'}
                cached=body[6:14]
                if cached[6:8]!=b'\xff\xff':values=[(c,xf,numeric(struct.unpack('<d',cached)[0]),'n')]
                elif cached[0]==0:pending=address(r+1,c+1);formats[pending]=codes.get(styles[xf],BUILTINS.get(styles[xf],'UNSUPPORTED'));continue
                elif cached[0]==1:values=[(c,xf,bool(cached[2]),'b')]
                elif cached[0]==2:values=[(c,xf,ERRORS.get(cached[2],'#ERROR'),'e')]
                elif cached[0]==3:values=[(c,xf,'','s')]
                else:raise Unsupported('unknown BIFF formula cache type')
            for col,style,value,kind in values:
                ref=address(r+1,col+1)
                if ref in cells:raise Unsupported('duplicate BIFF value cell')
                cells[ref]=value;types[ref]=kind
                fmt=styles[style] if style<len(styles) else 0;formats[ref]=codes.get(fmt,BUILTINS.get(fmt,'UNSUPPORTED'))
                total+=1
                if total>MAX_CELLS:raise Unsupported('BIFF cell count exceeds budget')
        if pending is not None:raise Unsupported('BIFF formula string cache missing')
        result[name]=Sheet(cells,merges,formats,types,epoch,formulas=formulas)
    if not result:raise Unsupported('BIFF worksheet streams missing')
    return result


def xlsb_records(data):
    at=0
    while at<len(data):
        ident=0
        for shift in (0,7):
            if at>=len(data):raise Unsupported('truncated XLSB identifier')
            byte=data[at];at+=1;ident|=(byte&127)<<shift
            if byte<128:break
        else:raise Unsupported('invalid XLSB identifier')
        size=0
        for shift in (0,7,14,21):
            if at>=len(data):raise Unsupported('truncated XLSB size')
            byte=data[at];at+=1;size|=(byte&127)<<shift
            if byte<128:break
        else:raise Unsupported('invalid XLSB length')
        if at+size>len(data):raise Unsupported('truncated XLSB body')
        yield ident,data[at:at+size];at+=size


def wide(body,at):
    count=struct.unpack_from('<I',body,at)[0];end=at+4+2*count
    if count>MAX_CELLS or end>len(body):raise Unsupported('invalid XLSB wide string')
    return body[at+4:end].decode('utf-16-le'),end


def part(z,name):
    if z.getinfo(name).file_size>128*1024*1024:raise Unsupported('binary ZIP part exceeds byte budget')
    return z.read(name)


def xlsb(path):
    with ZipFile(path) as z:
        rels={n.get('Id'):n.get('Target') for n in ET.fromstring(part(z,'xl/_rels/workbook.bin.rels')) if n.get('TargetMode')!='External'}
        records=list(xlsb_records(part(z,'xl/workbook.bin')));tabs=[];epoch=False;shared=[];codes={};styles=[];inside=False
        for ident,body in records:
            if ident==153:epoch=bool(body[0]&1)
            elif ident==156:
                if struct.unpack_from('<I',body,8)[0]==0xffffffff:continue
                relation,at=wide(body,8);name,_=wide(body,at);target=rels.get(relation)
                if target is None:raise Unsupported('XLSB worksheet relationship missing/external')
                tabs.append((name,target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)))
        if 'xl/sharedStrings.bin' in z.namelist():
            for ident,body in xlsb_records(part(z,'xl/sharedStrings.bin')):
                if ident==19:shared.append(wide(body,1)[0])
                if len(shared)>MAX_CELLS:raise Unsupported('XLSB shared strings exceed budget')
        if 'xl/styles.bin' in z.namelist():
            for ident,body in xlsb_records(part(z,'xl/styles.bin')):
                if ident==44:codes[struct.unpack_from('<H',body)[0]]=wide(body,2)[0]
                elif ident in (617,618):inside=ident==617
                elif ident==47 and inside:styles.append(struct.unpack_from('<H',body,2)[0])
        result={};total=0
        for name,member in tabs:
            if name in result:raise Unsupported('duplicate XLSB sheet name')
            cells={};types={};formats={};merges=[];formulas={};row=None;column=-1
            for ident,body in xlsb_records(part(z,member)):
                if ident==0:row=struct.unpack_from('<I',body)[0];column=-1;continue
                if ident==176:
                    top,bottom,left,right=struct.unpack_from('<IIII',body);merges.append((top+1,left+1,bottom+1,right+1));continue
                if ident not in range(1,19) and ident!=62:continue
                if row is None or row>=1048576:raise Unsupported('XLSB cell without a valid row')
                previous=column
                if 12<=ident<=18:
                    if previous<0:raise Unsupported('XLSB short cell needs a preceding full cell in its row')
                    column+=1;style=struct.unpack_from('<I',body)[0];at=4;kind=ident-11
                else:
                    column,style=struct.unpack_from('<II',body);at=8;kind=ident
                if column<=previous:raise Unsupported('XLSB cells are not in increasing column order')
                if column>=16384:raise Unsupported('XLSB column outside grid')
                if kind<=7:
                    size={1:0,2:4,3:1,4:1,5:8,7:4}.get(kind)
                    if size is not None and len(body)!=at+size:raise Unsupported('invalid XLSB value record length')
                if kind==1:continue
                if kind==2:value=rk(body[at:at+4]);typ='n'
                elif kind in (3,11):value=ERRORS.get(body[at],'#ERROR');typ='e'
                elif kind in (4,10):value=bool(body[at]);typ='b'
                elif kind in (5,9):value=numeric(struct.unpack_from('<d',body,at)[0]);typ='n'
                elif kind in (6,8):
                    value,end=wide(body,at);typ='s'
                    if kind==6 and end!=len(body):raise Unsupported('invalid XLSB string record length')
                elif kind==7:value=shared[struct.unpack_from('<I',body,at)[0]];typ='s'
                elif kind==62:value=wide(body,at+1)[0];typ='s'
                else:raise Unsupported('unsupported XLSB cell record')
                ref=address(row+1,column+1)
                if kind in (8,9,10,11):formulas[ref]={'kind':'binary'}
                if ref in cells:raise Unsupported('duplicate XLSB value cell')
                cells[ref]=None if typ=='s' and value=='' else value;types[ref]=typ;fmt=styles[style&0xffffff] if style&0xffffff<len(styles) else 0
                formats[ref]=codes.get(fmt,BUILTINS.get(fmt,'UNSUPPORTED'));total+=1
                if total>MAX_CELLS:raise Unsupported('XLSB cells exceed budget')
            result[name]=Sheet(cells,merges,formats,types,epoch,formulas=formulas)
        if not result:raise Unsupported('XLSB worksheets missing')
        return result
