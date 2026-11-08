"""Read-only MS-CFB container, implemented from the public format specification."""
from pathlib import Path
import struct
from .recipes import Unsupported

END=0xfffffffe
FREE=0xffffffff
MAX_BYTES=128*1024*1024


def workbook_stream(path):
    if Path(path).stat().st_size>MAX_BYTES:raise Unsupported('compound input exceeds byte budget')
    data=Path(path).read_bytes()
    if len(data)<512 or len(data)>MAX_BYTES or data[:8]!=bytes.fromhex('d0cf11e0a1b11ae1'):
        raise Unsupported('not a supported bounded compound workbook')
    u16=lambda at:struct.unpack_from('<H',data,at)[0]
    u32=lambda at:struct.unpack_from('<I',data,at)[0]
    version=u16(26);shift=u16(30)
    if u16(28)!=0xfffe or (version,shift) not in ((3,9),(4,12)) or u16(32)!=6:
        raise Unsupported('unsupported compound sector geometry')
    size=1<<shift;count=len(data)//size-1
    def sector(index):
        if not 0<=index<count:raise Unsupported('compound sector outside file')
        return data[(index+1)*size:(index+2)*size]
    def words(raw):return list(struct.unpack('<'+'I'*(len(raw)//4),raw))
    fat_sectors=[v for v in struct.unpack_from('<109I',data,76) if v!=FREE]
    difat=u32(68);seen=set()
    if u32(72)>count or u32(44)>count:raise Unsupported('compound allocation count exceeds file')
    for _ in range(u32(72)):
        if difat in seen:raise Unsupported('cyclic compound DIFAT')
        seen.add(difat);entries=words(sector(difat));fat_sectors.extend(v for v in entries[:-1] if v!=FREE);difat=entries[-1]
    if len(fat_sectors)!=u32(44) or len(set(fat_sectors))!=len(fat_sectors):raise Unsupported('invalid compound FAT declaration')
    fat=[v for index in fat_sectors for v in words(sector(index))]
    def chain(first,table,fetch,limit):
        parts=[];visited=set();current=first
        while current!=END:
            if current in visited or not 0<=current<min(len(table),limit):raise Unsupported('invalid/cyclic compound chain')
            visited.add(current);parts.append(fetch(current));current=table[current]
        return b''.join(parts)
    directory=chain(u32(48),fat,sector,count)
    entries=[]
    for at in range(0,len(directory)-127,128):
        raw=directory[at:at+128];length=struct.unpack_from('<H',raw,64)[0]
        if length>64 or length%2:raise Unsupported('invalid compound directory name')
        name=raw[:max(0,length-2)].decode('utf-16-le') if length else ''
        left,right,child=struct.unpack_from('<III',raw,68)
        begin=struct.unpack_from('<I',raw,116)[0];length=struct.unpack_from('<Q',raw,120)[0]
        if version==3:length&=0xffffffff
        entries.append((name,raw[66],left,right,child,begin,length))
    if not entries or entries[0][1]!=5:raise Unsupported('compound root missing')
    # Only root-level streams; do not confuse an embedded Workbook with this file.
    pending=[entries[0][4]];seen=set();books=[]
    while pending:
        index=pending.pop()
        if index==FREE:continue
        if index in seen or not 0<=index<len(entries):raise Unsupported('invalid compound directory tree')
        seen.add(index);entry=entries[index];pending.extend(entry[2:4])
        if entry[1]==2 and entry[0].casefold() in ('workbook','book'):books.append(entry)
    if len(books)!=1:raise Unsupported('unique root workbook stream missing')
    entry=books[0];length=entry[6]
    if length>MAX_BYTES:raise Unsupported('compound workbook exceeds byte budget')
    if length<u32(56):
        if u32(56)!=4096:raise Unsupported('unsupported compound mini-stream cutoff')
        mini_raw=chain(u32(60),fat,sector,count) if u32(64) else b''
        mini_fat=words(mini_raw)
        root=entries[0];mini_stream=chain(root[5],fat,sector,count)[:root[6]]
        result=chain(entry[5],mini_fat,lambda n:mini_stream[n*64:(n+1)*64],len(mini_stream)//64)
    else:result=chain(entry[5],fat,sector,count)
    if len(result)<length:raise Unsupported('truncated compound workbook stream')
    return result[:length]
