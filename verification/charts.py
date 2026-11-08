"""Check generated category/XY caches and every referenced embedded XLSX cell."""
import datetime as dt
from io import BytesIO
import posixpath
import re
from fractions import Fraction
import xml.etree.ElementTree as ET
from .kernel import same_ints
from .rawxlsx import workbook,position,address
from .recipes import figure,near,Unsupported

C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def same_text(first,second):
    return same_ints(map(ord,str(first)),map(ord,str(second)))


def _cells(formula):
    sheet,mark,area=formula.rpartition('!')
    if not mark:raise ValueError('chart reference has no worksheet')
    if sheet.startswith("'") and sheet.endswith("'"):sheet=sheet[1:-1].replace("''", "'")
    bits=area.replace('$','').split(':');a,b=bits[0],bits[-1]
    top,left=position(a);bottom,right=position(b)
    if bottom<top or right<left or (bottom-top+1)*(right-left+1)>1000000:raise ValueError('invalid embedded chart range')
    return sheet,[address(r,c) for r in range(top,bottom+1) for c in range(left,right+1)]


def _serial(value):
    date=value.date() if isinstance(value,dt.datetime) else value
    result=(date-dt.date(1899,12,30)).days
    return result-1 if date<dt.date(1900,3,1) else result


def check(z,part,grid,mode,date_axis):
    errors=[];partial=[];unknown=object();root=ET.fromstring(z.read(part))
    ext=root.find(C+'externalData')
    if ext is None:return ['chart embedded workbook missing'],[]
    folder=posixpath.dirname(part)
    rels=ET.fromstring(z.read(folder+'/_rels/'+posixpath.basename(part)+'.rels'))
    rel=next((n for n in rels if n.get('Id')==ext.get(R+'id')),None)
    if rel is None or rel.get('TargetMode')=='External':return ['chart workbook is not embedded'],[]
    target=rel.get('Target');member=target.lstrip('/') if target.startswith('/') else posixpath.normpath(folder+'/'+target)
    book=workbook(BytesIO(z.read(member)))
    def expected_number(raw):
        try:return figure(raw)
        except Unsupported as exc:partial.append(str(exc));return unknown
    def cache_number(raw):
        if raw in (None,''):return None
        try:return Fraction(str(raw))
        except ValueError:raise ValueError('invalid numeric chart cache or embedded value')

    def check_ref(container,values,numeric):
        if container is None:errors.append('chart data field missing');return
        ref=container.find(C+('numRef' if numeric else 'strRef'))
        if ref is None:errors.append('chart data reference missing or wrong type');return
        cache=ref.find(C+('numCache' if numeric else 'strCache'))
        if cache is None:errors.append('chart cache missing');return
        count=cache.find(C+'ptCount')
        if count is None or int(count.get('val','-1'))!=len(values):errors.append('chart point count differs')
        points={};indices=[]
        for node in cache.findall(C+'pt'):
            i=int(node.get('idx'));indices.append(i);points[i]=node.findtext(C+'v')
        if len(indices)!=len(set(indices)) or any(i<0 or i>=len(values) for i in indices):errors.append('chart point indices invalid/duplicate')
        sheet,refs=_cells(ref.findtext(C+'f',''))
        if len(refs)!=len(values):errors.append('embedded chart reference length differs')
        for i,wanted in enumerate(values):
            cached=points.get(i)
            embedded=book[sheet].get(refs[i]) if i<len(refs) else None
            if numeric:
                if wanted is unknown:continue
                if wanted is None:
                    if cached not in (None,'') or embedded not in (None,''):errors.append('chart gap replaced with a number')
                elif not near(cache_number(cached),wanted) or not near(cache_number(embedded),wanted):errors.append(f'chart numeric cache/embedded mismatch at point{i}')
            else:
                if not same_text(cached or '',wanted or '') or not same_text(embedded or '',wanted or ''):errors.append(f'chart category/name mismatch at point{i}')

    saved=root.findall('.//'+C+'ser')
    expected_series=(len(grid[0])-1)//2 if mode=='bubble' else len(grid[0])-1
    if len(saved)!=expected_series:errors.append('chart series count differs')
    dates=date_axis and all(isinstance(line[0][0],dt.date) for line in grid[1:])
    categories=[_serial(line[0][0]) if dates else (line[0][1] or '') for line in grid[1:]]
    for index,node in enumerate(saved):
        col=1+index*(2 if mode=='bubble' else 1)
        if col>=len(grid[0]):break
        check_ref(node.find(C+'tx'),[grid[0][col][1] or ''],False)
        if mode in ('xy','bubble'):
            pairs=[tuple(expected_number(row[c][0]) for c in ([0,col,col+1] if mode=='bubble' else [0,col])) for row in grid[1:]]
            if any(unknown in pair for pair in pairs):partial.append('scatter source point interpretation is incomplete');continue
            pairs=[p for p in pairs if all(v is not None for v in p)]
            check_ref(node.find(C+'xVal'),[p[0] for p in pairs],True)
            check_ref(node.find(C+'yVal'),[p[1] for p in pairs],True)
            if mode=='bubble':check_ref(node.find(C+'bubbleSize'),[p[2] for p in pairs],True)
        else:
            check_ref(node.find(C+'cat'),categories,dates)
            check_ref(node.find(C+'val'),[expected_number(row[col][0]) for row in grid[1:]],True)
    return errors,partial
