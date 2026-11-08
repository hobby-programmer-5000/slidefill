"""Read XLSX evidence with zip/XML only; no openpyxl or production reader."""
from dataclasses import dataclass
import hashlib
import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def position(ref):
    m = re.fullmatch(r'([A-Z]+)([1-9]\d*)', ref)
    if not m:
        raise ValueError('invalid cell reference')
    column = 0
    for ch in m[1]:
        column = column * 26 + ord(ch) - 64
    return int(m[2]), column


def address(row, col):
    letters = ''
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return f'{letters}{row}'


@dataclass
class Sheet:
    cells: dict
    merges: list
    formats: dict
    types: dict = None
    epoch1904: bool = False
    row_levels: dict = None
    condition_values: dict = None  # independent delimited-text normalization for filters

    def get(self, ref):
        return self.cells.get(ref)

    def anchor(self, r, c):
        for top, left, bottom, right in self.merges:
            if top <= r <= bottom and left <= c <= right:
                return address(top, left)
        return address(r, c)


def workbook(path):
    with ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(n.itertext()) if not n.findall('.//' + S + 't') else ''.join(t.text or '' for t in n.findall('.//' + S + 't'))
                      for n in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(S + 'si')]
        styles = {0: 'General'}
        if 'xl/styles.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/styles.xml'))
            custom = {int(n.get('numFmtId')): n.get('formatCode') for n in root.findall('.//' + S + 'numFmt')}
            xfs = root.find(S + 'cellXfs')
            if xfs is not None:
                styles = {i: custom.get(int(n.get('numFmtId', '0')), {0:'General',1:'0',3:'#,##0',2:'0.00',4:'#,##0.00'}.get(int(n.get('numFmtId','0')), 'UNSUPPORTED')) for i,n in enumerate(xfs)}
        relations = {n.get('Id'): n.get('Target') for n in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')) if n.get('TargetMode') != 'External'}
        result = {}
        bookroot=ET.fromstring(z.read('xl/workbook.xml'))
        properties=bookroot.find(S+'workbookPr')
        epoch1904=properties is not None and properties.get('date1904') in ('1','true')
        for tab in bookroot.findall('.//' + S + 'sheet'):
            target = relations[tab.get(R + 'id')]
            member = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            root = ET.fromstring(z.read(member))
            cells, formats, merges, types = {}, {}, [], {}
            for n in root.findall('.//' + S + 'sheetData/' + S + 'row/' + S + 'c'):
                ref, typ = n.get('r'), n.get('t')
                text = n.findtext(S + 'v')
                if typ == 's' and text is not None:
                    text = shared[int(text)]
                elif typ == 'inlineStr':
                    text = ''.join(t.text or '' for t in n.findall('.//' + S + 't'))
                elif typ == 'b':
                    text = True if text == '1' else False
                elif typ == 'e':
                    text = 'ERROR:' + str(text)
                cells[ref] = text
                types[ref] = typ or 'n'
                formats[ref] = styles.get(int(n.get('s','0')), 'UNSUPPORTED')
            for n in root.findall('.//' + S + 'mergeCell'):
                a,b = n.get('ref').split(':');top,left=position(a);bottom,right=position(b)
                merges.append((top,left,bottom,right))
            levels={int(n.get('r')):int(n.get('outlineLevel','0')) for n in root.findall('.//'+S+'sheetData/'+S+'row') if n.get('r')}
            if any(not 0<=v<=7 for v in levels.values()):raise ValueError('invalid row outline level')
            result[tab.get('name')] = Sheet(cells,merges,formats,types,epoch1904,row_levels=levels)
    return result


def digest(path):
    h = hashlib.sha256()
    with open(path,'rb') as f:
        for part in iter(lambda:f.read(1024*1024),b''):
            h.update(part)
    return h.hexdigest()
