"""Independent saved-PPTX XML checks; production code is never imported."""
from io import BytesIO
import posixpath
import re
from fractions import Fraction
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from .arithmetic import exact
from .oracle import expected_text, findings, context_labels, clean, grouping_caveat_visible
from .rawxlsx import workbook

A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
ACCOUNTING_INTEGER = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
ACCOUNTING_BLANK_ZERO = '_-* #,##0_-;\\-* #,##0_-;_-* ""_-;_-@_-'


def _equal_cell(sheet, item, shown, lang):
    expected=expected_text(sheet,item,lang)
    if item['op']=='cell' and expected!='':
        try:
            want=exact(expected)
        except ValueError:
            return shown==expected
        fmt=sheet.formats.get(item['refs'][0],'General')
        # Only lossless number formats in this initial verification contract.
        if fmt not in ('General','0','#,##0','0.00','#,##0.00','#,##0_ ',ACCOUNTING_INTEGER,ACCOUNTING_BLANK_ZERO):
            raise ValueError(f'number format outside independent checker contract: {fmt}')
        if fmt==ACCOUNTING_INTEGER and want==0 and shown.strip()=='-':
            return True  # only a numeric source zero under this exact zero section
        if fmt==ACCOUNTING_BLANK_ZERO and want==0 and shown.strip()=='':
            return True
        try:return exact(shown)==want
        except ValueError:return False
    return shown==expected


def _slides(z):
    rels={n.get('Id'):n.get('Target') for n in ET.fromstring(z.read('ppt/_rels/presentation.xml.rels')) if n.get('TargetMode')!='External'}
    root=ET.fromstring(z.read('ppt/presentation.xml'))
    return [posixpath.normpath('ppt/'+rels[n.get(R+'id')]) for n in root.findall('.//'+P+'sldId')]


def tables(path, book, certificate):
    problems=[];actual=[];paragraphs=[];chart_count=0
    with ZipFile(path) as z:
        for slide in _slides(z):
            root=ET.fromstring(z.read(slide))
            chart_count+=len(root.findall('.//'+C+'chart'))
            slide_words=[''.join(n.text or '' for n in p.findall('.//'+A+'t')) for p in root.findall('.//'+A+'p')]
            paragraphs.extend(slide_words)
            for frame in root.findall('.//'+P+'graphicFrame'):
                tbl=frame.find('.//'+A+'tbl')
                if tbl is None:continue
                prop=frame.find('.//'+P+'cNvPr')
                descr=prop.get('descr','') if prop is not None else ''
                rows=[[''.join(t.text or '' for t in tc.findall('.//'+A+'t')) for tc in tr.findall(A+'tc')] for tr in tbl.findall(A+'tr')]
                actual.append((descr,rows,slide,slide_words))
            if any('{{' in (n.text or '') for n in root.findall('.//'+A+'t')):
                problems.append(f'{slide}: unresolved placeholder')
    if len(actual)!=len(certificate['tables']):problems.append('saved table count differs from certificate')
    if chart_count:problems.append('unverified chart in structured-table report')
    for (descr,rows,slide,slide_words),page in zip(actual,certificate['tables']):
        if page['view'].startswith('outline-') and not grouping_caveat_visible(slide_words):problems.append('grouping interpretation caveat missing')
        if descr:problems.append(f'{slide}: unresolved table marker or unexpected alternative text')
        if not rows or rows[0]!=page['headers']:problems.append(f'{slide}: header differs')
        if len(rows)!=len(page['rows'])+1:problems.append(f'{slide}: missing/extra table rows')
        sheet=book[page['sheet']]
        for label in context_labels(sheet,page['view'],page['rows']):
            if not any(clean(label) in clean(words) for words in slide_words):
                problems.append(f'{slide}: source period/group label missing or changed: {label}')
        for i,(shown,wanted) in enumerate(zip(rows[1:],page['rows']),1):
            if len(shown)!=len(wanted):problems.append(f'{slide}: column count at row{i}')
            for j,(text,item) in enumerate(zip(shown,wanted),1):
                if not _equal_cell(sheet,item,text,page['lang']):
                    problems.append(f'{slide}: source/output mismatch at row{i}, column{j}, {item["refs"]}')
    for item in findings(book):
        relevant=[p for p in paragraphs if item['cell'] in p]
        if not relevant:problems.append(f'{item["cell"]}: source warning missing from saved report')
        elif item['kind']=='source_total_mismatch':
            values=[{Fraction(n) for n in re.findall(r'(?<![A-Za-z0-9_.])-?\d+(?:\.\d+)?(?![A-Za-z0-9_]|\.\d)',p)} for p in relevant]
            if not any({Fraction(item['source']),Fraction(item['recalculated'])} <= nums for nums in values):
                problems.append(f'{item["cell"]}: source/recalculated distinction missing')
    return problems


def chart(path, part, categories, series):
    """Compare expected values to chart caches AND its embedded XLSX.

    categories: list[str]; series: [(name, list[exact numeric strings])].
    These must come from an independent source oracle, not chart caches.
    """
    problems=[]
    def points(node,tag):
        if node is None:
            problems.append('chart category/value node missing')
            return []
        cache=node.find('.//'+C+tag)
        if cache is None:return []
        count=cache.find(C+'ptCount')
        values={int(n.get('idx')):n.findtext(C+'v') for n in cache.findall(C+'pt')}
        if count is None or int(count.get('val'))!=len(categories) or set(values)!=set(range(len(categories))):
            problems.append('chart cache count/index mismatch')
        return [values.get(i) for i in range(len(categories))]
    with ZipFile(path) as z:
        root=ET.fromstring(z.read(part))
        saved=root.findall('.//'+C+'ser')
        if len(saved)!=len(series):problems.append('chart series count mismatch')
        for node,(name,values) in zip(saved,series):
            tx=node.find(C+'tx')
            title=tx.findtext('.//'+C+'v') if tx is not None else None
            if title!=name:problems.append('chart series name mismatch')
            cats=points(node.find(C+'cat'),'strCache')
            if cats!=categories:problems.append('chart categories mismatch')
            nums=points(node.find(C+'val'),'numCache')
            try:
                if [exact(n) for n in nums]!=[exact(n) for n in values]:problems.append('chart numeric cache mismatch')
            except ValueError:problems.append('chart missing/nonnumeric cache')
        external=root.find(C+'externalData')
        if external is None:
            problems.append('chart embedded workbook missing')
        else:
            folder=posixpath.dirname(part)
            relpart=folder+'/_rels/'+posixpath.basename(part)+'.rels'
            relations={n.get('Id'):n for n in ET.fromstring(z.read(relpart))}
            relation=relations.get(external.get(R+'id'))
            if relation is None or relation.get('TargetMode')=='External':
                problems.append('chart workbook relationship is not embedded')
            else:
                member=posixpath.normpath(folder+'/'+relation.get('Target'))
                tabs=workbook(BytesIO(z.read(member)))
                sheet=next(iter(tabs.values()))
                from .rawxlsx import address
                if [sheet.get(address(i+2,1)) for i in range(len(categories))]!=categories:problems.append('embedded categories mismatch')
                for col,(name,values) in enumerate(series,2):
                    if sheet.get(address(1,col))!=name:problems.append('embedded series name mismatch')
                    try:
                        if [exact(sheet.get(address(i+2,col))) for i in range(len(values))]!=[exact(v) for v in values]:problems.append('embedded values mismatch')
                    except ValueError:problems.append('embedded missing/nonnumeric values')
    return problems
