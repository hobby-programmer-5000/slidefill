"""Independent schema/coverage oracle over raw XML; no application imports.

Coverage is bounded to the declared rate/organization grammars. Agreement with
this oracle is empirical evidence, not proof that either parser knows intent.
"""
import re
from .arithmetic import exact
from .rawxlsx import address, position


def clean(value):
    return ''.join(str(value or '').split())


def numeric(value):
    try:
        return exact(value)
    except ValueError:
        return None


def token(op, *refs, text=''):
    return {'op':op,'refs':list(refs),'text':text}


def expected(sheet):
    found = {}
    rows = sorted({position(k)[0] for k,v in sheet.cells.items() if v not in (None,'')})
    last_row=max(rows,default=0);last_col=max((position(k)[1] for k,v in sheet.cells.items() if v not in (None,'')),default=0)
    candidates={position(k) for k,v in sheet.cells.items() if clean(v).lower() in {'최저','최소','min','minimum'}}
    for top,left,bottom,right in sheet.merges:
        if (top,left) in candidates:candidates.update((r,c) for r in range(top,bottom+1) for c in range(left,right+1))
    rows=sorted({r for r,c in candidates}|{position(k)[0] for k,v in sheet.cells.items() if clean(v).lower() in {'현원','정원','headcount','quota'}})
    for r in rows:
        # Rate grammars are anchored by explicit min/max header cells.
        for low in sorted(c for rr,c in candidates if rr==r and r>=2):
            at = lambda rr,cc: sheet.get(sheet.anchor(rr,cc))
            if clean(at(r,low)).lower() not in {'최저','최소','min','minimum'} or clean(at(r,low+1)).lower() not in {'최고','최대','max','maximum'}:
                continue
            for col in range(low+2,last_col-1):
                headings = tuple(clean(at(r,c)).lower() for c in range(col,col+3))
                if headings not in {('식품비','운영비','계'),('food','operations','total')}:
                    continue
                items, bands = [],0
                for rr in range(r+1,last_row+1):
                    a,b=address(rr,low),address(rr,low+1)
                    costs=[address(rr,c) for c in range(col,col+3)]
                    values=[sheet.get(ref) for ref in costs]
                    lo,hi=sheet.get(a),sheet.get(b)
                    if lo in (None,'') and hi in (None,'') and all(x in (None,'') for x in values):break
                    n,m=numeric(lo),numeric(hi)
                    if n is not None:
                        if n < 0 or n.denominator != 1 or (hi not in (None,'') and (m is None or m < n or m.denominator != 1)):break
                        label=token('band',a,b);bands+=1
                    elif isinstance(lo,str) and lo.strip() and hi in (None,'') and any(numeric(x) is not None for x in values):
                        label=token('exception',a)
                    else:break
                    items.append([label]+[token('cell',ref) for ref in costs])
                if bands>=2:found[f'rates-{r}-{col}']=items
        # Organization grammars use explicit metric rows and header rectangles.
        for ref,value in sheet.cells.items():
            rr,labelcol=position(ref)
            if rr != r or r<3 or clean(value).lower() not in {'현원','정원','headcount','quota'}:continue
            groups={}
            for col in range(labelcol+1,last_col+1):
                parent=sheet.anchor(r-2,col)
                pr,pc=position(parent)
                name=sheet.get(parent)
                if pc<=labelcol or not isinstance(name,str) or not clean(name):continue
                groups.setdefault(parent,[]).append(col)
            items=[]
            for parent,columns in groups.items():
                name=clean(sheet.get(parent)).lower()
                if name in {'계','총계','합계','total','subtotal','일반직','별정직','정무직'}:continue
                slots={};seen=set();ambiguous=False
                for col in columns:
                    h=sheet.anchor(r-1,col)
                    if h in seen:continue
                    seen.add(h)
                    if h==parent and len(columns)==1:slot=0
                    else:slot={'계':0,'총계':0,'합계':0,'total':0,'subtotal':0,'일반직':1,'general':1,'별정직':2,'special':2,'정무직':3,'political':3}.get(clean(sheet.get(h)).lower())
                    if slot is not None:
                        if slot in slots:ambiguous=True
                        slots[slot]=address(r,col)
                if ambiguous or 0 not in slots or numeric(sheet.get(slots[0])) is None:continue
                items.append([token('compact',parent)]+[token('cell',slots[i]) if i in slots else token('literal',text='') for i in range(4)])
            if len(items)>=2:found[f'staff-{r}-{labelcol}']=items
    # Explicit Excel row grouping plus adjacent item/value role headers only.
    for ref,name in sheet.cells.items():
        if clean(name).lower() not in {'항목','기관','부서','item','organization','department'}:continue
        r,c=position(ref)
        if clean(sheet.get(address(r,c+1))).lower() not in {'값','금액','인원','현원','정원','value','amount','headcount'}:continue
        records=[];levels=[]
        for rr in range(r+1,last_row+1):
            label=sheet.get(address(rr,c));value=sheet.get(address(rr,c+1))
            if label in (None,'') and value in (None,''):break
            if not isinstance(label,str) or not label.strip() or value not in (None,'') and numeric(value) is None:break
            level=(sheet.row_levels or {}).get(rr,0)
            if not 0<=level<=7:break
            levels.append(level)
            records.append([token('cell',address(rr,c)),token('outline_depth',address(rr,c)),token('cell',address(rr,c+1))])
        if len(records)>=2 and any(levels):found[f'outline-{r}-{c}']=records
    return found


def expected_text(sheet, item, lang):
    op, refs = item['op'],item['refs']
    if op=='outline_depth':return str((sheet.row_levels or {}).get(position(refs[0])[0],0))
    if op=='literal':
        if item['text']!='':raise ValueError('nonempty literal in data row')
        return ''
    values=[sheet.get(ref) for ref in refs]
    if op=='band':return f"{values[0]}–{values[1] if values[1] not in (None,'') else ('상한 미기재' if lang=='ko' else 'upper unstated')}"
    if op=='compact':return clean(values[0])
    if op=='exception':return ('학교 예외: ' if lang=='ko' else 'School exception: ')+str(values[0]).strip()
    if op=='cell':return '' if values[0] is None else str(values[0])
    raise ValueError('unknown operation')


def coverage(book, certificate):
    problems=[]
    declared={}
    for page in certificate['tables']:
        lang=page['lang']
        if lang not in ('ko','en'):
            problems.append('unknown certificate language')
            continue
        headers=schema_headers(book[page['sheet']],page['view'],lang)
        if page['headers']!=headers:problems.append('header meaning/language differs from schema')
        key=f"layout:{page['view']}~{page['start']}~{page['stop']}:{lang}:{page['sheet']}!grid"
        if key!=page['key']:problems.append('certificate binding differs from declared scope')
        identity=(page['sheet'],page['view'])
        declared.setdefault(identity,[]).append(page)
    all_expected={(name,key):rows for name,sheet in book.items() for key,rows in expected(sheet).items()}
    if set(declared)!=set(all_expected):problems.append('schema coverage mismatch: missing/extra source blocks')
    for identity,pages in declared.items():
        if identity not in all_expected:continue
        wanted=all_expected[identity];actual=[];offset=0
        for page in pages:
            if page['start']!=offset:problems.append(f'{identity}: missing/reordered/duplicate page')
            if page['stop']-page['start']!=len(page['rows']):problems.append(f'{identity}: incorrect page count')
            actual.extend(page['rows']);offset=page['stop']
        if actual!=wanted:problems.append(f'{identity}: source roles/references/order/row coverage mismatch')
    return problems


def findings(book):
    """Source inconsistencies are findings even if the app preserved them exactly."""
    result=[]
    for name,sheet in book.items():
        for ident,rows in expected(sheet).items():
            if not ident.startswith('rates-'):continue
            for row in rows:
                label=row[0]
                if label['op']=='band' and sheet.get(label['refs'][1]) in (None,''):
                    result.append({'kind':'unknown_upper_bound','sheet':name,'cell':label['refs'][1]})
                refs=[item['refs'][0] for item in row[1:]]
                nums=[numeric(sheet.get(ref)) for ref in refs]
                if all(v is not None for v in nums) and nums[0]+nums[1]!=nums[2]:
                    expected_value=nums[0]+nums[1]
                    result.append({'kind':'source_total_mismatch','sheet':name,'cell':refs[2],
                                   'source':str(nums[2]),'recalculated':str(expected_value)})
    return result


def context_labels(sheet, ident, rows):
    """Source period text and rate-group identity must survive on each page."""
    kind,row,col=ident.split('-');row=int(row);col=int(col)
    labels=[]
    for ref,value in sheet.cells.items():
        rr,cc=position(ref)
        if rr<=row and isinstance(value,str) and len(value)<=120 and numeric(value) is None and re.search(r'(?:19|20)\d\d',value):
            labels.append(value)
    if kind=='rates':
        group=sheet.get(sheet.anchor(row-1,col))
        if clean(group).lower() in {'food','operations','total','식품비','운영비','계'} or not group:
            low=position(rows[0][0]['refs'][0])[1]
            group=next((sheet.get(address(back,low)) for back in range(row-1,max(0,row-4),-1)
                        if sheet.get(address(back,low)) and clean(sheet.get(address(back,low))).lower() not in {'학생수구간','min','minimum','최저'}),None)
        if group:labels.append(str(group))
    return labels


def schema_headers(sheet,ident,lang):
    if ident.startswith('outline-'):
        return ['항목 (원본)','그룹 수준 (원본)','값 (원본)'] if lang=='ko' else ['Item (source)','Grouping level (source)','Value (source)']
    if ident.startswith('rates-'):
        return ['학생 수 구간 / 학교 예외','식품비','운영비','계 (원본)'] if lang=='ko' else ['Population band / school exception','Food','Operations','Total (source)']
    _,row,col=ident.split('-');quota=clean(sheet.get(address(int(row),int(col)))).lower() in {'정원','quota'}
    return ['기관 / 원본 구역','정원 (원본)' if quota else '현원 (원본)','일반직','별정직','정무직'] if lang=='ko' else ['Organization / source block','Quota (source)' if quota else 'Headcount (source)','General','Special','Political']


def warning_visible(finding,words):
    candidates=[p for p in words if finding['cell'] in p]
    if finding['kind']=='unknown_upper_bound':
        return any('상한 미기재' in p or 'upper bound not stated' in p for p in candidates)
    from fractions import Fraction
    required={Fraction(finding['source']),Fraction(finding['recalculated'])}
    for text in candidates:
        values={Fraction(n) for n in re.findall(r'(?<![A-Za-z0-9_.])-?\d+(?:\.\d+)?(?![A-Za-z0-9_]|\.\d)',text)}
        if required<=values and any(word in text for word in ('≠','differs','불일치')):return True
    return False


def grouping_caveat_visible(words):
    return any('Grouping levels are source metadata, not verified business membership.' in p or '그룹 수준은 원본 메타데이터이며 업무상 소속 관계를 인증하지 않습니다.' in p for p in words)
