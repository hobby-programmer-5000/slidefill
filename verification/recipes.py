"""Evaluate explicit source-cell recipes independently of application arithmetic."""
import datetime as dt
from fractions import Fraction
import math
import re
from .kernel import equal_rationals,sum_error_bounds
from numfmt import format_value
from .rawxlsx import address, position


class Unsupported(ValueError):
    pass


def figure(value):
    if value is None or isinstance(value,bool):return None
    if isinstance(value,(int,float,Fraction)):
        if isinstance(value,float) and not math.isfinite(value):return None
        return Fraction(str(value))
    text=str(value).strip().replace('\u2212','-').replace('\u00a0','').replace('\u202f','')
    if re.fullmatch(r'[+-]?\d{1,3}(?: +\d{3})+(?:\.\d+)?',text):text=text.replace(' ','')
    if text in ('','-','—','–'):return None
    sign=1
    if text.startswith(('△','▲')):sign=-1;text=text[1:]
    if text.startswith('(') and text.endswith(')'):
        if len(re.sub(r'\D','',text))<3:return None
        sign=-1;text=text[1:-1]
    pattern=r'([+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?)(?:\s*(?:\([^)]*%\)|건|명|개|회|원|톤|kg))?'
    match=re.fullmatch(pattern,text,re.I)
    if match:return sign*Fraction(match[1].replace(',',''))
    if any(ch.isdigit() for ch in text):raise Unsupported('numeric text outside independent parser contract')
    return None


def near(a,b):
    if a is None or b is None:return a is b
    a,b=Fraction(a),Fraction(b)
    # Integers are exact. Decimal/float bridges get a bounded number of binary
    # ULPs, never a blanket relative budget that could hide whole currency units.
    if a.denominator==b.denominator==1:tolerance=Fraction(0)
    else:
        try:
            ulp=math.ulp(float(max(abs(a),abs(b))))
            if not math.isfinite(ulp):raise OverflowError
        except OverflowError:raise Unsupported('number outside finite binary chart precision')
        tolerance=Fraction.from_float(ulp)*2
    return equal_rationals(a,b,tolerance)


def _raw(book,node):
    sheet=book[node['sheet']];value=sheet.get(node['cell'])
    if value is not None and (sheet.types or {}).get(node['cell'])=='d':
        try:return dt.datetime.fromisoformat(value)
        except ValueError:raise Unsupported('invalid ISO source date')
    if value is not None and (sheet.types or {}).get(node['cell'])=='n':
        try:return Fraction(value)
        except (ValueError,ZeroDivisionError):pass
    return value


def _matches(book,node,wanted):
    sheet=book[node['sheet']]
    value=evaluate(book,node) if node['op']=='date' else (sheet.condition_values or {}).get(node['cell'],sheet.get(node['cell']))
    if value is None:return False
    text=str(value).strip().casefold();word=str(wanted).strip().casefold()
    year=lambda x:re.fullmatch(r'\[?((?:19|20)\d\d)\]?년?',x)
    if year(word) and year(text):return year(word)[1]==year(text)[1]
    domain=node['domain'];dates=node.get('date_domain',())
    if not hasattr(sheet,'_condition_domains'):sheet._condition_domains={}
    key=(id(domain),id(dates));cached=sheet._condition_domains.get(key)
    if cached is None:
        choices=set();date_domain=set(dates)
        for ref in domain:
            raw=evaluate(book,{'op':'date','sheet':node['sheet'],'cell':ref}) if ref in date_domain else (sheet.condition_values or {}).get(ref,sheet.get(ref))
            choices.add(str(raw if raw is not None else '').strip().casefold())
        # Retain the domain objects so their identities cannot be reused in this session.
        cached=(domain,dates,choices);sheet._condition_domains[key]=cached
    choices=cached[2]
    return text==word if word in choices else word in text


def calendar_text(value):
    text=str(value).strip().strip('[]')
    text=re.sub(r'(?<=[년월])\s+','',text)
    if re.fullmatch(r'(?:19|20)\d\d년(?:도)?',text):return dt.datetime(int(text[:4]),1,1)
    match=re.fullmatch(r'(\d{4})(?:[-./년 ](\d{1,2})(?:[-./월 ](\d{1,2}))?[일월]?)?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?',text)
    if match:
        year,month,day,h,m,s=match.groups()
        return dt.datetime(int(year),int(month or 1),int(day or 1),int(h or 0),int(m or 0),int(s or 0))
    if re.fullmatch(r'\d{8}',text):return dt.datetime.strptime(text,'%Y%m%d')
    months={name:i for i,name in enumerate(('jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'),1)}
    names='January February March April May June July August September October November December'.split()
    month_words='|'.join(names+[name[:3] for name in names])
    for pattern,order in [(rf'(\d{{4}})[ -]({month_words})\.?',('y','m')), (rf'({month_words})\.?[ -](\d{{4}})',('m','y'))]:
        found=re.fullmatch(pattern,text,re.I)
        if found:
            parts=dict(zip(order,found.groups()));return dt.datetime(int(parts['y']),months[parts['m'][:3].lower()],1)
    match=re.fullmatch(r'(?:(\d{4})[ -]*Q([1-4])|Q([1-4])[ -]*(\d{4})|(\d{4})년[ ]*([1-4])분기)',text,re.I)
    if match:
        y,q,q2,y2,y3,q3=match.groups();return dt.datetime(int(y or y2 or y3),(int(q or q2 or q3)-1)*3+1,1)
    return None


def _decoded_operand(book,cell,value):
    if value is None:return None
    sheet=book[cell['sheet']];literal=sheet.get(cell['cell'])
    if not isinstance(literal,str) or ((sheet.types or {}).get(cell['cell'])=='n' and not any(c in literal for c in '.eE')):return value
    try:decoded=float(value)
    except OverflowError:raise Unsupported('aggregate operand outside finite decoded precision')
    if not math.isfinite(decoded):raise Unsupported('aggregate operand outside finite decoded precision')
    normalized=Fraction(str(decoded))
    if normalized==value:return value
    # Preserve exact integers; a short-decimal spelling may differ from the
    # stored decimal by <=1 ULP after the decimal/binary/short-decimal bridge.
    # More than 17 significant source decimal digits remain explicitly partial.
    denominator=value.denominator;twos=fives=0
    while denominator%2==0:denominator//=2;twos+=1
    while denominator%5==0:denominator//=5;fives+=1
    places=max(twos,fives)
    coefficient=abs(value.numerator)*2**(places-twos)*5**(places-fives)
    digits=len(str(coefficient).rstrip('0'))
    if value.denominator==1 or denominator!=1 or digits>17 or not equal_rationals(value,normalized,Fraction.from_float(math.ulp(decoded))):
        raise Unsupported('aggregate operand loses source decimal precision')
    return normalized


def evaluate(book,node,decoded=False):
    op=node['op']
    if op=='cell':return _raw(book,node)
    if op=='literal':return node['value']
    if op=='label':
        for ref in reversed(node['cells']):
            value=_raw(book,ref)
            if value in (None,'') or not isinstance(value,str):continue
            try:Fraction(value.replace(',',''));continue
            except ValueError:pass
            parts=value.split()
            return ''.join(parts) if parts and all(len(p)==1 for p in parts) else ' '.join(parts)
        raise Unsupported('source label could not be reconstructed')
    if op=='sum_terms':
        values=[figure(evaluate(book,term,decoded)) for term in node['terms']]
        return sum((v for v in values if v is not None),Fraction()) if any(v is not None for v in values) else None
    if op=='ratio':
        numerator=figure(evaluate(book,node['numerator'],decoded));denominator=figure(evaluate(book,node['denominator'],decoded))
        return None if numerator is None or not denominator else numerator/denominator
    if op=='date_parts':
        parts=[]
        for ref in node['cells']:
            raw=_raw(book,ref)
            if isinstance(raw,str):
                match=re.fullmatch(r'\s*\[?(\d{1,4})\]?(?:년|월|일)?\s*',raw)
                if not match:raise Unsupported('split date notation is unsupported')
                parts.append(int(match[1]))
            else:
                value=Fraction(raw)
                if value.denominator!=1:raise Unsupported('noninteger split date')
                parts.append(int(value))
        if node['quarter']:parts[1]=(parts[1]-1)*3+1
        return dt.datetime(parts[0],parts[1],parts[2] if len(parts)>2 else 1)
    if op=='aggregate':
        values=[];originals=[]
        for row in node['rows']:
            if not all(_matches(book,condition,condition['wanted']) for condition in row['filters']):continue
            raw=[_raw(book,cell) for cell in row['cells']]
            kind=node['kind']
            if kind=='count':value=Fraction(1) if any(v not in (None,'') for v in raw) else None
            elif kind=='rate':
                numerator=row['numerator'];raw_n=_raw(book,numerator)
                value=None if raw_n in (None,'') else Fraction(int(_matches(book,numerator,numerator['wanted'])))
            else:
                nums=[figure(v) for v in raw]
                original=None if node['product'] and any(v is None for v in nums) else math.prod(nums) if node['product'] else sum((v for v in nums if v is not None),Fraction()) if any(v is not None for v in nums) else None
                if decoded:nums=[_decoded_operand(book,cell,value) for cell,value in zip(row['cells'],nums)]
                if node['product']:value=None if any(v is None for v in nums) else math.prod(nums)
                else:value=sum((v for v in nums if v is not None),Fraction()) if any(v is not None for v in nums) else None
            if value is not None:values.append(value);originals.append(value if kind in ('count','rate') else original)
        if not values:return None
        total=sum(values,Fraction())
        if decoded:
            differences=[a-b for a,b in zip(values,originals)]
            delta,bound=sum_error_bounds(differences)
            if delta!=total-sum(originals,Fraction()) or not equal_rationals(delta,0,bound):raise ValueError('aggregate source-decoding error contract failed')
        return total/len(values) if node['kind'] in ('mean','rate') else total
    if op=='date':
        value=_raw(book,node)
        if value is None:return None
        if isinstance(value,dt.datetime):return value
        if isinstance(value,dt.date):return dt.datetime.combine(value,dt.time())
        text=str(value).strip()
        fmt=book[node['sheet']].formats.get(node['cell'],'General')
        if isinstance(value,str) or fmt=='General':
            try:
                parsed=calendar_text(text)
                if parsed is not None:return parsed
                return dt.datetime.fromisoformat(text)
            except ValueError:pass
        try:
            serial=float(Fraction(text));epoch=getattr(book[node['sheet']],'epoch1904',False)
            base=dt.datetime(1904,1,1) if epoch else dt.datetime(1899,12,30)
            if not epoch and 0<serial<60:serial+=1
            return base+dt.timedelta(days=serial)
        except (ValueError,OverflowError):raise Unsupported('date representation outside independent parser contract')
    raise Unsupported('unavailable independent source recipe')


def _cell_matches(book,expr,actual):
    wanted=evaluate(book,expr,decoded=True);raw=actual[0]
    if expr['op']=='cell' and isinstance(wanted,Fraction) and isinstance(raw,(int,float)) and not isinstance(raw,bool):
        wanted=_decoded_operand(book,expr,wanted)
    if isinstance(raw,bool) or isinstance(wanted,bool):return type(raw) is type(wanted) and raw==wanted
    if expr['op'] in ('date','date_parts'):
        if isinstance(raw,dt.date) and not isinstance(raw,dt.datetime):raw=dt.datetime.combine(raw,dt.time())
        return wanted==raw
    if expr['op']=='label':
        if wanted==raw:return True
        for context in expr.get('context',[]):
            value=_raw(book,context)
            if not isinstance(value,str) or not value.strip():continue
            parts=value.split();word=''.join(parts) if all(len(p)==1 for p in parts) else ' '.join(parts)
            composed=word+' '+wanted if context['before'] else wanted+' ('+word+')'
            if raw==composed:return True
        return False
    if expr['op']=='cell' and not isinstance(raw,(int,float)):
        if raw is None and wanted is None:return True
        if isinstance(raw,str) and isinstance(wanted,str):return raw==wanted
        if raw==wanted:return True
        a,b=figure(wanted),figure(raw)
        return a is not None and b is not None and near(a,b)
    if expr['op']=='aggregate':
        if wanted is None:return raw is None
        if not isinstance(raw,(int,float,Fraction)) or isinstance(raw,bool):return False
        wanted=Fraction(wanted)
        try:rounded=wanted.numerator if wanted.denominator==1 else float(wanted)
        except OverflowError:raise Unsupported('aggregate result outside finite decoded precision')
        if isinstance(rounded,float) and not math.isfinite(rounded):raise Unsupported('aggregate result outside finite decoded precision')
        # Exact integer result or a single correctly rounded float conversion.
        # No operation-count multiplier and no tolerance at the result scale.
        if raw!=(wanted if isinstance(raw,Fraction) else rounded):return False
        if actual[2] not in ('General','@'):
            expected_text,exact=format_value(wanted,actual[2])
            if not exact:raise Unsupported('aggregate output format cannot be independently rendered')
            actual_text=actual[1]
            if actual[2]=='0':expected_text=expected_text.replace(',','');actual_text=actual_text.replace(',','')
            if expected_text!=actual_text:return False
        return True
    return near(figure(wanted),figure(raw))


def audit(book,recipe,grid):
    if recipe is None:raise Unsupported('source recipe is unavailable for this reference')
    if len(recipe)!=len(grid):return ['source recipe row count differs']
    errors=[];partial=[]
    for i,(exprs,line) in enumerate(zip(recipe,grid)):
        if len(exprs)!=len(line):errors.append(f'source recipe column count at row{i}');continue
        for j,(expr,actual) in enumerate(zip(exprs,line)):
            if expr is None:continue  # interpreted header/label selection remains a trusted boundary
            try:
                if not _cell_matches(book,expr,actual):errors.append(f'source value/text/date/calculation differs at row{i}, column{j}')
            except Unsupported as exc:partial.append(str(exc))
    if not errors and partial:raise Unsupported('; '.join(dict.fromkeys(partial)))
    return errors
