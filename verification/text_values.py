"""Independent arithmetic for supported numeric calculation placeholders.

Selection and business definitions are declared inputs; arithmetic is reconstructed
from raw source recipes, not from the application's result or facts module.
"""
import datetime as dt
from fractions import Fraction
from numfmt import format_value
from . import recipes
from .kernel import round_rational, sum_rationals


def expected(book, binding):
    spec=binding['calculation'];name=spec['name'];rows=binding['grid'][1:]
    expressions=binding['recipe'][1:]
    if not rows or len(rows)!=len(expressions):raise recipes.Unsupported('calculation input coverage is incomplete')
    values=[]
    for expr in expressions:
        value=recipes.figure(recipes.evaluate(book,expr[1],decoded=True))
        if value is None:raise recipes.Unsupported('calculation input is not a supported number')
        if expr[1]['op']=='cell':value=recipes._decoded_operand(book,expr[1],value)
        values.append(value)
    count=spec['count']
    if name in ('peak','low'):
        if count<1:raise recipes.Unsupported('extreme window must be positive')
        start=max(0,len(values)-count);target=(max if name=='peak' else min)(values[start:])
        index=next(i for i in range(len(values)-1,start-1,-1) if values[i]==target)
        raw=recipes.evaluate(book,expressions[index][0],decoded=True)
        if spec.get('date_parts'):
            if spec.get('quarter'):return f'{raw.year} Q{(raw.month-1)//3+1}'
            return raw.strftime('%Y-%m-%d' if spec['date_parts']==3 else '%Y-%m')
        rendered,covered=format_value(raw,rows[index][0][2])
        if not covered:raise recipes.Unsupported('extreme date format is outside independent rendering')
        return rendered
    if name=='prev':base=values[-2] if len(values)>=2 else None
    elif name=='avg':base=sum_rationals(values[-count-1:-1])/count if count>=1 and len(values)>count else None
    elif name=='week':
        dates=[]
        for expr in expressions:
            value=recipes.evaluate(book,expr[0],decoded=True)
            date=value if isinstance(value,dt.datetime) else recipes.calendar_text(value)
            if date is None:raise recipes.Unsupported('week comparison date cannot be independently reconstructed')
            dates.append(date.date())
        wanted=dates[-1]-dt.timedelta(days=7)
        base=next((values[i] for i in range(len(values)-1,-1,-1) if dates[i]==wanted),None)
    else:raise recipes.Unsupported('calculation is outside the independent numeric placeholder grammar')
    if base is None:raise recipes.Unsupported('calculation has no comparison base')
    difference=values[-1]-base
    if spec['rate']:
        if spec['scale']==0:
            size=abs(difference)
            pattern='#,##0' if size>=1000 or difference.denominator==1 else '0.0' if size>=10 else '0.00' if size>=1 else '0.000'
            text,covered=format_value(difference,pattern)
            if not covered:raise recipes.Unsupported('amount change format is unsupported')
            text=text.rstrip('0').rstrip('.') if '.' in text else text
            return ('+' if difference>0 else '')+text
        rounded=round_rational(difference*Fraction(str(spec['scale'])),1)
        if not rounded:rounded=abs(rounded)
        return f'{rounded:+.1f}'+('포인트' if spec['lang']=='ko' else ' points')
    if not base:raise recipes.Unsupported('relative change has zero comparison base')
    value=difference/abs(base)
    zero,covered=format_value(abs(value),'0.0%')
    if not covered:raise recipes.Unsupported('relative change format is unsupported')
    return '0.0%' if zero=='0.0%' else format_value(value,'+0.0%;-0.0%;0.0%')[0]
