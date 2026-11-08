"""Typed bridge to the actual Dafny-generated code; conversions are tested TCB."""
from decimal import Decimal
from fractions import Fraction
from math import lcm
from .verified_kernel import default__ as verified


def same_ints(first,second):
    a,b=list(first),list(second)
    if any(type(x) is not int for x in a+b):raise TypeError('integer sequences required')
    return verified.Same(a,b)


def sum_decimals(first,second):
    values=[Decimal(first),Decimal(second)]
    if not all(x.is_finite() for x in values):raise ValueError('finite decimals required')
    scale=max([0]+[-x.as_tuple().exponent for x in values])
    integers=[]
    for value in values:
        sign,digits,exponent=value.as_tuple()
        coefficient=int(''.join(map(str,digits)) or '0')*(-1 if sign else 1)
        integers.append(coefficient*10**(exponent+scale))
    total=verified.Sum(integers)
    return Decimal((int(total<0),tuple(map(int,str(abs(total)))),-scale))


def equal_rationals(first,second,tolerance=Fraction(0)):
    a,b,eps=Fraction(first),Fraction(second),Fraction(tolerance)
    if eps<0:raise ValueError('negative tolerance')
    return verified.RationalWithin(a.numerator,a.denominator,b.numerator,b.denominator,eps.numerator,eps.denominator)


def sum_rationals(values):
    """Exact common-denominator bridge to the generated integer sum contract."""
    values=[Fraction(value) for value in values]
    denominator=lcm(*(value.denominator for value in values))
    return Fraction(verified.Sum([value.numerator*(denominator//value.denominator) for value in values]),denominator)


def sum_error_bounds(differences):
    """Proved triangle bound for exact rational row-decoding errors."""
    values=[Fraction(value) for value in differences]
    denominator=lcm(*(value.denominator for value in values))
    deltas=[value.numerator*(denominator//value.denominator) for value in values]
    errors=[abs(value) for value in deltas]
    if len(deltas)!=len(errors) or any(not 0<=e or not -e<=d<=e for d,e in zip(deltas,errors)):raise ArithmeticError('invalid error-bound precondition')
    total,budget=verified.SumBound(deltas,errors)
    return Fraction(total,denominator),Fraction(budget,denominator)


def round_decimal(value,places):
    if type(places) is not int or not 0<=places<=1000:raise ValueError('decimal places outside supported range')
    value=Decimal(value)
    if not value.is_finite():raise ValueError('finite decimal required')
    exact=Fraction(value);n,d=exact.numerator,exact.denominator;scale=10**places
    quotient=abs(n)*scale//d
    # Establish the generated method's precondition with integer arithmetic.
    if not quotient*d<=abs(n)*scale<(quotient+1)*d:raise ArithmeticError('invalid floor quotient')
    rounded=verified.RoundScaled(n,d,scale,quotient)
    return Decimal((int(rounded<0),tuple(map(int,str(abs(rounded)))),-places))
