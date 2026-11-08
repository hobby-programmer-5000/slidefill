"""Independent, bounded formula-cache validation. Never imports application code.

Exact rational arithmetic covers an explicit Excel grammar. Unknown functions,
coercions, external links, circular/array formulas and resource limits leave
coverage incomplete. A supported formula with an inconsistent cache is a failure.
This checks saved results; it does not modify or recalculate the input workbook.
"""
from fractions import Fraction
import math
import re

from .rawxlsx import position, address
from .kernel import equal_rationals, round_rational
from .recipes import _decoded_operand, Unsupported as RecipeUnsupported


class Unsupported(ValueError):
    pass


class Mismatch(ValueError):
    pass


MAX_CELLS = 100_000
MAX_DEPTH = 64
MAX_BITS = 65_536
NUMBER = r'(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d{1,4})?'
REF = r"(?:(?:'(?:[^']|'')+'|[^\s!()+*/^%<>=,:;\"\[\]-]+)!)?\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}"
TOKEN = re.compile(r'\s*(?:(' + NUMBER + r')|(' + REF + r')(?![\w.])|([A-Za-z_][A-Za-z_0-9.]*)|("(?:[^"]|"")*")|(<=|>=|<>|[()+*/^%<>=,:;\-]))')
CELL = re.compile(r'(?:(.+)!)?(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)')
FUNCTIONS = {'IF','TRUE','FALSE','SUM','AVERAGE','COUNT','MIN','MAX','ABS','INT','ROUND','ROUNDUP','ROUNDDOWN'}


def bounded(value):
    if isinstance(value, Fraction) and max(value.numerator.bit_length(), value.denominator.bit_length()) > MAX_BITS:
        raise Unsupported('formula result exceeds exact-arithmetic budget')
    return value


class Parser:
    def __init__(self, text):
        text = text.strip().removeprefix('=')
        if len(text) > 8192:
            raise Unsupported('formula exceeds length budget')
        self.tokens = []
        at = 0
        while at < len(text):
            match = TOKEN.match(text, at)
            if not match:
                raise Unsupported('formula syntax outside independent grammar')
            self.tokens.append(next((kind, word) for kind, word in zip(('number','ref','name','text','op'), match.groups()) if word is not None))
            at = match.end()
        if len(self.tokens) > 2048:
            raise Unsupported('formula exceeds token budget')
        self.at = 0
        self.depth = 0

    def peek(self, word=None):
        token = self.tokens[self.at] if self.at < len(self.tokens) else ('end', '')
        return token if word is None else token[1] == word

    def take(self, word=None):
        token = self.peek()
        if token[0] == 'end' or word is not None and token[1] != word:
            raise Unsupported('incomplete formula syntax')
        self.at += 1
        return token

    def parse(self):
        node = self.expression()
        if self.at != len(self.tokens):
            raise Unsupported('formula reference operators or trailing syntax unsupported')
        return node

    def expression(self):
        self.depth += 1
        if self.depth > MAX_DEPTH:
            raise Unsupported('formula nesting exceeds budget')
        try:
            node = self.addition()
            while self.peek()[1] in ('=','<>','<','>','<=','>='):
                node = ('binary', self.take()[1], node, self.addition())
            return node
        finally:
            self.depth -= 1

    def addition(self):
        node = self.product()
        while self.peek()[1] in ('+','-'):
            node = ('binary', self.take()[1], node, self.product())
        return node

    def product(self):
        node = self.power()
        while self.peek()[1] in ('*','/'):
            node = ('binary', self.take()[1], node, self.power())
        return node

    def power(self):
        # Excel: negation, percent, then exponentiation; equal precedence is left-to-right.
        node = self.percent()
        while self.peek('^'):
            self.take('^')
            node = ('binary', '^', node, self.percent())
        return node

    def percent(self):
        node = self.unary()
        while self.peek('%'):
            self.take('%')
            node = ('binary', '/', node, ('literal', Fraction(100)))
        return node

    def unary(self):
        signs = 1
        while self.peek()[1] in ('+','-'):
            if self.take()[1] == '-': signs *= -1
        node = self.atom()
        return node if signs == 1 else ('binary', '*', ('literal', Fraction(-1)), node)

    def atom(self):
        kind, word = self.take()
        if word == '(':
            node = self.expression(); self.take(')'); return node
        if kind == 'number':
            if 'e' in word.lower() and abs(int(word.lower().split('e')[1])) > 1000:
                raise Unsupported('formula decimal exponent exceeds budget')
            return ('literal', bounded(Fraction(word)))
        if kind == 'text': return ('literal', word[1:-1].replace('""', '"'))
        if kind == 'ref':
            if self.peek(':'):
                self.take(':'); end = self.take()
                if end[0] != 'ref': raise Unsupported('unsupported formula range')
                return ('range', word, end[1])
            return ('cell', word)
        if kind == 'name':
            if word.upper() in ('TRUE','FALSE') and not self.peek('('):
                return ('literal', word.upper() == 'TRUE')
            if word.upper() not in FUNCTIONS:raise Unsupported('formula function outside independent grammar: '+word)
            self.take('('); args = []
            if not self.peek(')'):
                args.append(self.expression())
                while self.peek()[1] in (',',';'):
                    self.take(); args.append(self.expression())
            self.take(')')
            return ('call', word.upper(), args)
        raise Unsupported('unsupported formula expression')


def references(recipe, sheet=None, book=None):
    """Collect cells including filter domains and structured-source roles."""
    found = set();seen=set()
    def visit(node, current):
        if isinstance(node,(dict,list,tuple)):
            identity=(id(node),current)
            if identity in seen:return
            seen.add(identity)
        if isinstance(node, dict):
            current = node.get('sheet', current)
            if node.get('op')=='selection_domain' and book is not None:
                source=book[current];lines=set(node['lines']);by_row=node['by_row']
                for ref in source.formulas or {}:
                    r,c=position(ref)
                    if (r if by_row else c) in lines and node['start']<=(c if by_row else r)<=node['stop']:
                        found.add((current,ref))
                for span in source.formula_regions or ():
                    parts=span.replace('$','').split(':');top,left=position(parts[0]);bottom,right=position(parts[-1])
                    low,high=(top,bottom) if by_row else (left,right)
                    first,last=(left,right) if by_row else (top,bottom)
                    for line in lines:
                        pos=max(first,node['start'])
                        if low<=line<=high and pos<=min(last,node['stop']):
                            found.add((current,address(line,pos) if by_row else address(pos,line)))
            if current and isinstance(node.get('cell'), str): found.add((current, node['cell']))
            if current:
                for key in ('refs', 'domain', 'date_domain'):
                    for ref in node.get(key, ()):
                        if isinstance(ref, str): found.add((current, ref))
            for value in node.values(): visit(value, current)
        elif isinstance(node, (list, tuple)):
            for value in node: visit(value, current)
    visit(recipe, sheet)
    return found


class Audit:
    def __init__(self, book):
        self.book = book
        self.names = {name.casefold(): name for name in book}
        self.active = set()
        self.values = {}
        self.checked = set()
        self.visited = 0

    def cached(self, name, sheet, ref):
        value = sheet.get(ref)
        kind = (sheet.types or {}).get(ref)
        if kind=='d':raise Unsupported('ISO date formula operands are outside the independent grammar')
        if kind not in (None,'n','s','inlineStr','str','b','e'):
            raise Unsupported('formula dependency has an unsupported source cell type')
        if kind == 'e': raise Unsupported('source formula dependency is an Excel error')
        if kind == 'n' and value is not None:
            if len(str(value))>8192 or not re.fullmatch(r'[+-]?'+NUMBER,str(value)):
                raise Unsupported('numeric formula cache/operand exceeds lexical budget')
            try: return bounded(_decoded_operand(self.book,{'sheet':name,'cell':ref},Fraction(value)))
            except RecipeUnsupported as exc:raise Unsupported(str(exc)) from exc
            except (ValueError, ZeroDivisionError): raise Unsupported('invalid numeric formula dependency')
        return value

    def cell(self, name, ref):
        name = self.names.get(name.casefold())
        if name is None: raise Unsupported('formula references an unavailable sheet')
        key = (name, ref)
        if key in self.values: return self.values[key]
        self.visited += 1
        if self.visited > MAX_CELLS or len(self.active) >= MAX_DEPTH:
            raise Unsupported('formula dependency budget exceeded')
        if key in self.active: raise Unsupported('circular formula dependency is not independently supported')
        sheet = self.book[name]; cached = self.cached(name,sheet,ref)
        formula = (sheet.formulas or {}).get(ref)
        for span in sheet.formula_regions or ():
            if inside(ref, span): raise Unsupported('array/data-table formula cache is not independently supported')
        if formula is None:
            self.values[key] = cached
            return cached
        if formula.get('kind') not in ('normal','shared','ods'):
            raise Unsupported('binary/array formula cache is not independently supported')
        text = formula.get('text', '')
        if not text: raise Unsupported('formula expression is missing')
        origin = formula.get('origin', ref)
        if formula['kind'] == 'shared' and (not formula.get('range') or not inside(ref, formula['range'])):
            raise Unsupported('shared formula range is unavailable or inconsistent')
        if formula['kind'] == 'ods':
            text = text.removeprefix('of:=').removeprefix('=')
            text = re.sub(r'\[\.([A-Za-z$0-9:]+)\]', r'\1', text)
        self.active.add(key)
        try:
            expected = self.evaluate(Parser(text).parse(), name, ref, origin)
            if isinstance(expected, list): raise Unsupported('array-valued formula outside supported grammar')
            if expected is None: expected = Fraction(0)
            if expected == '' and cached is None and (sheet.types or {}).get(ref) == 'str': cached = ''
            if not same(expected, cached): raise Mismatch(f'stale or inconsistent formula result at {name}!{ref}')
            self.checked.add(key); self.values[key] = expected
            return expected
        finally:
            self.active.remove(key)

    def locate(self, token, name, ref, origin):
        match = CELL.fullmatch(token)
        if not match: raise Unsupported('unsupported formula reference')
        title, abs_col, col, abs_row, row = match.groups()
        if title:
            name = title[1:-1].replace("''", "'") if title.startswith("'") else title
        r, c = position(col.upper() + row)
        here_r, here_c = position(ref); base_r, base_c = position(origin)
        r += 0 if abs_row else here_r - base_r
        c += 0 if abs_col else here_c - base_c
        if not 1 <= r <= 1048576 or not 1 <= c <= 16384: raise Unsupported('formula reference outside worksheet grid')
        return name, r, c

    def evaluate(self, node, name, ref, origin, depth=0):
        if depth > MAX_DEPTH: raise Unsupported('formula expression depth exceeds budget')
        def run(child): return self.evaluate(child, name, ref, origin, depth+1)
        if node[0] == 'literal': return node[1]
        if node[0] == 'cell':
            title, r, c = self.locate(node[1], name, ref, origin)
            return self.cell(title, address(r,c))
        if node[0] == 'range':
            title, top, left = self.locate(node[1], name, ref, origin)
            other, bottom, right = self.locate(node[2], title, ref, origin)
            if title.casefold() != other.casefold() or top > bottom or left > right:
                raise Unsupported('formula range order/sheet is unsupported')
            if (bottom-top+1)*(right-left+1) > MAX_CELLS: raise Unsupported('formula range exceeds budget')
            return [self.cell(title,address(r,c)) for r in range(top,bottom+1) for c in range(left,right+1)]
        if node[0] == 'binary':
            op = node[1]; first = run(node[2]); second = run(node[3])
            if op in ('=','<>','<','>','<=','>='):
                if not isinstance(first,Fraction) or not isinstance(second,Fraction):
                    raise Unsupported('non-numeric formula comparison is outside the independent grammar')
                a,b=first,second
                return {'=':a==b,'<>':a!=b,'<':a<b,'>':a>b,'<=':a<=b,'>=':a>=b}[op]
            a,b=number(first),number(second)
            if op == '+': value = a+b
            elif op == '-': value = a-b
            elif op == '*': value = a*b
            elif op == '/':
                if b == 0: raise Mismatch('source formula divides by zero')
                value = a/b
            elif op == '^':
                if b.denominator != 1 or abs(b)>1000: raise Unsupported('noninteger/large formula exponent unsupported')
                if not a and b<0: raise Mismatch('source formula divides by zero')
                if max(a.numerator.bit_length(),a.denominator.bit_length())*abs(b)>MAX_BITS:
                    raise Unsupported('formula power exceeds exact-arithmetic budget')
                value = a**int(b)
            else: raise Unsupported('unsupported formula operator')
            return bounded(value)
        function, args = node[1:]
        if function == 'IF' and len(args) in (2,3):
            condition=run(args[0])
            if not isinstance(condition,(Fraction,bool)) and condition is not None:
                raise Unsupported('text/array formula condition is outside the independent grammar')
            return run(args[1]) if number(condition) else run(args[2]) if len(args)==3 else False
        if function in ('TRUE','FALSE') and not args: return function=='TRUE'
        if function in ('SUM','AVERAGE','COUNT','MIN','MAX'):
            values=[]
            for arg in args:
                value=run(arg); items=value if isinstance(value,list) else [value]
                referenced=arg[0] in ('cell','range')
                for item in items:
                    if item is None or referenced and not isinstance(item,Fraction): continue
                    values.append(number(item))
            if function=='COUNT': return Fraction(len(values))
            if function=='SUM': return bounded(sum(values,Fraction()))
            if function in ('MIN','MAX'): return (min if function=='MIN' else max)(values) if values else Fraction()
            if not values: raise Mismatch('source formula average has no numeric operands')
            return bounded(sum(values,Fraction())/len(values))
        values=[number(run(arg)) for arg in args]
        if function=='ABS' and len(values)==1: return abs(values[0])
        if function=='INT' and len(values)==1: return Fraction(values[0].numerator//values[0].denominator)
        if function in ('ROUND','ROUNDUP','ROUNDDOWN') and len(values)==2:
            value, places=values
            if places.denominator!=1 or abs(places)>1000: raise Unsupported('formula rounding precision unsupported')
            scale=Fraction(10)**int(places); scaled=abs(value)*scale
            if function=='ROUND':
                result=Fraction(round_rational(value,int(places))) if places>=0 else Fraction(round_rational(value*scale,0))/scale
                return bounded(result)
            whole=scaled.numerator//scaled.denominator
            if function=='ROUNDUP': whole += int(scaled!=whole)
            return bounded(Fraction((-1 if value<0 else 1)*whole)/scale)
        raise Unsupported('formula function outside independent grammar: '+function)


def number(value):
    if isinstance(value,list): raise Unsupported('range cannot be used as a scalar formula operand')
    if value is None: return Fraction()
    if isinstance(value,bool): return Fraction(int(value))
    if isinstance(value,Fraction): return value
    if isinstance(value,str) and re.fullmatch(r'[+-]?'+NUMBER,value.strip()): return bounded(Fraction(value.strip()))
    raise Unsupported('formula text/locale coercion outside independent grammar')


def inside(ref, span):
    parts=span.replace('$','').split(':')
    top,left=position(parts[0]); bottom,right=position(parts[-1]); r,c=position(ref)
    return top<=r<=bottom and left<=c<=right


def same(a,b):
    if isinstance(a,Fraction) and isinstance(b,Fraction):
        if a.denominator==b.denominator==1: return a==b
        try: tolerance=Fraction.from_float(math.ulp(float(max(abs(a),abs(b)))))*2
        except (OverflowError,ValueError): raise Unsupported('formula result outside finite cache precision')
        return equal_rationals(a,b,tolerance)
    return type(a) is type(b) and a==b


def inspect(book, refs):
    audit=Audit(book); failures=[]; partial=[]
    for name,ref in sorted(set(refs)):
        try:
            sheet=book[name]
            if ref not in (sheet.formulas or {}) and not any(inside(ref,span) for span in sheet.formula_regions or ()):
                continue
            audit.cell(name,ref)
        except Mismatch as exc: failures.append(str(exc))
        except (Unsupported,ValueError,TypeError,KeyError,IndexError,RecursionError) as exc:
            partial.append(f'{name}!{ref}: {exc}')
    return {'checked':len(audit.checked),'failures':list(dict.fromkeys(failures)),
            'partial':list(dict.fromkeys(partial))}
