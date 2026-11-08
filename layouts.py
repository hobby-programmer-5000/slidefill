"""Bounded, source-linked views of rate schedules and repeated organization blocks.

No totals are inferred from matching sums. Original cells stay untouched. These
views are descriptive tables, not spending/ranking/time-series calculations.
"""
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, localcontext
import re
from verification.kernel import sum_decimals

PAGE_ROWS = 8


def compact(value):
    return re.sub(r"\s+", "", str(value or ""))


def number(value):
    if isinstance(value, bool) or value is None:
        return None
    try:
        text = str(value).strip()
        if not re.fullmatch(r"-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?", text):
            return None
        result = Decimal(text.replace(",", ""))
        return result if result.is_finite() else None
    except InvalidOperation:
        return None


def origin(ws, r, c):
    for area in ws.merged_cells.ranges:
        if area.min_row <= r <= area.max_row and area.min_col <= c <= area.max_col:
            return ws.cell(area.min_row, area.min_col)
    return ws.cell(r, c)


def _cost_sum(first, second):
    return sum_decimals(number(first), number(second))


@dataclass(frozen=True)
class Token:
    op: str
    refs: tuple = ()
    text: str = ""

    def certificate(self):
        return {"op": self.op, "refs": list(self.refs), "text": self.text}


def literal(text):
    return Token("literal", text=text)


def cell(ws, r, c, op="cell"):
    return Token(op, (ws.cell(r, c).coordinate,))


@dataclass
class View:
    id: str
    kind: str
    title: str
    headers: tuple
    rows: list
    warnings: list = field(default_factory=list)  # bilingual pairs
    scope: str = ""

    def note(self, lang):
        index = 0 if lang == "ko" else 1
        if self.kind == "rates":
            base = ("학생 수 구간별 단가와 학교별 예외입니다. 단가를 더해 지출액으로 만들지 않았으며 빈칸은 0으로 바꾸지 않았습니다."
                    if lang == "ko" else "Rates are shown by population band, with school exceptions marked. They are not added into expenditure. Blank cells are not zero.")
        elif self.kind == "outline":
            base = ("그룹 수준은 원본 메타데이터이며 업무상 소속 관계를 인증하지 않습니다. 원본 행을 유지하고 합계와 하위 항목을 더하지 않았습니다." if lang=="ko" else
                    "Grouping levels are source metadata, not verified business membership. Source rows are retained; totals and their children are not added together.")
        else:
            base = ("기관·직종별 원본 값입니다. 빈칸은 0으로 바꾸지 않았습니다." if lang == "ko" else
                    "Source values by organization and job type. Blank cells are not zero.")
        return " ".join([base] + [pair[index] for pair in self.warnings])


def _bounds(ws):
    cells=[(r,c) for (r,c),cell in ws._cells.items() if cell.value not in (None, "")]
    return max((r for r,c in cells),default=0),max((c for r,c in cells),default=0)


def _anchors(ws,words):
    anchors=set()
    for (r,c),cell in list(ws._cells.items()):
        if compact(cell.value).lower() in words:
            anchors.add((r,c))
            for area in ws.merged_cells.ranges:
                if (area.min_row,area.min_col)==(r,c):
                    anchors.update((rr,cc) for rr in range(r,area.max_row+1) for cc in range(c,area.max_col+1))
    return sorted(anchors)


def _context(ws, row):
    entries=sorted((r,c,cell.value) for (r,c),cell in ws._cells.items() if r<=row and cell.value not in (None,""))
    first=entries[0][0] if entries else None
    words=[" ".join(v.split()) for r,c,v in entries if isinstance(v,str) and len(v)<=120 and (re.search(r"(?:19|20)\d\d",v) or r==first)]
    return " · ".join(dict.fromkeys(words))


def _rates(ws):
    views = []
    last_row,last_col=_bounds(ws)
    for r,low in _anchors(ws,{"최저","최소","min","minimum"}):
        if r<2:continue
        if compact(origin(ws, r, low).value).lower() not in ("최저", "최소", "min", "minimum"):
            continue
        if compact(origin(ws, r, low + 1).value).lower() not in ("최고", "최대", "max", "maximum"):
            continue
        c = low + 2
        while c + 2 <= last_col:
            heads = [compact(origin(ws, r, x).value) for x in (c, c + 1, c + 2)]
            if heads != ["식품비", "운영비", "계"] and [h.lower() for h in heads] != ["food", "operations", "total"]:
                c += 1
                continue
            group = origin(ws, r - 1, c)
            title = " ".join(str(group.value or "").split())
            if group.row == r or compact(title) in heads:
                title = ""
            if not title:
                title = next((" ".join(str(ws.cell(back, low).value).split()) for back in range(r - 1, max(0, r - 4), -1)
                              if isinstance(ws.cell(back, low).value, str) and compact(ws.cell(back, low).value) not in ("학생수구간", "최저")), ws.title)
            rows, warnings, previous, band_count = [], [], None, 0
            for rr in range(r + 1, last_row + 1):
                lo, hi = ws.cell(rr, low).value, ws.cell(rr, low + 1).value
                amounts = [ws.cell(rr, x).value for x in (c, c + 1, c + 2)]
                if lo in (None, "") and hi in (None, "") and all(v in (None, "") for v in amounts):
                    break
                nlo, nhi = number(lo), number(hi)
                if nlo is not None:
                    if nlo < 0 or nlo != int(nlo) or (hi not in (None, "") and (nhi is None or nhi < nlo or nhi != int(nhi))):
                        break
                    band_count += 1
                    label = Token("band", (ws.cell(rr, low).coordinate, ws.cell(rr, low + 1).coordinate))
                    if hi in (None, ""):
                        warnings.append((f"{ws.cell(rr, low + 1).coordinate}: 상한 미기재이며 무한대로 추정하지 않았습니다.",
                                         f"{ws.cell(rr, low + 1).coordinate}: upper bound not stated; not assumed infinite."))
                    if previous is not None and nlo <= previous:
                        warnings.append(("인원 구간이 겹칩니다. 조회 조건을 확인하세요.", "Population bands overlap; check lookup conditions."))
                    previous = nhi
                elif isinstance(lo, str) and lo.strip() and hi in (None, "") and any(number(v) is not None for v in amounts):
                    label = Token("exception", (ws.cell(rr, low).coordinate,))
                else:
                    break
                if not all(number(v) is not None for v in amounts):
                    warnings.append((f"{rr}행: 비용 값이 비어 있거나 수치가 아닙니다. 원본 표기를 유지했습니다.",
                                     f"Row {rr}: a cost is missing or nonnumeric; original notation retained."))
                elif _cost_sum(amounts[0], amounts[1]) != number(amounts[2]):
                    refs = [ws.cell(rr, x).coordinate for x in (c, c + 1, c + 2)]
                    expected = _cost_sum(amounts[0], amounts[1])
                    warnings.append((f"{refs[2]} 원본 합계 {amounts[2]} ≠ {refs[0]}+{refs[1]} {expected}. 원본 값을 유지했습니다.",
                                     f"{refs[2]} source total {amounts[2]} differs from {refs[0]}+{refs[1]} {expected}. Source value retained."))
                rows.append([label] + [cell(ws, rr, x) for x in (c, c + 1, c + 2)])
            if band_count >= 2:
                views.append(View(f"rates-{r}-{c}", "rates", title,
                                  (("학생 수 구간 / 학교 예외", "Population band / school exception"), ("식품비", "Food"),
                                   ("운영비", "Operations"), ("계 (원본)", "Total (source)")), rows, warnings, _context(ws, r)))
            c += 3
    return views

def _organizations(ws):
    views = []
    total_words = {"계", "총계", "합계", "total", "subtotal"}
    types = {"일반직": 2, "별정직": 3, "정무직": 4, "general": 2, "special": 3, "political": 4}
    last_row,last_col=_bounds(ws)
    for r,labelcol in _anchors(ws,{"현원","정원","headcount","quota"}):
        if r<3:continue
        metric = compact(ws.cell(r, labelcol).value).lower()
        if metric not in ("현원", "정원", "headcount", "quota"):
            continue
        candidates, used = [], set()
        for c in range(labelcol + 1, last_col + 1):
            top = origin(ws, r - 2, c)
            if not isinstance(top.value, str) or not compact(top.value) or top.column <= labelcol:
                continue
            name = compact(top.value)
            if name.lower() in total_words or name in ("일반직", "별정직", "정무직"):
                continue
            identity = top.coordinate
            if identity in used:
                continue
            used.add(identity)
            columns = [x for x in range(c, last_col + 1) if origin(ws, r - 2, x).coordinate == identity]
            values = [None] * 4
            # Vertically merged organization name with exactly one count.
            if len(columns) == 1 and top.row + 1 <= r - 1 and origin(ws, r - 1, c).coordinate == identity:
                values[0] = cell(ws, r, c)
            else:
                header_origins = set()
                for x in columns:
                    heading_cell = origin(ws, r - 1, x)
                    if heading_cell.coordinate in header_origins:
                        continue
                    header_origins.add(heading_cell.coordinate)
                    heading = compact(heading_cell.value).lower()
                    index = 1 if heading in total_words else types.get(heading)
                    if index is not None:
                        if values[index - 1] is not None:
                            values = [None] * 4  # ambiguous duplicate role
                            break
                        values[index - 1] = cell(ws, r, x)
            if values[0] is None or number(ws[values[0].refs[0]].value) is None:
                continue
            candidates.append([Token("compact", (identity,))] + [v or literal("") for v in values])
        if len(candidates) >= 2:
            headers = (("기관 / 원본 구역", "Organization / source block"),
                       ("정원 (원본)" if metric in ("정원", "quota") else "현원 (원본)", "Quota (source)" if metric in ("정원", "quota") else "Headcount (source)"),
                       ("일반직", "General"), ("별정직", "Special"), ("정무직", "Political"))
            views.append(View(f"staff-{r}-{labelcol}", "staff", _context(ws, r) or ws.title, headers, candidates,
                              [("합계와 하위 부서가 함께 있을 수 있습니다. 각 행은 원본 값이며 행끼리 더하지 않았습니다.",
                                "Totals and their departments may coexist. Each row is a source value; rows were not added together.")], _context(ws, r)))
    return views

def _outlines(ws):
    views=[];last_row,last_col=_bounds(ws)
    for r,c in _anchors(ws,{'항목','기관','부서','item','organization','department'}):
        if c>=last_col or compact(ws.cell(r,c+1).value).lower() not in {'값','금액','인원','현원','정원','value','amount','headcount'}:continue
        rows=[];has_level=False
        for rr in range(r+1,last_row+1):
            label=ws.cell(rr,c).value;value=ws.cell(rr,c+1).value
            if label in (None,'') and value in (None,''):break
            if not isinstance(label,str) or not label.strip() or value not in (None,'') and number(value) is None:break
            level=ws.row_dimensions[rr].outlineLevel
            if not 0<=level<=7:break
            has_level|=bool(level)
            rows.append([cell(ws,rr,c),Token('outline_depth',(ws.cell(rr,c).coordinate,)),cell(ws,rr,c+1)])
        if len(rows)>=2 and has_level:
            views.append(View(f'outline-{r}-{c}','outline',ws.title,
                (("항목 (원본)","Item (source)"),("그룹 수준 (원본)","Grouping level (source)"),("값 (원본)","Value (source)")),rows,[],_context(ws,r)))
    return views


def discover(ws):
    """Only explicit bounded schemas; no file names, source IDs or magic cell addresses."""
    return _rates(ws) + _organizations(ws) + _outlines(ws)


def page_key(view, start, stop, lang, sheet, part="grid"):
    return f"layout:{view.id}~{start}~{stop}:{lang}:{sheet}!{part}"


def resolve(source, key):
    match = re.fullmatch(r"layout:([a-z]+-\d+-\d+)~(\d+)~(\d+):(ko|en):(.+)!(grid|note)", key, re.S)
    if not match:
        raise ValueError("invalid layout reference")
    ident, first, last, lang, sheet, part = match.groups()
    ws = source._sheet(sheet)
    cache = getattr(source, "_layout_views", {})
    if ws.title not in cache:
        cache[ws.title] = discover(ws)
        source._layout_views = cache
    view = next((v for v in cache[ws.title] if v.id == ident), None)
    start, stop = int(first), int(last)
    if view is None or not 0 <= start < stop <= len(view.rows) or stop - start > PAGE_ROWS:
        raise ValueError("layout no longer matches source")
    return ws, view, start, stop, lang, part


def token_value(source, ws, token, lang):
    if token.op == "literal":
        return token.text, token.text, "General", []
    if token.op == "outline_depth":
        level=ws.row_dimensions[ws[token.refs[0]].row].outlineLevel
        return level,str(level),"General",[]
    values = [ws[ref].value for ref in token.refs]
    if token.op == "cell":
        c = ws[token.refs[0]]
        return source.read(ws, c.row, c.column)
    if token.op == "compact":
        text = compact(values[0])
    elif token.op == "exception":
        text = ("학교 예외: " if lang == "ko" else "School exception: ") + str(values[0]).strip()
    elif token.op == "band":
        upper = str(values[1]) if values[1] not in (None, "") else ("상한 미기재" if lang == "ko" else "upper unstated")
        text = f"{values[0]}–{upper}"
    else:
        raise ValueError("unknown token")
    return text, text, "General", []


def table(source, ws, view, start, stop, lang):
    index = 0 if lang == "ko" else 1
    rows = [[(pair[index], pair[index], "General") for pair in view.headers]]
    issues = []
    for line in view.rows[start:stop]:
        cells = []
        for token in line:
            raw, text, fmt, notes = token_value(source, ws, token, lang)
            cells.append((raw, text, fmt))
            issues.extend(notes)
        rows.append(cells)
    return rows, issues
