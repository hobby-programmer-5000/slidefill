"""Second stage: calculated facts -> a local model -> checked candidate sentences. No GUI.

The model is never asked to calculate, and its sentences never reach a report by themselves: the person chooses
among the candidates, and core.py only reads what was chosen. This is the only module that makes a network
call, and it only ever talks to the model server on this PC.
"""

import json
import re
import socket
import urllib.error
import urllib.request
from dataclasses import dataclass, field

from facts import HANGUL, LATIN
from numeric_text import NUMBER, NOT_A_FIGURE, figure_tokens, number_identity, unbacked_numbers

URL = "http://127.0.0.1:11434/api/"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # a system proxy must never see this data
# Chosen by testing on real data with 8GB of video memory as the limit. See doc/how-it-works.md.
MODELS = {"ko": "hf.co/mykor/A.X-4.0-Light-gguf:Q4_K_M", "en": "qwen3.5:4b"}
RUNS = 3
MAX_CALCULATED = 8
TIMEOUT = 300
TEMPERATURE = 0.2  # sentences that stay close to the facts; more runs still give some choice

# 원인 alone would also match 지원인원, so it must stand as a word.
CAUSE = re.compile(r"때문|덕분|로 인해|으로 인해|(?<![가-힣])원인|영향|가능성|시사|추정|외부|상관|한 결과|된 결과|의 결과|수 있습니다|수 있다|because|due to|caused by|driven by|resulted|resulting|led to|leading to|thanks to|correlat|suggest|indicat|likely|possibl|\bmay\b|\bmight\b", re.I)
# A figure can be stated without being judged: 19th of 21 is not "still high", and one quarter is not "stable" or "a strong performance".
JUDGED = re.compile(r"안정적|긍정적|부정적|양호|우수|성과|주목할|여전히|상대적으로|활발|모멘텀|심각|집중|두드러|뚜렷|눈에\s*띄|매우|상당히|크게|대폭|소폭|급증|급감|급등|급락|주요|큰\s*(?:차이|격차|폭)|(?:차이|격차)가\s*(?:크(?!지)|큽)|(?<!가장\s)(?<!번째로\s)(?:높은|낮은)\s*수준"
                    r"|\b(?:stable|positive|negative|healthy|strong|robust|impressive|notable|remarkable|significant|substantial|surge|surged|concentrat\w*|dominat\w*|sharp|sharply|very|considerabl[ey]|dramatic|dramatically|wide|widely)\b", re.I)
PERIOD = re.compile(r"\d{4}년\s*(?:\d{1,2}월|\d분기)|Q\d \d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{4}")
COMPARING = re.compile(r"대비|보다|비해|compared|against|than|versus|since|from|above|below|earlier|\bvs\b", re.I)
EXTREME_FACT = re.compile(r"(?:^|_)(?:high|low)$")
# Where a sentence turns to another thought, so each part can be held to one fact.
CLAUSES = re.compile(r",(?!\d)|;|\bwhile\b|\bbut\b|\bwhereas\b|으며|하며|지만|으나|했고|이며|이고|였고|\.\s")
RESPECTIVELY = re.compile(r"각각|respectively", re.I)
RUN_WORD = re.compile(r"연속\s*(?:하락|감소|증가|상승)|(?:rose|risen|fell|fallen|increased|decreased|declined) for \d+ consecutive", re.I)
YEAR_WORD = re.compile(r"전년|같은\s*(?:달|분기)|동월|동분기|year|same quarter|same month", re.I)
WHOLE_FIGURE = re.compile(r"합계|총계|전체|^total|^sum\b|\btotal$", re.I)
WINDOW = re.compile(r"(?:최근|지난|직전)\s*(\d+)\s*(?:개월|개|달|분기|년|주|일)|(?:last|latest|past)\s+(?:\d+\s+)?(\d+)\s+(?:months?|quarters?|years?|weeks?|days?|values?)", re.I)
UP_WORD = re.compile(r"증가|상승|높|늘|\b(?:above|higher|more|up|rose|risen|increase[ds]?|grew|growth)\b", re.I)
DOWN_WORD = re.compile(r"감소|하락|낮|줄|\b(?:below|lower|less|down|fell|fallen|decrease[ds]?|declined?|drop(?:ped)?)\b", re.I)
TOP_ITEM = re.compile(r"(?:1위|number 1)\s+(.+?)\s+[\d,.-]+(?:,|$)")
# Each comparison a sentence can claim. It stands only if a fact it cites makes the same one: the model's habit of writing "the same month a year earlier" for anything must not reach a report.
SAME_PERIOD_LAST_YEAR = re.compile(r"같은\s*(?:달|분기|기간|월)|동월|동분기|동기|same\s+(?:month|quarter|period)|(?:a|one)\s+year\s+(?:earlier|ago)|same\s+\w+\s+last\s+year|year[- ]on[- ]year", re.I)
LAST_WEEK = re.compile(r"지난\s*주|전주|last\s+week|previous\s+week|a\s+week\s+(?:earlier|ago)", re.I)
A_YEAR_BEFORE = re.compile(r"전년|작년|지난해|직전\s*해|last\s+year|previous\s+year|year\s+before|a\s+year", re.I)
BACKS_A_YEAR_BEFORE = re.compile(r"전년|작년|직전\s*해|\d{4}년\s*월평균|12개월치|last\s+year|previous\s+year|year\s+before|a\s+year|12\s+months\s+of|per\s+month\s+in", re.I)
COMPARISONS = ((SAME_PERIOD_LAST_YEAR, SAME_PERIOD_LAST_YEAR), (LAST_WEEK, LAST_WEEK), (A_YEAR_BEFORE, BACKS_A_YEAR_BEFORE))
CURRENT_PERIOD_WORD = re.compile(r"이번\s*기간|current\s+period|this\s+period", re.I)
PART_OF_A_YEAR = re.compile(r"(까지|\s*year\s*to\s*date)$", re.I)
SAYS_LENGTHS_DIFFER = re.compile(r"기간이\s*달|개월치|그대로\s*읽|cannot be read|months of|lengths? differ", re.I)
PERCENT = re.compile(r"[+−-]?(\d[\d,]*(?:\.\d+)?)\s*%")
COMPARED_WITH_NOW = re.compile(r"(?:^|_)(?:prev|year|week|avg\d*|read_prev_year|read_year)$")
ABOUT_NAMED_ITEMS = re.compile(r"(?:^|_)(?:item_\d+|read_top|read_gap|read_top_three|read_range)$")  # 서울 23.9%, or the top 3: the name of the thing compared is in the sentence, and the slide says what was counted
MONTHLY_BASE = re.compile(r"\S*\s*월평균\s*(?:\d[\d,.]*\s*)?(?:대비|보다|에\s*비해)")  # what it is compared with; the thing compared must be a monthly mean as well
A_MULTIPLE = re.compile(r"\d\s*(?:배|times)", re.I)
SECOND_NAME = re.compile(r"2위\s+(.+?)\(|the second, (.+?) \(")
# A name for the figure that came from the sheet's title, or from nowhere. 전체 1,658,715 of 합계 under the title 세대 및 인구 was written up as 세대 수, and it was people.
SHEET_LABEL = re.compile(r"이 수치가 있는 시트의 이름:|시트 제목:|The sheet these figures are on is named:|Its title says:")
SHEET_WORD = re.compile(r"[가-힣]{2,}|[A-Za-z]{3,}")
PLAIN_SHEET_WORD = re.compile(r".*별$|현황|통계|자료|시트|데이터|목록|보고|결과|단위|기준|전체|합계|총계|summary|data|sheet|table|report|total|statistics|and|the|for|with", re.I)
SEVERAL_THINGS = re.compile(r"\s및\s|,|/|\sand\s", re.I)
MADE_UP_MEASURE = re.compile(r"([가-힣]{2,})\s*(?:수|량|액)(?=[는은이가의를을로와과도만]|\s|$|[,.])|(?<=\d)\s*(명|건|톤|달러|kg|km)(?![a-z])|\b(?:number|count|amount|total|sum) of (?:the )?([a-z]+)", re.I)
PLAIN_MEASURE = re.compile(r"항목|items?|rows?|sheet|all|값", re.I)
# A period that had not ended when the file was saved. -88% is mostly the months still to come, so it may only be said with that said.
UNFINISHED_LEAD = re.compile(r"끝나지 않은|had not ended")
SAYS_UNFINISHED = re.compile(r"끝나지 않|아직|덜 쌓|잠정|현재까지|not ended|not over|incomplete|so far|to date|partial", re.I)
RISING_OR_RUN_WORD = re.compile(r"연속|최고|가장 높|증가|상승|늘었|\b(?:in a row|highest|above|higher|rose|risen|rises|rise|up|increased?)\b", re.I)
SIGN_BEFORE_WORD = re.compile(r"[+−-](\d[\d,]*(?:\.\d+)?\s*%)(?=\s*(?:above|below|higher|lower|more|less|up\b|down\b|높|낮|많|적|증가|감소|상승|하락))", re.I)
FALLING_WORD = re.compile(r"최저|가장 낮|감소|하락|낮습|줄었|\b(?:lowest|below|lower|fell|fallen|falls|fall|down|decreased?)\b", re.I)
ITEM_FACT = re.compile(r"(?:^|_)item_\d+$")
ITEM_NAME = re.compile(r"(?:(?:\d+위|number \d+)\s+)?(.+?)\s+(?:실적\s+|came to\s+)?[\d,.-]+(?:,|$|\s)")  # 1위 연구소 19,468, or against a budget 연구소 실적 19,468
SUPERLATIVE = re.compile(r"가장|최고|최대|최저|최소|\b(?:highest|largest|biggest|most|lowest|smallest|least)\b", re.I)
RANKED = re.compile(r"(?<!\d)1\s*(?:위|번째로)|\bnumber 1\b", re.I)  # first place backs a superlative: 최근 12개 값 중 1번째로 높음, 1위 연구소; a third place does not
SUM_IN_A_FACT = re.compile(r"(?:전체|항목\s*합계)\s+(\d[\d,]*(?:\.\d+)?)|(?:items' sum|the total|the whole),?\s+(\d[\d,]*(?:\.\d+)?)", re.I)
SAYS_THE_WHOLE = re.compile(r"전체|합계|총|합은|\b(?:sum|total|whole|all)\b", re.I)
MIN_ITEM_NAME = 2  # 가 alone would be found in almost any sentence
BRACKET = re.compile(r"^\[[^\]]*\]\s*")
FILE_PREFIX = re.compile(r"^file\d+_")
HAN = re.compile(r"[一-鿿]")
LEAK = re.compile(r"신호_|signal_|read_|fact_ids?|fact_\d|item_\d")
CITATION_TEXT = re.compile(r"\s*\(fact_ids?:[^)]*\)")
# 평년 for a recent mean, 17th of 30 as "the top 17%" or "near the top", and a station's figure as the whole system's.
MISREADINGS = (re.compile(r"평년|이상적|전체\s*가치"), re.compile(r"[가-힣]\s*[A-Za-z]{5,}|[A-Za-z]{5,}\s*[가-힣]"), re.compile(r"상위\s*\d+\s*%|상위권|하위권"))
RANK_LOW = re.compile(r"\d+\s*번째로\s*낮")
TREND = re.compile(r"지속|꾸준|추세|모멘텀|\b(?:steady|steadily|sustained|trend)\b", re.I)
RUN = re.compile(r"연속|in a row|a row")

PROMPT = {
    "ko": """당신은 운영 보고서를 검토하는 분석가입니다. 아래 [계산 결과]는 프로그램이 계산해 검증한 값입니다. 다시 계산하지 마세요.
read로 시작하는 항목은 프로그램이 정한 읽는 법이고, signal로 시작하는 항목은 프로그램이 찾아낸 주목할 점입니다. 이 판단을 바꾸지 말고 이것을 우선 사용해, 보고받는 사람이 어떻게 읽어야 하는지 쓰세요. 수치를 나열하지 마세요.

규칙:
1. 통찰은 2개 또는 3개만 씁니다. 각각 한 문장입니다.
2. fact_ids에 근거가 된 사실의 id를 모두 적습니다. 문장의 숫자는 그 사실들에 적힌 숫자만 씁니다. 서로 다른 사실의 숫자를 한 절에 묶지 않습니다.
3. 사실에 적힌 이름과 비교 대상을 바꾸지 않습니다. 예를 들어 '직전 7개 평균'을 '평년'이나 '지난주'로 부르지 않습니다. 비교 대상은 그 숫자가 적힌 사실의 말을 그대로 씁니다. 이번 기간이 적혀 있지 않으면 기간을 말하지 않습니다.
4. 문장마다 subject에 적힌 수치 이름을 그대로 씁니다. 파일 이름만으로 전체를 뜻하는 것처럼 쓰지 않습니다. 시트 이름이나 제목에서 다른 수치 이름을 가져오지 않고, 무엇을 세었는지 짐작해 쓰지 않습니다.
5. 최고, 최저, 연속 같은 말은 이번 기간에 대해서만 씁니다. 비교 대상 기간의 값을 이번 값처럼 쓰지 않습니다. 전체나 합계의 값을 한 항목의 값처럼 쓰지 않습니다. read_lead에 값이 덜 쌓였다고 적혀 있으면 증감을 말할 때 아직 끝나지 않은 기간이라는 말도 함께 씁니다.
6. 좋다, 나쁘다, 안정적이다, 높은 수준이다, 집중도가 높다, 차이가 크다, 두드러진다 같은 평가를 하지 않습니다. 사실에 없는 말로 여러 항목을 묶지 않습니다.
7. 데이터에 없는 원인은 statement에 쓰지 않습니다. '~한 결과', '~때문', '~로 이어져'를 쓰지 않습니다. 날짜 때문에 떠오르는 원인(공휴일 등)은 hypothesis에 "확인 필요:"로 시작해 씁니다. 없으면 빈 문자열입니다.
8. 한국어로만 씁니다.

좋은 형식: {good_form}
나쁜 형식: [수치]는 연속 하락한 결과 낮은 수준입니다. (수치 이름 없음, 전년 대비를 연속 하락으로 바꿈, 원인과 평가를 씀)

[계산 결과]
{facts}""",
    "en": """You review operations reports. The [calculated results] below were worked out and checked by a program. Do not recalculate them.
Items starting with read are how the program decided the figures should be read, and items starting with signal are points it found. Do not change those judgements. Use them first, and say how the reader should read the figures. Do not list the figures.

Rules:
1. Write only 2 or 3 insights, one sentence each.
2. List every supporting fact id in fact_ids. Use only numbers written in those facts. Do not join numbers from different facts in one clause.
3. Keep the names and comparisons used in the facts. For example, do not call "the mean of the previous 7" something else. Say what a number is compared with in the words of the fact that holds it. If no current period is given, do not mention a period.
4. Use the figure's name from the subject item in each sentence. Do not let a file's name stand for the whole. Do not take another name from the sheet's name or title, and do not guess what was counted.
5. Words like highest, lowest and in a row are only about the current period. Do not write a comparison period's value as if it were the current one. Do not write a total or sum as one item's value. If read_lead says the value may be incomplete, say that the period had not ended whenever you state a change.
6. Do not judge: no good, bad, stable, strong, significant, concentrated, or sharp. Do not group things under a word the facts do not use.
7. Do not put causes that are not in the data in the statement. Do not write "resulted in" or "led to". Put a possible cause in hypothesis, starting with "To confirm:". Otherwise use an empty string.
8. Write in English only.

Good form: {good_form}
Bad form: [figure] fell for months as a result, to a low level. (no figure from a fact, says why, judges)

[calculated results]
{facts}""",
}


# A model copies the form it is shown, placeholders and all, so it is shown this figure and this period. A list of companies has no period for it to invent.
GOOD_FORM = {
    "ko": {"dated": "{period} {figure}은 [그 사실에 적힌 비교 대상]보다 [사실의 숫자]% 낮습니다.", "items": "[항목 이름]의 {figure}은 [사실의 숫자]입니다. 기간은 쓰지 않습니다.",
           "several": "[이번 기간] [수치 이름]은 [그 사실에 적힌 비교 대상]보다 [사실의 숫자]% 낮습니다."},
    "en": {"dated": "{figure} in {period} was [number from a fact]% below [the comparison written in that fact].", "items": "{figure} of [an item's name] was [number from a fact]. Do not mention a period.",
           "several": "[figure's name] in [current period] was [number from a fact]% below [the comparison written in that fact]."},
}


def prompt_for(pack, lang):
    """What the model is asked, with the facts and the form that fits them."""
    subjects = [fact for fact in pack if fact.id.endswith("subject")]
    if len(subjects) != 1 or not subjects[0].figure:
        form = GOOD_FORM[lang]["several"]
    else:
        form = GOOD_FORM[lang]["dated" if subjects[0].period else "items"].format(period=subjects[0].period, figure=subjects[0].figure)
    return PROMPT[lang].format(facts="\n".join(f"{fact.id}: {fact.text}" for fact in pack), good_form=form)


class InsightError(Exception):
    """The model could not be reached. code is the end of an i18n key; params fill its message."""

    def __init__(self, code, **params):
        super().__init__(code)
        self.code, self.params = code, params


@dataclass(frozen=True)
class Candidate:
    statement: str
    fact_ids: tuple
    calculated: bool = False  # written by the program from its own results, not by a model


@dataclass
class Candidates:
    model: str
    candidates: list = field(default_factory=list)
    hypotheses: list = field(default_factory=list)
    rejected: int = 0


def post(path, body, timeout=TIMEOUT):
    request = urllib.request.Request(URL + path, json.dumps(body).encode("utf-8"), {"Content-Type": "application/json"})
    try:
        with OPENER.open(request, timeout=timeout) as response:
            return json.load(response)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise InsightError("model_missing", model=body.get("model", "")) from exc
        raise InsightError("other", detail=f"HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, (socket.timeout, TimeoutError)):
            raise InsightError("timeout") from exc
        raise InsightError("ollama_missing") from exc
    except (socket.timeout, TimeoutError) as exc:
        raise InsightError("timeout") from exc
    except ValueError as exc:
        raise InsightError("other", detail=str(exc)) from exc


def schema_for(pack):
    """The id list is part of the schema, so a citation cannot be misspelled."""
    return {"type": "object", "required": ["insights"], "properties": {"insights": {
        "type": "array", "minItems": 2, "maxItems": 3, "items": {
            "type": "object", "required": ["statement", "fact_ids", "hypothesis"], "properties": {
                "statement": {"type": "string"},
                "fact_ids": {"type": "array", "minItems": 1, "items": {"type": "string", "enum": [fact.id for fact in pack]}},
                "hypothesis": {"type": "string"}}}}}}


def _number(text):
    """Exact magnitude identity; sign/direction is checked separately below."""
    return number_identity(text, magnitude=True)


def _with_missing_citations(cited, pack, text):
    """A model often forgets to cite a fact whose figure it used rightly. If one fact holds the figure, it is added; if several do, which was meant cannot be told, so it stays missing."""
    held = {_number(n) for fact in cited for n in fact.numbers}
    added = list(cited)
    for raw in figure_tokens(text):
        value = _number(raw)
        if value is None or value in held:
            continue
        holders = [fact for fact in pack if value in {_number(n) for n in fact.numbers} and not fact.id.endswith(("subject", "files"))]
        current = [fact for fact in holders if fact.id.endswith("value")]  # the current value is also the highest of the last five, and it is still the current value
        chosen = holders if len(holders) == 1 else current
        if len(chosen) == 1:
            added.append(chosen[0])
            held |= {_number(n) for n in chosen[0].numbers}
    return added


def check(item, pack, lang):
    """Reasons to leave a sentence out. This catches wrong citations and known misreadings, not every error of meaning."""
    if not isinstance(item, dict):
        return ["statement"]
    return _check(item.get("statement"), item.get("fact_ids"), pack, lang)[0]


def _check(text, fact_ids, pack, lang):
    """Return rejection reasons and the exact facts used by the checks."""
    if not isinstance(text, str) or not text.strip():
        return ["statement"], ()
    known = {fact.id: fact for fact in pack}
    if (not isinstance(fact_ids, (list, tuple)) or not fact_ids
            or any(not isinstance(name, str) or name not in known for name in fact_ids)):
        return ["citation"], ()
    cited = _with_missing_citations([known[name] for name in dict.fromkeys(fact_ids)], pack, text)
    problems = []
    if unbacked_numbers(text, (n for fact in cited for n in fact.numbers)):
        problems.append("number")
    if CAUSE.search(text):
        problems.append("cause")
    if _judges(text, cited, pack):
        problems.append("judged")
    if _value_of_another(text, pack):
        problems.append("item_value")
    if _borrowed_name(text, pack) or _made_up_measure(text, pack):
        problems.append("name")
    if _reads_unfinished_as_finished(text, pack, cited):
        problems.append("unfinished")
    if _misnames_figure(text, cited, pack):
        problems.append("scope")
    if _wrong_period(text, cited, pack):
        problems.append("period")
    if _joins_facts(text, cited, pack):
        problems.append("join")
    if _unknown_window(text, cited):
        problems.append("window")
    if SUPERLATIVE.search(text) and not any(SUPERLATIVE.search(fact.text) or RANKED.search(fact.text) for fact in cited):
        problems.append("superlative")  # 연구소 +10.3%로 가장 높은 증가율, citing only 연구소's own line, when 관리부 rose 33.0%
    if _wrong_way(text, cited):
        problems.append("direction")
    if _reversed_multiple(text, pack):
        problems.append("multiple")
    if _unbacked_comparison(text, cited) or _no_such_period(text, pack) or _part_of_a_year_cut_short(text, pack) or _unreadable_change(text, pack):
        problems.append("period_word")
    if _old_value_as_current(text, cited, pack):
        problems.append("old_value")
    if RUN_WORD.search(text) and any(fact.id.endswith("read_year_run") and YEAR_WORD.search(fact.text) for fact in cited) and not YEAR_WORD.search(text):
        problems.append("run")  # 6번 연속 전년보다 낮다 is not 6번 연속 하락
    if TREND.search(text) and not any(RUN.search(fact.text) for fact in cited):  # a trend needs a run to stand on, not one month's change
        problems.append("trend")
    if LEAK.search(text):
        problems.append("leak")
    if HAN.search(text) or (lang == "ko" and len(LATIN.findall(text)) > len(HANGUL.findall(text))) or (lang == "en" and HANGUL.search(text)):
        problems.append("language")
    if lang == "ko":
        low_in_fact = any(RANK_LOW.search(fact.text) for fact in cited)
        if any(pattern.search(text) for pattern in MISREADINGS) or (RANK_LOW.search(text) and not low_in_fact):
            problems.append("meaning")
    return problems, tuple(fact.id for fact in cited)


def _reads_unfinished_as_finished(text, pack, cited=()):
    """A change, or a low, in a period the program said had not ended, with nothing to say so."""
    if not any(fact.id.endswith("read_lead") and UNFINISHED_LEAD.search(fact.text) for fact in pack) or SAYS_UNFINISHED.search(text):
        return False
    ended = [fact.period for fact in cited if fact.id.endswith("read_ended") and fact.period]
    def mentions(words, period):
        return re.search(r"(?<!\d)" + re.escape(period).replace(r"\ ", r"\s*") + r"(?!\d)", words, re.I) is not None
    current = [fact.period for fact in pack if fact.id.endswith("subject") and fact.period]
    for period in ended:
        if not mentions(text, period) or any(now != period and mentions(text, now) for now in current):
            continue
        changing = [clause for clause in re.split(r"\band\b|그리고|또한", text, flags=re.I)
                    if PERCENT.search(clause) or FALLING_WORD.search(clause) or RISING_OR_RUN_WORD.search(clause)]
        if changing and all(mentions(clause, period) for clause in changing):
            return False
    return PERCENT.search(text) is not None or FALLING_WORD.search(text) is not None or RISING_OR_RUN_WORD.search(text) is not None


def _plain_sign(text):
    """Remove a redundant sign after both it and the direction word have been checked."""
    return SIGN_BEFORE_WORD.sub(r"\1", text)


def _known_words(pack):
    """What the program gave as names: the figure's and everything in the facts but the sheet's name, which is what a model borrows from."""
    return "".join([fact.figure or "" for fact in pack] + [fact.text for fact in pack if not fact.id.endswith("sheet")]).casefold()


def _borrowed_name(text, pack):
    """세대 수, from a sheet titled 세대 및 인구, about a column headed only 합계. A title naming one thing can name a total, since nothing else does; a title naming several cannot say which, and it cannot rename a figure that has its own name."""
    known = re.sub(r"\s+", "", _known_words(pack))
    lowered = text.casefold()
    specific = [fact for fact in pack if fact.id.endswith("subject") and fact.figure and not WHOLE_FIGURE.search(fact.figure)]
    for fact in pack:
        if not fact.id.endswith("sheet"):
            continue
        title = SHEET_LABEL.sub(" ", BRACKET.sub("", fact.text))
        if not specific and not SEVERAL_THINGS.search(title):
            continue
        for word in SHEET_WORD.findall(title):
            if not PLAIN_SHEET_WORD.fullmatch(word) and word.casefold() in lowered and word.casefold() not in known:
                return True
    return False


def _made_up_measure(text, pack):
    """방문객 수 or 877,691명 when no fact says who was counted. The number may be right; what it counts is a guess."""
    known = re.sub(r"\s+", "", _known_words(pack))
    for found in MADE_UP_MEASURE.finditer(text):
        word = next(group for group in found.groups() if group)
        whole = re.sub(r"\s+", "", found.group(0)).casefold()
        if not PLAIN_MEASURE.fullmatch(word) and word.casefold() not in known and whole not in known:
            return True
    return False


def _judges(text, cited, pack):
    """A word of judgement, unless the facts it cites use that word themselves or the figure is called that: 집중치료 환자수 says nothing about concentration."""
    known = " ".join([fact.text for fact in cited] + [fact.figure for fact in pack if fact.figure]).casefold()
    return any(match.group(0).casefold() not in known for match in JUDGED.finditer(text))


def _value_of_another(text, pack):
    """Woolworths was 4,393,656, when that is the sum of every company, or Woolworths was 117,552, when that is Coles. Only a clause naming one item is judged: with two names, which figure is whose cannot be told."""
    sums, items = set(), {}
    for fact in pack:
        sums |= {_number(found[0] or found[1]) for found in SUM_IN_A_FACT.findall(fact.text)}
        name = ITEM_NAME.search(fact.text) if ITEM_FACT.search(fact.id) else None
        if name and len(name.group(1)) >= MIN_ITEM_NAME:
            items[name.group(1).casefold()] = {_number(n) for n in NUMBER.findall(fact.text[name.end(1):])}
    for clause in CLAUSES.split(text):
        lowered = clause.casefold()
        named = [name for name in items if name in lowered]
        if len(named) != 1:
            continue
        own = items[named[0]]
        figures = {_number(n) for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", clause))} - own
        if figures & sums and not SAYS_THE_WHOLE.search(clause):
            return True
        if figures & (set().union(*items.values()) - sums):
            return True
    return False


def _subject_of(fact, pack):
    prefix = FILE_PREFIX.match(fact.id)
    return next((other for other in pack if other.id == (prefix.group(0) if prefix else "") + "subject"), None)


def _misnames_figure(text, cited, pack):
    """A sentence about a figure that is part of a file must name it: 내국인 192 written as 매출액 192 reads as all sales."""
    plain = re.sub(r"\s+", "", text).casefold()
    one_figure = sum(fact.id.endswith("subject") for fact in pack) == 1
    for fact in cited:
        if one_figure and ABOUT_NAMED_ITEMS.search(fact.id):
            continue
        subject = _subject_of(fact, pack)
        figure = re.sub(r"\([^)]*\)", "", subject.figure).strip() if subject else ""
        if not figure or WHOLE_FIGURE.search(figure):
            continue
        if re.sub(r"\s+", "", figure).casefold() not in plain and not any(len(word) >= 3 and word.casefold() in text.casefold() for word in figure.split()):
            return True
    return False


def _wrong_period(text, cited, pack):
    """A period named as if it were now. It may be the current one, the one an extreme was reached in, or something compared with."""
    current = {re.sub(r"\s+", "", fact.period) for fact in pack if fact.period}
    for clause in CLAUSES.split(text):
        for token in PERIOD.findall(clause):
            if re.sub(r"\s+", "", token) in current or COMPARING.search(clause):
                continue
            if any(token in fact.text and (EXTREME_FACT.search(fact.id) or fact.id.startswith("doc")) for fact in cited):
                continue
            return True
    return False


def _old_value_as_current(text, cited, pack):
    """Q2 2018 was 515,260, when 515,260 is Q2 2017's: a number that a fact gives beside another period, in a clause about the current one with nothing to say it is being compared."""
    current = {re.sub(r"\s+", "", fact.period) for fact in pack if fact.period}
    now = {_number(n) for fact in pack if fact.id.endswith("value") for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", fact.text))}
    for clause in CLAUSES.split(text):
        if COMPARING.search(clause) or not any(re.sub(r"\s+", "", token) in current for token in PERIOD.findall(clause)):
            continue
        for raw in NUMBER.findall(NOT_A_FIGURE.sub(" ", clause)):
            value = _number(raw)
            if value in now:
                continue
            for fact in cited:
                others = [token for token in PERIOD.findall(fact.text) if re.sub(r"\s+", "", token) not in current]
                if others and value in {_number(n) for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", fact.text))}:
                    return True
    return False


def _direction_of(fact, value):
    """Which way a fact says a percentage went: 1, -1, or 0 when it gives no direction, as with a share of a total."""
    for match in PERCENT.finditer(fact.text):
        if _number(match[1]) == value:
            sign = match[0][0]
            if sign in "+−-":
                return 1 if sign == "+" else -1
            up, down = bool(UP_WORD.search(fact.text)), bool(DOWN_WORD.search(fact.text))
            return 1 if up and not down else -1 if down and not up else 0
    return None


def _comparison_direction(fact):
    if fact.relative_change is not None:
        return (fact.relative_change > 0) - (fact.relative_change < 0)
    matches = list(PERCENT.finditer(fact.text))
    if len(matches) == 1 and (value := _number(matches[0][1])) and value[1] != "0":
        return _direction_of(fact, value)
    return None


def _wrong_way(text, cited):
    """Match percentage signs and verbal directions to the stated comparison."""
    for clause in CLAUSES.split(text):
        if RESPECTIVELY.search(clause):
            continue
        up, down = bool(UP_WORD.search(clause)), bool(DOWN_WORD.search(clause))
        said = 0 if up == down else 1 if up else -1
        bases = [(match, backed) for named, backed in COMPARISONS if (match := named.search(clause))]
        matching = [fact for fact in cited if all(basis.search(fact.text) for _, basis in bases)]
        relevant = matching or cited
        percents = list(PERCENT.finditer(clause))
        if not percents and said and any(COMPARING.search(clause[:match.start()]) or
                                         re.search(r"대비|보다|비해", clause[match.end():]) for match, _ in bases):
            ways = {_comparison_direction(fact) for fact in matching if COMPARED_WITH_NOW.search(fact.id)} - {None}
            if len(ways) == 1 and said not in ways:
                return True
        for match in percents:
            value = _number(match[1])
            ways = {_direction_of(fact, value) for fact in relevant if not fact.id.startswith("cross_")} - {None}
            if any(fact.id.startswith("cross_") and value in {_number(n) for n in fact.numbers} for fact in relevant):
                continue
            if not ways:
                return True
            sign = match[0][0]
            if value and value[1] != "0" and sign in "+−-" and (1 if sign == "+" else -1) not in ways:
                return True
            if said and said not in ways:
                return True
    return False


def _reversed_multiple(text, pack):
    """교육부 is 보건복지부의 1.3배 had the largest item as the smaller one. A multiple is always of the second, so the largest cannot be what it is a multiple of."""
    top = next((fact for fact in pack if fact.id == "item_1" or fact.id.endswith("_item_1")), None)
    words = TOP_ITEM.search(top.text) if top else None
    if not words:
        return False
    name = words.group(1).strip()
    if re.search(rf"{re.escape(name)}\S{{0,3}}\s*(?:\S+\s*)?의\s*[\d.]+\s*배|[\d.]+\s*times\s+(?:that of\s+|the\s+\w+\s+(?:of|for|in)\s+)?{re.escape(name)}", text, re.I):
        return True
    # VIC was 1.2 times the second, when VIC is the second: the only multiple calculated is the largest's, so it cannot be about another item alone.
    others = [found.group(1).casefold() for fact in pack if fact is not top and ITEM_FACT.search(fact.id) and (found := ITEM_NAME.search(fact.text)) and len(found.group(1)) >= MIN_ITEM_NAME]
    others += [(found.group(1) or found.group(2)).casefold() for fact in pack if fact.id.endswith("read_gap") and (found := SECOND_NAME.search(fact.text))]  # with several files, only the first item is listed
    lowered = text.casefold()  # the whole sentence: ", which is 1.4 times the second" leaves its name behind the comma
    return A_MULTIPLE.search(text) is not None and name.casefold() not in lowered and any(other in lowered for other in others)


def _unbacked_comparison(text, cited):
    """Yearly figures have no same month last year, and 6.4% against last week is not 6.4% against a year ago. A comparison named in a sentence must be one a cited fact makes."""
    return any(said.search(text) and not any(backed.search(fact.text) for fact in cited) for said, backed in COMPARISONS)


def _no_such_period(text, pack):
    """A list of companies has no current period for a sentence to be about."""
    return CURRENT_PERIOD_WORD.search(text) is not None and not any(fact.period for fact in pack)


def _part_of_a_year_cut_short(text, pack):
    """2022년 10월까지 is ten months. Written as 2022년 10월 it reads as that one month."""
    plain = re.sub(r"\s+", "", text)
    for fact in pack:
        whole = re.sub(r"\s+", "", fact.period or "")
        stem = PART_OF_A_YEAR.sub("", whole)
        if stem != whole and stem in plain and whole not in plain:
            return True
    return False


def _unreadable_change(text, pack):
    """Ten months against twelve: the program said this percentage cannot be read as it stands, so it may only be repeated with that said."""
    for fact in pack:
        if fact.id.endswith("read_partial_year"):
            shown = PERCENT.findall(fact.text)
            said = {_number(raw) for raw in PERCENT.findall(text)}
            if shown and _number(shown[0]) in said and not SAYS_LENGTHS_DIFFER.search(text):
                return True
            if shown and _number(shown[-1]) in said and "per month" not in text.casefold() and "월평균" not in MONTHLY_BASE.sub(" ", text):
                return True  # 외래연인원은 2021년 월평균 대비 4.2% 감소: the ten months' total is not 4.2% below one month
    return False


def _unknown_window(text, cited):
    """A rank or extreme over a stretch nobody calculated. "5th of the last 6 months" came with no such result behind it, and was wrong."""
    for match in WINDOW.finditer(text):
        size = match.group(1)
        if not any(size in NUMBER.findall(fact.text) and (WINDOW.search(fact.text) or "개" in fact.text or "latest" in fact.text.lower() or "last" in fact.text.lower()) for fact in cited):
            return True
    return False


def _joins_facts(text, cited, pack=()):
    """Figures in one clause must come from one fact. $64 million across 8,019 estates took one from each, and the document never said so.

    The exception is the current value beside a comparison made with it: 739,226, against 714,365 the year before, +3.5%. A monthly mean was not compared with it, so it has no such place.
    """
    allowed = {_number(n) for fact in cited for n in fact.numbers}
    now = {_number(n) for fact in pack if fact.id.endswith("value") for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", fact.text))}
    for clause in CLAUSES.split(text):
        if RESPECTIVELY.search(clause):
            continue
        figures = {_number(n) for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", clause))} & allowed  # a figure no fact holds is another problem, already counted
        if len(figures) < 2 or any(figures <= {_number(n) for n in fact.numbers} for fact in cited):
            continue
        if not any(figures - now <= {_number(n) for n in fact.numbers} and COMPARED_WITH_NOW.search(fact.id) for fact in cited):
            return True
    return False


def calculated_candidates(pack, lang):
    """Sentences from the program's own results, which cannot misread them: what was calculated across files, what a document lines up with, and each figure's readings. They need no model and no waiting."""
    found, seen = [], set()
    subjects = {prefix.group(0) if (prefix := FILE_PREFIX.match(fact.id)) else "": fact for fact in pack if fact.id.endswith("subject")}

    def add(text, fact):
        if text not in seen:
            seen.add(text)
            found.append(Candidate(text, (fact.id,), True))

    for fact in pack:
        if fact.id.startswith("cross_"):
            add(fact.text, fact)
    for fact in pack:
        if re.match(r"doc\d+_link", fact.id):
            add(BRACKET.sub("", fact.text), fact)
    for fact in pack:
        prefix = FILE_PREFIX.match(fact.id)
        name = fact.id[prefix.end():] if prefix else fact.id
        if name.startswith("read_") and name != "read_separate":
            subject = subjects.get(prefix.group(0) if prefix else "")
            lead = f"{subject.period} {subject.figure}".strip() if subject else ""
            body = BRACKET.sub("", fact.text)
            warning = next((BRACKET.sub("", other.text) for other in pack if other.id == (prefix.group(0) if prefix else "") + "read_lead" and UNFINISHED_LEAD.search(other.text)), "")
            if warning and name != "read_lead" and (PERCENT.search(body) or FALLING_WORD.search(body)):  # someone may choose this sentence alone
                body = f"{warning} {body}"
            add(f"{lead}: {body}" if lead else body, fact)
    return found[:MAX_CALCULATED]


def make_candidates(pack, lang, runs=RUNS, first_seed=0, on_run=None):
    """Ask the model for the given language several times, and keep the sentences that pass the checks.

    on_run(done, runs, what has passed so far) is called after each answer, since one can take a minute and someone is waiting.
    """
    model = MODELS[lang]
    result = Candidates(model)
    prompt = prompt_for(pack, lang)
    seen = set()
    for done, seed in enumerate(range(first_seed, first_seed + runs), start=1):
        reply = post("generate", {"model": model, "prompt": prompt, "stream": False, "think": False, "format": schema_for(pack),
                                  "keep_alive": "5m", "options": {"seed": seed, "num_ctx": 4096, "num_predict": 900, "temperature": TEMPERATURE}}, TIMEOUT)
        try:
            items = json.loads(reply.get("response", "")).get("insights", [])
        except (ValueError, AttributeError, TypeError):
            items = []
        if not isinstance(items, list):
            items = []
        for item in items:
            if not isinstance(item, dict):
                result.rejected += 1
                continue
            text = item.get("statement")
            if isinstance(text, str):
                text = CITATION_TEXT.sub("", text).strip()
            problems, cited = _check(text, item.get("fact_ids"), pack, lang)
            if problems:
                result.rejected += 1
                continue
            text = _plain_sign(text)
            if text not in seen:
                seen.add(text)
                result.candidates.append(Candidate(text, cited))
            hypothesis = item.get("hypothesis")
            if isinstance(hypothesis, str) and (hypothesis := hypothesis.strip()) and hypothesis not in result.hypotheses:
                result.hypotheses.append(hypothesis)
        if on_run:
            on_run(done, runs, Candidates(model, list(result.candidates), list(result.hypotheses), result.rejected))
    return result
