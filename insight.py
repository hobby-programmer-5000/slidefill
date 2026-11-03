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

from facts import HANGUL, LATIN, NUMBER

URL = "http://127.0.0.1:11434/api/"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # a system proxy must never see this data
# Chosen by testing on real data with 8GB of video memory as the limit. See doc/how-it-works.md.
MODELS = {"ko": "hf.co/mykor/A.X-4.0-Light-gguf:Q4_K_M", "en": "qwen3.5:4b"}
RUNS = 3
TIMEOUT = 300

CAUSE = re.compile(r"때문|덕분|로 인해|으로 인해|원인|영향|가능성|시사|추정|외부|because|due to|caused by|driven by|suggest|indicat|likely|possibl", re.I)
HAN = re.compile(r"[一-鿿]")
NOT_A_FIGURE = re.compile(r"(?<![\d.,])\d+\s*(?:월|일|개|번|주|년|위)|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? \d+|(?<![\d.,])\d{4}(?![\d,.])")
LEAK = re.compile(r"신호_|signal_|read_|fact_ids?|fact_\d|item_\d")
CITATION_TEXT = re.compile(r"\s*\(fact_ids?:[^)]*\)")
MISREADINGS = (re.compile(r"평년|이상적"), re.compile(r"[가-힣]\s*[A-Za-z]{5,}|[A-Za-z]{5,}\s*[가-힣]"))
RANK_LOW = re.compile(r"\d+\s*번째로\s*낮")

PROMPT = {
    "ko": """당신은 운영 보고서를 검토하는 분석가입니다. 아래 [계산 결과]는 프로그램이 계산해 검증한 값입니다. 다시 계산하지 마세요.
read로 시작하는 항목은 프로그램이 정한 읽는 법이고, signal로 시작하는 항목은 프로그램이 찾아낸 주목할 점입니다. 이 판단을 바꾸지 말고 이것을 우선 사용해, 보고받는 사람이 어떻게 읽어야 하는지 쓰세요. 수치를 나열하지 마세요.

규칙:
1. 통찰은 2개 또는 3개만 씁니다. 각각 한 문장입니다.
2. fact_ids에 근거가 된 사실의 id를 모두 적습니다. 문장의 숫자는 그 사실들에 적힌 숫자만 씁니다.
3. 사실에 적힌 이름을 바꾸지 않습니다. 예를 들어 '직전 7개 평균'을 '평년'이나 '지난주'로 부르지 않습니다.
4. 데이터에 없는 원인은 statement에 쓰지 않습니다. 날짜 때문에 떠오르는 원인(공휴일 등)은 hypothesis에 "확인 필요:"로 시작해 씁니다. 없으면 빈 문자열입니다.
5. 한국어로만 씁니다.

[계산 결과]
{facts}""",
    "en": """You review operations reports. The [calculated results] below were worked out and checked by a program. Do not recalculate them.
Items starting with read are how the program decided the figures should be read, and items starting with signal are points it found. Do not change those judgements. Use them first, and say how the reader should read the figures. Do not list the figures.

Rules:
1. Write only 2 or 3 insights, one sentence each.
2. List every supporting fact id in fact_ids. Use only numbers written in those facts.
3. Keep the names used in the facts. For example, do not call "the mean of the previous 7" something else.
4. Do not put causes that are not in the data in the statement. Put a possible cause in hypothesis, starting with "To confirm:". Otherwise use an empty string.
5. Write in English only.

[calculated results]
{facts}""",
}


class InsightError(Exception):
    """The model could not be reached. code is the end of an i18n key; params fill its message."""

    def __init__(self, code, **params):
        super().__init__(code)
        self.code, self.params = code, params


@dataclass(frozen=True)
class Candidate:
    statement: str
    fact_ids: tuple


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
    return f"{float(text.replace(',', '').rstrip('.')):g}"


def check(item, pack, lang):
    """Reasons to leave a sentence out. This catches wrong citations and known misreadings, not every error of meaning."""
    known = {fact.id: fact for fact in pack}
    text = item.get("statement", "")
    problems = []
    cited = [known[name] for name in item.get("fact_ids", []) if name in known]
    if not cited:
        problems.append("citation")
    allowed = {_number(n) for fact in cited for n in fact.numbers}
    if {_number(n) for n in NUMBER.findall(NOT_A_FIGURE.sub(" ", text))} - allowed:
        problems.append("number")
    if CAUSE.search(text):
        problems.append("cause")
    if LEAK.search(text):
        problems.append("leak")
    if HAN.search(text) or (lang == "ko" and len(LATIN.findall(text)) > len(HANGUL.findall(text))) or (lang == "en" and HANGUL.search(text)):
        problems.append("language")
    if lang == "ko":
        low_in_fact = any(RANK_LOW.search(fact.text) for fact in cited)
        if any(pattern.search(text) for pattern in MISREADINGS) or (RANK_LOW.search(text) and not low_in_fact):
            problems.append("meaning")
    return problems


def make_candidates(pack, lang, runs=RUNS, first_seed=0):
    """Ask the model for the given language several times, and keep the sentences that pass the checks."""
    model = MODELS[lang]
    result = Candidates(model)
    prompt = PROMPT[lang].format(facts="\n".join(f"{fact.id}: {fact.text}" for fact in pack))
    seen = set()
    for seed in range(first_seed, first_seed + runs):
        reply = post("generate", {"model": model, "prompt": prompt, "stream": False, "think": False, "format": schema_for(pack),
                                  "keep_alive": "5m", "options": {"seed": seed, "num_ctx": 4096, "num_predict": 900, "temperature": 0.4}}, TIMEOUT)
        try:
            items = json.loads(reply.get("response", "")).get("insights", [])
        except (ValueError, AttributeError):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            item["statement"] = CITATION_TEXT.sub("", str(item.get("statement", ""))).strip()
            if check(item, pack, lang):
                result.rejected += 1
                continue
            if item["statement"] not in seen:
                seen.add(item["statement"])
                result.candidates.append(Candidate(item["statement"], tuple(item.get("fact_ids", []))))
            hypothesis = str(item.get("hypothesis", "")).strip()
            if hypothesis and hypothesis not in result.hypotheses:
                result.hypotheses.append(hypothesis)
    return result
