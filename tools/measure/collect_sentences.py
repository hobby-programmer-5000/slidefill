"""Each case in a process of its own, with a time limit and a memory limit, so one file cannot hold up the rest or fill this PC.

Whole run:  collect_sentences.py <data folder> <output.json>
One case:   collect_sentences.py --case <data folder> <name> <output.json>
"""
import ctypes
import ctypes.wintypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TIME_LIMIT = 420
MEMORY_LIMIT_GB = 6


def watch_memory():
    class Counters(ctypes.Structure):
        _fields_ = [("cb", ctypes.wintypes.DWORD), ("PageFaultCount", ctypes.wintypes.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t), ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t)]
    psapi, kernel32 = ctypes.windll.psapi, ctypes.windll.kernel32
    kernel32.GetCurrentProcess.restype = ctypes.wintypes.HANDLE
    psapi.GetProcessMemoryInfo.argtypes = (ctypes.wintypes.HANDLE, ctypes.POINTER(Counters), ctypes.wintypes.DWORD)
    while True:
        counters = Counters()
        counters.cb = ctypes.sizeof(counters)
        psapi.GetProcessMemoryInfo(kernel32.GetCurrentProcess(), ctypes.byref(counters), counters.cb)
        if counters.WorkingSetSize > MEMORY_LIMIT_GB * 2 ** 30:
            print("MEMORY LIMIT", flush=True)
            os._exit(3)
        time.sleep(1)


def one_case(base, case, output):
    threading.Thread(target=watch_memory, daemon=True).start()
    sys.path.insert(0, ROOT)
    import core
    import insight
    import suggest
    folder = os.path.join(base, case)
    work = tempfile.mkdtemp(prefix="measure-")
    paths = []
    for name in sorted(os.listdir(folder)):
        if name.endswith((".xlsx", ".docx")):
            target = os.path.join(work, ("자료" if name.endswith(".xlsx") else "문서") + os.path.splitext(name)[1])
            shutil.copyfile(os.path.join(folder, name), target)
            paths.append(target)
    source = core.ExcelSource(paths[0]) if len(paths) == 1 else core.Sources(paths)
    plan = suggest.plan_report(source) if len(paths) == 1 else suggest.plan_files(source)
    if plan is None:
        json.dump({"unread": True}, open(output, "w", encoding="utf-8"))
        return
    asks = suggest.ask_insights(source, plan.sections, "자료")
    if not asks:
        json.dump({"no_request": True, "what": plan.what, "sentences_in_report": [i.result for s in plan.sections for i in s.sentences]}, open(output, "w", encoding="utf-8"), ensure_ascii=False)
        return
    request = asks.get(core.ALL_FILES, next(iter(asks.values())))[0]
    pack = request.facts
    prompt = insight.prompt_for(pack, request.lang)
    entry = {"lang": request.lang, "what": plan.what, "model": insight.MODELS[request.lang], "facts": [{"id": f.id, "text": f.text, "figure": f.figure, "period": f.period} for f in pack],
             "calculated": [c.statement for c in insight.calculated_candidates(pack, request.lang)], "report": [i.result for s in plan.sections for i in s.sentences], "sentences": []}
    seen = set()
    started = time.time()
    for seed in range(insight.RUNS):
        reply = insight.post("generate", {"model": insight.MODELS[request.lang], "prompt": prompt, "stream": False, "think": False, "format": insight.schema_for(pack), "keep_alive": "5m",
                                          "options": {"seed": seed, "num_ctx": 4096, "num_predict": 900, "temperature": insight.TEMPERATURE}}, insight.TIMEOUT)
        try:
            items = json.loads(reply.get("response", "")).get("insights", [])
        except (ValueError, AttributeError):
            items = []
        for item in items:
            if not isinstance(item, dict):
                continue
            statement = insight.CITATION_TEXT.sub("", str(item.get("statement", ""))).strip()
            if statement not in seen:
                seen.add(statement)
                item["statement"] = statement
                entry["sentences"].append({"statement": statement, "fact_ids": list(item.get("fact_ids", [])), "problems": insight.check(item, pack, request.lang)})
    entry["model_seconds"] = round(time.time() - started)
    with open(output, "w", encoding="utf-8") as fh:
        json.dump(entry, fh, ensure_ascii=False, indent=1)
    sys.stdout.flush()
    os._exit(0)  # a large workbook can take minutes to free, and the result is already saved


if len(sys.argv) >= 5 and sys.argv[1] == "--case":
    one_case(sys.argv[2], sys.argv[3], sys.argv[4])
    sys.exit(0)

base, output = sys.argv[1:3]
results = {}
for case in sorted(name for name in os.listdir(base) if os.path.isdir(os.path.join(base, name))):
    piece = output + "." + case + ".json"
    started = time.time()
    try:
        done = subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.abspath(__file__), "--case", base, case, piece], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=TIME_LIMIT, cwd=os.path.dirname(output))
        if os.path.exists(piece):
            results[case] = json.load(open(piece, encoding="utf-8"))
        else:
            results[case] = {"failed": (done.stdout + done.stderr)[-300:], "exit": done.returncode}
    except subprocess.TimeoutExpired:
        results[case] = {"timed_out": TIME_LIMIT}
    results[case]["total_seconds"] = round(time.time() - started)
    entry = results[case]
    passed = sum(not s["problems"] for s in entry.get("sentences", []))
    print(f"{case}: {entry['total_seconds']}s lang={entry.get('lang')} wrote={len(entry.get('sentences', []))} passed={passed} {[k for k in ('unread', 'no_request', 'failed', 'timed_out') if k in entry]}", flush=True)
with open(output, "w", encoding="utf-8") as fh:
    json.dump(results, fh, ensure_ascii=False, indent=1)
print("DONE")
