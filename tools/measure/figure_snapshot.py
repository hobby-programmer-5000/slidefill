"""What the app would put in a report for every workbook, with no model. Each file in a process of its own, so changes can be compared file by file.

Whole run:  figure_snapshot.py <code folder> <output.json> <folder> [<folder> ...]
One file:   figure_snapshot.py --one <code folder> <file> <output.json>
"""
import ctypes
import ctypes.wintypes
import glob
import json
import os
import subprocess
import sys
import threading
import time

TIME_LIMIT = 500
MEMORY_LIMIT_GB = 6
MAX_BYTES = 15_000_000
KINDS = (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv", ".tsv")


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
            os._exit(3)
        time.sleep(1)


def one(code, path, output):
    threading.Thread(target=watch_memory, daemon=True).start()
    sys.path.insert(0, code)
    import core
    import suggest
    entry = {}
    try:
        source = core.ExcelSource(path)
        plan = suggest.plan_report(source)
        if plan is None:
            entry["none"] = True
        else:
            entry["what"] = plan.what
            entry["sections"] = [{"sheet": s.sheet, "figures": [f.name for f in s.figures], "sentences": [i.result[:150] for i in s.sentences]} for s in plan.sections]
    except Exception as exc:
        entry["error"] = repr(exc)[:200]
    with open(output, "w", encoding="utf-8") as fh:
        json.dump(entry, fh, ensure_ascii=False)
    sys.stdout.flush()
    os._exit(0)


if sys.argv[1] == "--one":
    one(*sys.argv[2:5])
    sys.exit(0)

code, output, folders = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), [os.path.abspath(folder) for folder in sys.argv[3:]]  # each file runs from the output's folder, so a relative path would find nothing
paths = []
for folder in folders:
    for path in sorted(glob.glob(os.path.join(folder, "**", "*"), recursive=True)):
        if path.lower().endswith(KINDS) and os.path.getsize(path) <= MAX_BYTES and not os.path.basename(path).startswith("~"):
            paths.append(path)
print("files", len(paths), flush=True)
results = {}
piece = output + ".piece"
for number, path in enumerate(paths, 1):
    if os.path.exists(piece):
        os.remove(piece)
    started = time.time()
    try:
        subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.abspath(__file__), "--one", code, path, piece], capture_output=True, timeout=TIME_LIMIT, cwd=os.path.dirname(output))
        results[path] = json.load(open(piece, encoding="utf-8")) if os.path.exists(piece) else {"error": "stopped (memory limit or crash)"}
    except subprocess.TimeoutExpired:
        results[path] = {"error": "time limit"}
    results[path]["seconds"] = round(time.time() - started)
    if number % 20 == 0:
        print(number, flush=True)
with open(output, "w", encoding="utf-8") as fh:
    json.dump(results, fh, ensure_ascii=False, indent=1)
print("DONE", flush=True)
