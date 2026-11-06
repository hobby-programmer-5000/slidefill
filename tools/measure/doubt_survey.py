"""Which workbooks the app would ask about, against what I found by checking them against their sources. One process per file; memory and time are limited by figure_snapshot's limits being copied here."""
import ctypes
import glob
import json
import os
import subprocess
import sys
import threading
import time


def one(path, output):
    def guard():
        import ctypes.wintypes
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
            if counters.WorkingSetSize > 6 * 2 ** 30:
                os._exit(3)
            time.sleep(1)
    threading.Thread(target=guard, daemon=True).start()
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    import core
    import suggest
    entry = {}
    try:
        plan = suggest.plan_report(core.ExcelSource(path))
        entry = {"none": True} if plan is None else {"what": plan.what, "doubts": [code for code, _p in plan.doubts]}
    except Exception as exc:
        entry = {"error": repr(exc)[:120]}
    json.dump(entry, open(output, "w", encoding="utf-8"), ensure_ascii=False)
    sys.stdout.flush()
    os._exit(0)


if sys.argv[1] == "--one":
    one(sys.argv[2], sys.argv[3])
    sys.exit(0)

output, folders = os.path.abspath(sys.argv[1]), [os.path.abspath(folder) for folder in sys.argv[2:]]  # each file runs from the output's folder, so a relative path would find nothing
paths = [p for folder in folders for p in sorted(glob.glob(os.path.join(folder, "*", "*.xlsx"))) if os.path.getsize(p) <= 15_000_000]
results, piece = {}, output + ".piece"
for path in paths:
    if os.path.exists(piece):
        os.remove(piece)
    try:
        subprocess.run([sys.executable, "-I", "-W", "ignore", "-X", "utf8", os.path.abspath(__file__), "--one", path, piece], capture_output=True, timeout=500, cwd=os.path.dirname(output))
        results[path] = json.load(open(piece, encoding="utf-8")) if os.path.exists(piece) else {"error": "stopped"}
    except subprocess.TimeoutExpired:
        results[path] = {"error": "time limit"}
    entry = results[path]
    print(os.path.basename(os.path.dirname(path))[:34].ljust(36), "|", str(entry.get("doubts", entry.get("error", "none"))), "|", str(entry.get("what", ""))[:60], flush=True)
json.dump(results, open(output, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("DONE", flush=True)
if not paths or all("error" in entry for entry in results.values()):
    print("NOTHING READ: no file found, or every file failed", flush=True)
    sys.exit(1)
