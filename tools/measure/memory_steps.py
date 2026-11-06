"""Memory at each step of planning one workbook. Stops itself past a limit, and reports merged ranges, the likeliest cause."""
import ctypes
import ctypes.wintypes
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import core
import suggest

LIMIT_GB = 6


class Counters(ctypes.Structure):
    _fields_ = [("cb", ctypes.wintypes.DWORD), ("PageFaultCount", ctypes.wintypes.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t), ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t)]


psapi, kernel32 = ctypes.windll.psapi, ctypes.windll.kernel32
kernel32.GetCurrentProcess.restype = ctypes.wintypes.HANDLE
psapi.GetProcessMemoryInfo.argtypes = (ctypes.wintypes.HANDLE, ctypes.POINTER(Counters), ctypes.wintypes.DWORD)


def gb():
    counters = Counters()
    counters.cb = ctypes.sizeof(counters)
    psapi.GetProcessMemoryInfo(kernel32.GetCurrentProcess(), ctypes.byref(counters), counters.cb)
    return counters.WorkingSetSize / 2 ** 30, counters.PeakWorkingSetSize / 2 ** 30


def guard():
    while True:
        if gb()[0] > LIMIT_GB:
            import traceback
            frames = sys._current_frames()
            main = frames[threading.main_thread().ident]
            print("OVER LIMIT at:", flush=True)
            print("".join(traceback.format_stack(main)[-8:]), flush=True)
            os._exit(3)
        time.sleep(0.5)


threading.Thread(target=guard, daemon=True).start()
started = time.time()
source = core.ExcelSource(sys.argv[1])
print(f"read {time.time() - started:.0f}s now {gb()[0]:.2f} GB peak {gb()[1]:.2f} GB", flush=True)
for ws in source.wb.worksheets:
    ranges = list(ws.merged_cells.ranges)
    biggest = max(((r.max_row - r.min_row + 1) * (r.max_col - r.min_col + 1) for r in ranges), default=0)
    print(ws.title[:30], "cells", len(ws._cells), "range", ws.max_row, ws.max_column, "merged", len(ranges), "biggest merged", biggest, flush=True)
started = time.time()
plan = suggest.plan_report(source)
print(f"plan {time.time() - started:.0f}s now {gb()[0]:.2f} GB peak {gb()[1]:.2f} GB", plan.what if plan else None, flush=True)
os._exit(0)
