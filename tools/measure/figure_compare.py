"""Which workbooks changed between two snapshots, to read by eye."""
import json
import os
import sys

before = json.load(open(sys.argv[1], encoding="utf-8"))
after = json.load(open(sys.argv[2], encoding="utf-8"))
same, changed = 0, []
for path in before:
    a, b = before[path], after.get(path, {"error": "missing"})
    if {k: v for k, v in a.items() if k != "seconds"} == {k: v for k, v in b.items() if k != "seconds"}:
        same += 1
    else:
        changed.append(path)
print("files", len(before), "same", same, "changed", len(changed))
print("slower by >20s", [(os.path.basename(os.path.dirname(p)), before[p]["seconds"], after[p]["seconds"]) for p in before if p in after and after[p]["seconds"] - before[p]["seconds"] > 20])
print("errors before", sum("error" in v for v in before.values()), "after", sum("error" in v for v in after.values()))
print("none before", sum("none" in v for v in before.values()), "after", sum("none" in v for v in after.values()))
read = [sum("error" not in v for v in snapshot.values()) for snapshot in (before, after)]
if not before or 0 in read:  # a comparison of nothing is not a pass
    print("NOTHING READ: before", read[0], "after", read[1])
    sys.exit(1)


def show(entry):
    if "error" in entry or "none" in entry:
        return str({k: entry[k] for k in ("error", "none") if k in entry})
    first = entry["sections"][0]["sentences"][:2] if entry.get("sections") else []
    return f"{entry.get('what', '')[:110]} || {' / '.join(s[:110] for s in first)}"


for path in changed:
    print("--", os.path.basename(os.path.dirname(path)), os.path.basename(path)[:40])
    print("   B", show(before[path]))
    print("   A", show(after.get(path, {"error": "missing"})))
