"""What the checks held back, to see whether right sentences are being lost."""
import json
import sys

results = json.load(open(sys.argv[1], encoding="utf-8"))
for case, entry in results.items():
    held = [s for s in entry.get("sentences", []) if s["problems"]]
    if held:
        print("--", case)
    for sentence in held:
        print("  ", sentence["problems"], sentence["fact_ids"], sentence["statement"][:170])
