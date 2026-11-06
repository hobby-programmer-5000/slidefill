"""Every sentence that passed the checks, next to the facts it was given, for scoring by hand."""
import json
import os
import sys

results = json.load(open(sys.argv[1], encoding="utf-8"))
sources = json.load(open(sys.argv[2], encoding="utf-8"))
for case, entry in results.items():
    print("=" * 8, case, "|", sources.get(case, {}).get("title"), "|", sources.get(case, {}).get("licence"))
    print("lang", entry.get("lang"), "| what", str(entry.get("what"))[:150])
    for key in ("unread", "no_request", "failed", "timed_out", "exit"):
        if key in entry:
            print(key.upper(), entry[key])
    for fact in entry.get("facts", []):
        print("  F", fact["id"], fact["text"][:230])
    for statement in entry.get("calculated", []):
        print("  CALC", statement[:200])
    for sentence in entry.get("sentences", []):
        if not sentence["problems"]:
            print("  PASS", sentence["fact_ids"], sentence["statement"])
    held = {}
    for sentence in entry.get("sentences", []):
        for problem in sentence["problems"]:
            kind = problem if isinstance(problem, str) else str(problem)
            kind = kind.split(":")[0]
            held[kind] = held.get(kind, 0) + 1
    print("  held", held, "of", len(entry.get("sentences", [])), "seconds", entry.get("model_seconds"))
