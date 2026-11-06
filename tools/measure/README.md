# Measuring accuracy

How to tell whether a change helped. A score on files the rules were tuned on says little: 41 of 43 on known files was 4 of 10 on unseen ones.

The rule:

1. Use public files never used before. Choose them by a rule, before opening them. Put each in its own folder, as `<folder>/<case>/data.xlsx` with an optional `data.docx`.
2. Freeze the code and note its hashes.
3. Collect real model output, then score every passed sentence by hand against the source cells, before changing any code.
4. Check the figures apart from the app, reading the cells directly. Do this for what the report states, not only for what the model wrote.
5. Report the whole score next to the score on files whose figure was chosen rightly. The two have very different causes.
6. Once fixes are made from a batch, that batch is tuning data. The next fair score needs new files.

The source files are other people's data and stay out of this repository.

## Scripts

Run them with the project's Python and `-I`, since they read downloaded files. The workbooks are read with SlideFill's own loader, which leaves pivot caches unread. Opening some public workbooks with plain openpyxl passes 6 GB.

| Script | What it does |
|---|---|
| `collect_sentences.py <folder> <out.json>` | Real model sentences for each case, one process per case with time and memory limits. Needs Ollama |
| `show_sentences.py <out.json> <sources.json>` | Every passed sentence beside the facts it was given, for scoring |
| `show_held.py <out.json>` | What the checks held back, to spot right sentences being lost |
| `figure_snapshot.py <code folder> <out.json> <folder> ...` | What a copy of the code would put in a report for every workbook. No model |
| `figure_compare.py <before.json> <after.json>` | Which workbooks changed between two snapshots, to read by eye |
| `doubt_survey.py <out.json> <folder> ...` | Which workbooks the app would ask a person about |
| `memory_steps.py <workbook>` | Memory at each step of planning one workbook. Stops itself at 6 GB |

To check a change for side effects, copy the `.py` files to another folder before changing anything, take a snapshot from that copy, make the change, take another, and compare.
