**English** | [한국어](verification.ko.md)

# CPU verification

This guide describes the **unreleased development source**, including [numeric reliability hardening](reliability-hardening.md). It does not add those changes to the published v0.3.0 executable. Historical validation applies to the artifact recorded with it.

Verification runs **automatically inside the app** for filled text, tables and charts, including structured rate/organization reports. No local model, GPU, separate checker window or user command is needed. Preview and generation use the same `core.run` guard.

## In the app

- **Passed:** the declared supported inputs, calculation checks and saved representations passed. Supported source recipes and formula caches are independently evaluated. Selected business scope and interpreted labels still need review; this is not certification of all prose or Office layout.
- **Partial:** a source format, formula or expression is outside the independent checker. The app states the gap and **refuses saving**. A known mismatch elsewhere takes precedence. Use the reason to correct or simplify the input; the ordinary confirmation cannot override incomplete coverage.
- **Blocked:** source drift, missing/changed values, categories, series, cache points, embedded data or required source warnings were found. Generate is disabled for that preview. Reopen the input and preview again. The ordinary “create anyway” path cannot override this check.

The guard checks a serialized preview, stages the final file, rechecks its bytes and source fingerprints, and only then publishes a new output. Existing output files are not overwritten. No sidecar certificate or extra copy of the source is created in normal app use. The preview's source snapshot is retained for its save operation.

Text-only reports also use this guard. Changed paragraphs are compared after serialization; direct/latest/rank values and supported numeric comparison/extreme placeholders have source checks. Source evidence accompanies prose and optional model text, but not every semantic claim is independently proved. Word-linked before/after means, period windows, quotation and change text are reconstructed from workbook/DOCX evidence. Used saved formulas and dependencies are checked against a bounded exact evaluator; stale supported caches block saving, while unsupported formulas cause partial coverage and also prevent saving. See the [grammar, precision and text boundaries](reliability-hardening.md).

Category, date-axis and XY charts are checked against their series names, categories/coordinates, gaps, point counts, caches and the cells referenced in the embedded XLSX. Direct ranges, ordinary latest series, split dates, grouped sums/means/counts/rates/products, rankings and unambiguous share denominators have source recipes. Selection, hierarchy interpretation and number formatting remain tested/trusted boundaries. Formats without an independent reader and ambiguous reconstructed expressions produce an explicit partial state. Bubble charts accept an explicit X column followed by Y/size pairs; size is nonnegative and each series needs a complete numeric point. Unreadable/incomplete points are omitted consistently, and the size caches and embedded cells are checked too. Independent source readers cover XLSX-family files, renamed XLSX exports, CSV/TSV/TXT and the supported ODS value subset. Supported BIFF2–5/8 XLS and XLSB cells/caches now have independent decoders; unsupported binary records/workspace variants, unsupported ODS duration grammar and other unknown expressions remain partial. HTML/delimited XLS exports are also covered. See [binary, grouping and rounding boundaries](binary-verification.md). Delimited/ODS reconstruction has a two-million-cell budget; it does not use the production spreadsheet reader.

Integer comparisons are exact. Decimal/float serialization has a two-ULP budget; explicit grouped aggregates retain exact rational intermediates through formatting or a required container conversion, independently reconstructed from source operands. A chart integer must survive float conversion exactly, and explicit supported formats must keep their displayed precision in the cache/embedded data. There is no operation-count ULP multiplier for aggregates. See the [aggregate contract](binary-verification.md#exact-aggregates-and-source-decoding-error). These are enforced comparison policies, not a formal floating-point error theorem or an independent proof of the shared display formatter.

## Optional insight and contents-sketch checks

The optional insight editor and core reuse check supported numeric magnitudes against the current calculated facts using bounded exact decimal identities. Equivalent spellings can match, while differing low digits, malformed forms and excessive literals are rejected. An invalid edit keeps the prior choice and leaves the editor open. Use a number from the displayed facts, or clear the field for the calculated sentence. An invalid stored choice is replaced with that fallback and a numeric warning. Signs/direction words, units, subjects, excluded dates/quarters and arbitrary sentence meaning remain separate. Scientific-literal support in the checker does not guarantee complete scientific metadata from every fact producer; conservative rejection is possible.

Cross-file direction uses exact relative-change payloads, including tiny nonzero signs; a move below one-decimal display precision is described as less than `0.05%`. It never reconstructs the change from rounded prose. The **Change contents** chart sketch preserves finite Fraction data, missing gaps and mixed-sign zero baselines, with exact normalization before screen-coordinate conversion. The sketch remains a schematic. See the [boundary contracts](reliability-hardening.md#preserving-information-between-calculations-and-wording) and [research rationale](research-methods.md). Manual content selections still lack a complete schema-drift contract.

## Special cases and notes in the report

| Situation | Existing app/report behavior | What the reader still needs to check |
|---|---|---|
| Rate schedule with an absent upper bound or a conflicting total | Retains the original values and writes the unknown bound or source-versus-component warning in the structured report note. Missing bounds are not inferred as infinity. | Which band or named exception applies; the app does not repair the source total. |
| Organization totals beside departments/job types | Keeps source records and descriptive notes; does not infer a hierarchy or add parents and children together. | Which records belong to the intended business scope. |
| Unknown fiscal-year end, incomplete periods or an unconfirmed automatic choice | Uses the existing qualified sentences/notes or withholds a comparison. | The actual year end, covered dates and chosen measure. |
| Unsupported independent source format/formula/expression | Shows partial verification and refuses saving; checks it can perform still run. | Correct or simplify the input using the stated reason. Re-preview; there is no generic override. |
| Changed input, altered chart/cache/embedded workbook or missing required source warning | Blocks saving; the generic “create anyway” confirmation cannot override it. | Reopen the intended input, correct the cause and preview again. |
| A chosen insight contains an unsupported or unbacked numeric token | The editor rejects the edit; core reuse substitutes the calculated fallback and adds a numeric warning. | Correct the token using the current facts, or clear the choice. Read the sentence for signs, units and meaning. |

For example, if food is 10, operations is 20 and the source total is 35, the report retains 35 and notes that the components add to 30. A passed check can mean that these values and the required warning were correctly preserved; it does not certify the source total as correct.

Report notes preserve information a recipient needs to interpret the figures, including supported structured-source findings. Not every app warning is printed on a slide. Automatic passed/partial/blocked status is displayed in the app; there is no universal “mathematically certified report” stamp. Normal generation does not append proof code, internal source recipes or a certificate to the PPTX. A recipient does not need a checker. This documentation describes the existing behavior, not a new embedding feature.

For why the checks exist, the mathematical algorithms, code ownership and maintenance checks, see [verification design](verification-design.md). Formal kernel reproduction is documented [separately](../verification/proofs/README.md).

## Developer reproduction commands

From the project directory:

```powershell
& .\.venv\Scripts\python.exe -X utf8 -m verification selftest
& .\.venv\Scripts\python.exe -X utf8 -m verification generate C:\data\rates.xlsx C:\data\check-001 --lang en
& .\.venv\Scripts\python.exe -X utf8 -m verification check C:\data\rates.xlsx C:\data\check-001\report.pptx C:\data\check-001\certificate.json
```

Choose a **new output directory**. `generate` refuses an existing one and creates `template.pptx`, `report.pptx`, `certificate.json` and `checks.json`. Inputs are read-only. Treat certificates and reports as private if the source is private. Exit0 means the declared checks passed, exit1 means a failed check, exit2 means the command could not run. A failed command can leave evidence in its new output directory; inspect it and use another directory for the next attempt.

For these standalone structured-table certificates, `source_findings` remains separate from `result`: PASS means supported source records and their warnings were preserved, not that the source itself is correct. This older CLI certificate contract remains limited to structured tables. General charts use the automatic application guard described above; users do not need this CLI.

## Supported table families

- Rate schedules: explicit adjacent min/max or 최저/최고 columns, then adjacent food/operations/total or 식품비/운영비/계 headers. At least two valid integer population bands are needed. Multiple school-level cost groups and named school exceptions are retained. Missing upper bounds remain unknown; rates are never added into expenditure. The source total is retained even when its components disagree.
- Organization blocks: 현원/정원 or headcount/quota rows, organization headers two rows above them, and explicit total/general/special/political job-type subheaders. Vertically merged single-count organization labels are supported. Original parent totals and departments remain separate records; their relationship is not inferred merely because numbers sum correctly. They are not ranked or summed across levels.
- Explicit grouping: adjacent item/value role headings plus actual XLSX row-outline metadata produce a descriptive source-row/level/value table. Parent and child values are not summed. Removing the required grouping caveat blocks the report.
- Discovery locates explicit schema anchors among cells with content, including beyond row300/column80; it ignores empty styled padding instead of scanning a fixed rectangle. Tables are paged in groups of8records. The schema grammar remains bounded and ambiguous hierarchies are not inferred. Original source labels remain in their original language; generated headings, notes and confirmation are English/Korean. Blank cells are not filled with zero. The app asks for confirmation before creating a structured report. Other sheets left out are counted.

The source/output checker independently verifies supported block coverage, row order, field roles, cell references, source hash, table headers, lossless numeric values and required source warnings. It supports General, integer/grouped and two-decimal formats plus the specific integer accounting formats covered by regression fixtures. Other formats fail verification rather than being silently accepted. Arbitrary business meaning, full Excel formula evaluation, all number-format rendering and Office layout are not proven.

## Components and guarantees

`layouts.py` implements the app's bounded parser. `verification/pipeline.py` connects report creation and certificates. `rawxlsx.py`, `oracle.py`, `output.py` and `check.py` independently read/check evidence without the production parsing/arithmetic functions. `formulas.py` checks bounded formula caches; `text_values.py` reconstructs supported numeric text placeholders. Reference arithmetic and production sums/means/changes/ratios use exact Fractions from decoded operands. Display formatting, decoding bridges, some statistical/selection paths and general business interpretation remain tested/trusted boundaries. The independent numeric-text/recipe checks deliberately share `numfmt` for final display, so arithmetic independence is not a second independent formatter.

`selftest` exhaustively checks growth over the finite integer domain−25..25, small closed interval boundaries, and all81assignments of four atoms to three groups. It also checks translated blocks, both languages, source warnings, duplicate/missing references, corrupted reports and overwrite refusal. Mutation tests change saved values, delete rows/warnings, and separately corrupt chart caches and embedded workbooks. The chart component has direct golden-fixture tests; generic source-linked chart certificates are not supported by the report CLI yet.

Tests and bounded exhaustive checks are separate from the [actual Dafny kernel proof](../verification/proofs/README.md). The app now calls the generated integer-sum, sequence-equality, tolerance-comparison and rational cross-product kernel. The rational transformation itself is proved relative to positive denominators; Fraction normalization/extraction and the tolerance budget remain tested/trusted boundaries. Dafny4.11.0 verified the contracts and termination, and implementation mutations were rejected. This is **not a universal proof of the app, spreadsheet interpretation or Office**. Compiler, solver, Python, integer/rational conversion, source recipes, selection and format rendering remain trust boundaries. Real company files and human executable trials remain separate evidence.

## Licensing

New original code follows the existing project MIT license. No dependencies, models, weights or third-party source were added. Existing dependency notices remain applicable. Public regression files and research PDFs are not bundled; use synthetic fixtures for redistribution.
