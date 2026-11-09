**English** | [한국어](limits.ko.md)

# Supported layouts and remaining verification

Automatic selection is a suggestion. Check the sheet, measure, period, conditions and unit shown before generating. An absent warning does not prove that the chosen figure matches your business question.

v0.4.0 checks supported formulas, numeric text and precision, and refuses saving on a mismatch or incomplete coverage. The Windows and Office evidence below applies to v0.3.0. [v0.4.0 verification scope](reliability-hardening.md).

## Supported boundaries

- Explicit population-band cost schedules and repeated organization/job-type blocks produce source-linked descriptive tables, with confirmation, unknown bounds and source-total warnings retained. They do not sum rates into spending or mix hierarchy levels. See [CPU verification and exact boundaries](verification.md).

- Months across columns work when the period row and measure rows are identifiable. References such as `latest:Monthly!1,2` select them explicitly.
- Up to three header levels under periods work when every column has a label, including actual merged labels, every combination appears once per period, and all periods repeat the same combinations. The selected combination is explicit and requires confirmation. Missing or duplicate combinations still cause automatic refusal.
- Numbered financial outlines compare the top level, excluding child rows. Income and expenses are separate concepts and are not added into one total. Unnumbered or ambiguous hierarchies and some pivot layouts remain unsupported.
- Numeric chart cells and text numbers use the same parsing rules. Bad text remains a warned gap. Tests check point counts as well as values, including the embedded workbook.
- Unknown or conflicting fiscal-year ends are not replaced with June. Raw period labels and values remain available, but completion and calculated period comparisons are withheld. State the end in the workbook to continue.
- Directly sourced current values in dated calculated sentences carry their number-format colour. Derived sums, means, ratios, historical comparisons and model text keep the template's colour. Black and white also retain the template colour.

- Explicit structured headers are discovered among actual cells beyond the former300row/80column probe. The grammar still requires stated min/max or organization/job roles; an arbitrary hierarchy is not inferred.
- CSV/TSV/TXT, supported ODS values and renamed XLSX exports now have independent source reconstruction. Supported BIFF2–5/8 XLS and XLSB cells now have independent decoding; unsupported binary records/workspace variants and unsupported ODS duration grammar/expression cases remain partial; text/ODS readers enforce a two-million-cell expansion budget.
- Used formula caches require the bounded independent formula audit. XLS/XLSB formula expressions, arrays, external links, unknown functions/coercions and resource-limit cases are partial and cannot be saved. The XLSX/ODS subset and exact-decimal decoding boundary are stated in the [reliability contract](reliability-hardening.md).
- Text-only fills use source/serialization checks too. Supported numeric comparison/extreme placeholders and Word-linked means have independent calculations; not every prose or model claim is semantically proved. Chart integers and explicit display precision must survive the numeric container or the chart is refused.
- Bubble templates accept explicit X plus Y/size pairs, with nonnegative sizes and at least one complete point per series. Size caches and embedded data are checked. Ambiguous columns and an empty series are explicit input errors.
- The small-window summary occupies its own row. The proved kernel also covers rational cross multiplication; the floating-error budget, Fraction normalization and general business selection remain outside that proof.

Explicit XLSX row-outline levels can be preserved as descriptive tables without hierarchy sums. Decimal/rational display rounding uses an actually proved half-up decision; grouped aggregates retain exact rational intermediates and use a proved sum of supplied decoding-error bounds. Production means/changes/ratios also retain rational intermediates. Source conversions, the shared display formatter and other statistical/selection paths remain trusted/tested. An equal-row mean of rates is not necessarily the intended total rate; weights and denominators still need an explicit definition. See [the numeric boundaries](binary-verification.md).

## Evidence and current test scope

The eighth measurement used eight previously unseen public files on frozen v0.2.0 code. It generated six reports and refused two complex layouts, and found an empty chart, missing years/units and excess sentence filtering. Once used to fix these defects, those inputs are regression data; rerunning them is not a new fair accuracy score.

Those two refused layouts now have bounded source-table support and independent regression checks. The app distinguishes passed, partial and blocked for tables and category/date/XY charts. v0.4.0 extends the guard to numeric text and supported formulas and refuses saving for partial coverage. The app uses a small Dafny-proved kernel for sum, equality and tolerance. [Exact boundaries](verification.md).

Existing sample use and automated checks are sufficient for the current scope by project decision. A separate human trial remains excluded. The later clean-Windows request was completed with the exact v0.3.0 EXE: 19 cases passed in a fresh offline Windows 10 Pro 19045.3803 VM with no Office or separate Python. Native PowerPoint 16.0.20326.20232 rendered 18 host and 18 actual guest reports (44 slides total); the guest's 22 slide images matched the inspected host images byte for byte. This does not establish Windows 11, another physical PC, private-company compatibility or an independent human trial.

No dependencies were added for these changes. The project remains MIT; component notices and source-data licenses remain separate and unchanged.

See [the environment evidence and earlier test decisions](testing-alternatives.md) and [development history](development-history.md). Earlier skipped checks are historical decisions, not the status of the completed Windows 10 VM test. No alternative office suite or host security/virtualization feature change was required.

## Optional future reproduction

This is a reference only, not a required current trial. If a concrete issue is reported, use an approved local copy of its workbook and template. Write down the expected sheet, period, unit and five source values before opening SlideFill. Start the executable, choose the files, press F5 and inspect those values. Use Ctrl+R to change the selected measure, then Ctrl+G to save a new report. Open it in PowerPoint and compare the text, table, chart and chart's Edit Data with the original. Cancel one doubtful selection and try one unreadable file to check recovery. Record the executable hash, Windows/Office version, result and discrepancies. Keep company files and screenshots private. This procedure needs an actual person; scripted clicks are not a completed human trial.
