**English** | [한국어](README.ko.md)

# Excel-hosted VBA comparison spike

This experiment keeps SlideFill's existing Python/Fraction calculations and compares them with one original VBA module hosted by installed desktop Excel. It does not change the normal app calculation backend. [SlideFillOracle.bas](SlideFillOracle.bas) has no file/network access or automatic workbook events.

The app's optional **Excel VBA Check** now uses the same `SFCalc` module for supported report aggregates. The bundled [engine workbook](../../assets/slidefill-vba.xlsm) contains two empty sheets and the exact VBA project from the native experiment. To reproduce it, run `python tools/vba/make_runtime_engine.py TESTED.xlsm NEW.xlsm` with the tested single-module workbook as input. The generator verifies the VBA project hash, writes a fresh workbook with no copied worksheet data, and refuses to overwrite a file.

The intended comparison has three paths:

1. Excel worksheet formulas call the VBA `SFCalc` UDF over declared cells.
2. The comparison runner calls the same VBA function through Excel `Application.Run` and receives its result.
3. Existing Python arithmetic computes the declared operations independently.

The first two paths share the VBA implementation and Excel runtime. Agreement confirms integration/transfer consistency; it is not independent proof of the algorithm. Python and fixed rational expectations help expose shared numeric assumptions. `Evaluate` or `WorksheetFunction` calls alone do not demonstrate that this VBA module executed.

## Prerequisites and first import

Use Windows desktop Excel with VBA installed and an interactive user session. The app does not bundle Microsoft Excel/VBA binaries. The supplied macro-enabled scaffold and original module must be allowed by the existing Office policy. This experiment does not change Trust Center, trusted locations, signing keys or registry settings.

Programmatic access to the VBA project was unavailable in the tested environment. The executable preparation/comparison tooling can be reviewed before this one manual step:

1. Open the generated **test scaffold** in Excel; preserve the original input workbook.
2. Press **Alt+F11** to open Visual Basic Editor. In Project Explorer, select **VBAProject (your test scaffold's filename)**. Choose **File → Import File…** and select this folder's `SlideFillOracle.bas`. Confirm that **Modules → SlideFillOracle** appears under that same project, then return to Excel. `Sheet1`, `Sheet2` and `ThisWorkbook` alone are the default objects, not the imported module.
3. Save the scaffold as its existing macro-enabled `.xlsm` format. Do not paste the module's `Attribute` line into an ordinary code window; File Import handles it.
4. If Office policy blocks this original test macro, record the refusal. Do not enable all macros or change project-access settings to force the test.
5. Run the supplied comparison workflow against the reviewed artifact. Until import and actual execution occur, both VBA routes are **NOT RUN**. Static source checks, workbook formulas showing `#NAME?`, or Excel worksheet-function calls are not a VBA execution pass.

Import `SlideFillOracle` once per test workbook. If a second import creates `SlideFillOracle1`, remove the duplicate from the test workbook. Removing it resolved the ambiguous `SFCalc` reference in this run.

For example, a normal English Excel formula `=SFCalc("mean",B2:B4)` calls the module on three cells. Regional Excel interfaces may use a different formula separator. A runner can call the verified artifact using a fully qualified name such as `'Engine.xlsm'!SlideFillOracle.SFCalc`; it must bind the actual artifact and input identities rather than relying on whichever workbook is active.

## Commands

Run from the repository root with the existing project virtual environment. Use a **new** run directory; keep earlier artifacts and receipts. This example uses `run-02` so an existing `run-01` is preserved.

```powershell
$spikeRun = '.cache/agent-work/vba-spike/run-02'
.venv/Scripts/python.exe tools/vba/compare.py prepare $spikeRun
powershell -NoProfile -File tools/vba/host.ps1 -Mode Build -InputWorkbook "$spikeRun/inputs.xlsx" -ModulePath tools/vba/SlideFillOracle.bas -EngineWorkbook "$spikeRun/engine.xlsm" -OutputJson "$spikeRun/build.json"
```

Read `build.json`. If project access prevented automatic module import, use the normal VBE File Import procedure above on that run's `engine.xlsm`, then save and close it. A blocked Build result is not a VBA execution result. Do not change project-access or macro-security settings to make Build pass.

```powershell
powershell -NoProfile -File tools/vba/host.ps1 -Mode Run -EngineWorkbook "$spikeRun/engine.xlsm" -AppInputJson "$spikeRun/app-input.json" -OutputJson "$spikeRun/native-results.json"
.venv/Scripts/python.exe tools/vba/compare.py summarize "$spikeRun/app-input.json" "$spikeRun/native-results.json" "$spikeRun/comparison.json"
```

The prepared manifest has21 cases. Inspect `native-results.json` and `comparison.json`; count the actually executed cases and distinguish MATCH, DIFFERENT, blocked/unavailable and errors. Expected binary/rounding differences must be retained as differences under the stated comparison contract. Keep source, module, workbook and input hashes with the result. If the macro was never imported/allowed/executed, preserve **NOT RUN** rather than inferring success from the workbook opening.

Run checks the original literal workbook against the actual engine cells by type and Double bits before calling VBA; declared metadata hashes alone are insufficient. Build preserves numeric-looking text and format strings explicitly. An earlier scaffold copied through Excel `Value2` could coerce these strings into numbers. After importing/saving its module, use `repair_fixture.py` to create a new typed-data copy without touching the original or changing its embedded VBA binary:

```powershell
.venv/Scripts/python.exe tools/vba/repair_fixture.py "$spikeRun/inputs.xlsx" "$spikeRun/engine-final.xlsm" "$spikeRun/engine-typed.xlsm"
```

Then use that new copy as `-EngineWorkbook`. The repair refuses a missing VBA project or existing output. Its package tests use non-executable sentinel bytes and do not establish VBA compilation or execution.

## Observed results

Excel 16.0 (build 20326) verified the operand types and numeric Double bits across the original workbook, engine copy and app input for all 21 cases. The worksheet UDF and `Application.Run` matched in **21 of 21 cases**, with identical Double bits in 16 cases and matching error codes in five. Against Python's nearest Double or expected error code, VBA matched **17 of 21**.

| Case | Both VBA routes | Python exact |
|---|---:|---:|
| `decimal_sum` — Double sum of `0.1 + 0.2` | `0.30000000000000004` | `3/10` |
| `cancellation` — `10^16 + 1 − 10^16` | `0` | `1` |
| `tiny_growth` — growth from `100` to `100.00000001` | `9.9999937219763522e-11` | `1/10000000000` |
| `banker_tie_probe` — round `2.5` to an integer | `2` | `3` (half away from zero) |

In `decimal_variant`, the bounded `CDec` sum of `0.1 + 0.2` returned the nearest Double to `3/10`. In `large_literal`, both VBA routes returned `9007199254740992`, one less than the exact input `9007199254740993`. That case matched at the Double boundary. Excel's `Range.Text` reported `9007199254740990` for the declared display format.

The 21 synthetic cases exercise the developer comparison runner. SlideFill uses Python/Fraction for report generation. Local evidence is in `native-results-after-dedup.json` and `comparison-after-dedup.json` under `.cache/agent-work/vba-spike/run-02/`.

## Fixed operation contract

`comparison.json` reports VBA route agreement, equality at the declared Double output boundary, and `native_binary_minus_python_exact` separately. A large integer can lose a unit on Double conversion even when the two VBA routes agree; the exact difference remains visible. The comparison is diagnostic and does not silently replace Python's result.

`SFCalc(operation, values, [digits])` accepts a contiguous Excel Range, a two-dimensional Variant array, or one scalar. Arrays/ranges are limited to **100,000 cells** and read in row-major order. Range inputs use `Value2`. Genuine empty cells are ignored; text (including empty strings), Booleans, dates supplied as typed Date values, Excel errors, objects and nonfinite values are refused. A formatted date serial from `Value2` is still a number, so the caller must select the intended value cells explicitly.

| Operation | Result and preconditions |
|---|---|
| `sum` | Double sum of retained numeric cells; empty input0. |
| `count` | Number of retained numeric cells as Double; empty input0. |
| `mean` | Double sum/count; empty input returns divide-by-zero status. |
| `growth` | Exactly two retained numeric values, **now then base**; `(now-base)/Abs(base)`; zero base unavailable. |
| `points` | Exactly two retained values, **now then base**; absolute difference `now-base`, not a relative percentage. |
| `round` | One retained numeric value; integer digits−15..15; `WorksheetFunction.Round`. |
| `vba_round_probe` | One value; integer digits0..15; deliberate VBA `Round` diagnostic, whose banker ties differ from worksheet rounding. |
| `decimal_sum` | Bounded `CDec` accumulation, converted to Double at return. Diagnostic only; not arbitrary exact rational arithmetic. |

Error results are actual VBA `CVErr` values: **2015 / `#VALUE!`** for unsupported operation, shape, input type, digits or arity; **2007 / `#DIV/0!`** for empty mean/zero growth base; **2036 / `#NUM!`** for arithmetic, Decimal conversion/overflow or nonfinite failure. A COM bridge may expose these as error variants/HRESULT-backed wrappers; preserve their type/status rather than interpreting an error integer as a successful number.

Use separate comparisons for contributing cells, status, raw numeric value, declared formatted string and final report transfer. Excel's formatted `Range.Text` can vary with column width. Do not change precision-as-displayed, source values or formatting to manufacture agreement.

## Limits and evidence

This spike covers only the listed operations over explicit inputs. It does not reimplement automatic source selection, every formula, dates/calendars, arbitrary workbook macros or report rendering. A Double loop or bounded Decimal intermediate can lose information or overflow where Python/Fraction remains exact. Numeric cells already limited by Excel storage cannot regain missing digits in VBA. Cancellation, rounding ties, negative/zero bases, blank/error inputs and tiny changes belong in the comparison fixture.

Production [Python arithmetic](../../facts.py), [core](../../core.py) and numeric guards supply the report's values. The app check compares independently verified table and chart outputs for supported single-column sums, means and counts. Numeric text, dates, formulas and filtered or multi-column aggregates are outside this first comparison scope.

## Primary sources and licensing

Microsoft documents [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run), [VBA Round](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function), [Double](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/double-data-type), [Decimal](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/decimal-data-type), [Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2), [Text](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text) and [macro security](https://support.microsoft.com/en-us/excel/change-macro-security-settings-in-excel). [VBA is integrated with Office](https://learn.microsoft.com/en-us/visualstudio/vsto/vba-and-office-solutions-in-visual-studio-compared?view=visualstudio); bundling a source module does not establish a standalone runtime or redistribution entitlement.

The original module, its compiled VBA project and the tooling follow the existing project [MIT license](../../LICENSE). Installed Office/VBA remains subject to the host's applicable license and policy.
