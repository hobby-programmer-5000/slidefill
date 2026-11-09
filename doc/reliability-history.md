**English** | [한국어](reliability-history.ko.md)

# From client-deck feedback to Excel VBA verification

This history follows the SlideFill work on 9 November 2026, Korea Standard Time. The project author set the direction, proposed embedding a VBA module and calling it from both Excel and the app, and handled the workbook-specific approvals in Excel. Codex helped research, implement, and test the changes. [Development and verification history](development-history.md)

## A polished slide can hide a wrong number

Feedback on the project introduction asked whether anyone should trust an automatically filled client deck without human review. The project author put the review boundary in a prominent README warning in commit `13dfc27`. A follow-up pointed out that a polished slide can make a wrong number easier to send. The author then asked to strengthen the implementation so people could spend less time rechecking the arithmetic. The question became which steps the app could verify and which still needed the report author's judgment.

The first implementation pass separated source selection, calculation, and saved PowerPoint checks. It retained exact rational intermediates for supported operands, independently evaluated supported formula caches and their dependencies, and withheld output if a mismatch or incomplete independent coverage was detected. The bounded Dafny kernel covered stated integer and rounding contracts. Checks on source decoding, chart data, and saved slides used separate tests. The algorithm guides then distinguished exact operations from decisions such as choosing the right sheet, denominator, unit, or period. [Numeric reliability](reliability-hardening.md), [Feature algorithms](feature-algorithms.md), [Verification design](verification-design.md)

## The paper search and the tests it informed

At the project author's request, the team used Liner Scholar for six search batches on spreadsheet headers and provenance, test oracles and metamorphic testing, contracts, and numeric deviations between implementations. Two fetches returned relevant body excerpts. Three other works informed the discussion at abstract level. Liner also listed full text for *jContractor* and *Metamorphic testing and beyond*, but both body fetches failed. The accounts below describe the evidence consulted during development. The links lead to primary copies or author repository records checked for this history. [Research methods and evidence levels](research-methods.md)

- **Robin Abraham and Martin Erwig, 2004.** The body excerpt showed how spatial header inference can support unit checks. We used that idea to keep a comparison's kind, period, and exact value explicit at module boundaries. SlideFill's supported selections still require the correct business measure and units. [Header and Unit Inference for Spreadsheets Through Spatial Analyses](https://web.engr.oregonstate.edu/~erwig/papers/HeaderInf_VLHCC04.pdf)
- **Huai Liu and colleagues, 2014 journal issue.** The body excerpt treated relations between executions as checks when expected results are difficult to specify for each case. We tested equivalent numeric spellings, positive rescaling, and preserved change direction within stated preconditions, then added concrete counterexamples. [How Effectively Does Metamorphic Testing Alleviate the Oracle Problem?](https://nottingham-repository.worktribe.com/output/720879/how-effectively-does-metamorphic-testing-alleviate-the-oracle-problem)
- **Earl T. Barr and colleagues, 2015.** The abstract describes specifications, contracts, metamorphic tests, and human domain knowledge as sources of expected behavior. We kept the exact mathematical reference, independent saved-output checks, and Office observations separate so their agreement could be interpreted within scope. [The Oracle Problem in Software Testing: A Survey](https://discovery.ucl.ac.uk/id/eprint/1471263/)
- **M. Moein Almasi and colleagues, 2018.** The abstract studied deviations after moving spreadsheet calculations to Java in two financial applications. It informed our choice to target adjacent large integers, long decimals, very small changes, and cancellation in our own synthetic cases. [Search-Based Detection of Deviation Failures in the Migration of Legacy Spreadsheet Applications](https://philmcminn.com/publications/almasi2018.pdf)
- **Tsong Yueh Chen and colleagues, 2018.** The abstract surveyed test generation and verification through necessary relations and the challenges of choosing those relations. We documented when each proposed relation should hold and tested both false acceptance and false rejection. [Metamorphic Testing: A Review of Challenges and Opportunities](https://hub.hku.hk/handle/10722/254895)

The resulting change preserved numeric token identity in optional wording, sent exact relative changes to the cross-file comparisons, and preserved missing gaps and the signed zero baseline in the contents canvas. These were SlideFill engineering decisions, validated with targeted regressions and saved artifacts. The papers studied their own systems and methods. [Information-preservation contracts](reliability-hardening.md#preserving-information-between-calculations-and-wording)

## Trying the actual Excel and PowerPoint renderers

The project author asked to open real Excel and use the real contents canvas for visual comparison. Excel 16.0 build 20326 recalculated six `AVERAGE` formulas and exported three chart images. Native PowerPoint exported six slides from the English and Korean reports. The inspected outputs showed the intended means, signed bars from zero, and gaps where data was missing. The normal app loaded the synthetic workbook and its six exact values into the real Tk contents canvas. Desktop capture timed out and the accessibility tree was insufficient in this run, leaving full window appearance and click-flow inspection unverified at this stage. The exported Office images document the selected chart and slide rendering cases. [Environment test scope](testing-alternatives.md)

## The project author's VBA proposal

The project author saw a possible source of disagreement in making different environments reproduce each other's behavior. They proposed implementing the same operations in an embedded VBA module, comparing a run inside Excel with a run called by the app, and keeping the existing Python code, formulas, and mathematical algorithms. That proposal led to an experiment with three paths.

1. The worksheet calls the original `SFCalc` module as a VBA user-defined function.
2. The app-side runner calls the same module through Excel `Application.Run`.
3. Python computes the declared operations using the existing exact arithmetic and rounding rules.

The two VBA routes share an implementation and the Excel runtime, so their agreement tests how inputs and results cross the app boundary. The Python result is a separate mathematical reference under the project's declared rules. [Original VBA source](../tools/vba/SlideFillOracle.bas), [Excel Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run)

## What went wrong before the test ran

Excel refused programmatic access to the VBA project. The initial blocked run verified 21 input cases and executed zero VBA cases. The author used the editor's normal file import, and a duplicate module caused `Ambiguous name detected: SFCalc`. The author removed the duplicate and saved a single-module workbook. Another setup check found that bulk `Value2` copying had converted numeric-looking text into numbers. The repaired fixture preserved the original workbook and VBA binary and checked the actual operand types and numeric Double bits before execution. The work used workbook-specific approval under the existing Office policy. [VBA setup and reproduction](../tools/vba/README.md), [Office developer security notes](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/security-notes-for-microsoft-office-solution-developers)

## The experiment results

After repair, desktop Excel 16.0 build 20326 executed **21 of 21** cases by both routes. The worksheet and the app-side call agreed in all 21 cases, with identical Double bits in 16 numeric results and matching error codes in five. VBA and Python matched in **17 of 21** at the declared nearest-Double or expected-error boundary. The remaining four cases were recorded as differences. [Full comparison and operation contract](../tools/vba/README.md#observed-results)

| Synthetic case | Result from both VBA routes | Python exact result |
|---|---:|---:|
| `0.1 + 0.2` using Double accumulation | `0.30000000000000004` | `3/10` |
| `10^16 + 1 − 10^16` | `0` | `1` |
| Growth from `100` to `100.00000001` | `9.9999937219763522e-11` | `1/10000000000` |
| Round `2.5` using the VBA `Round` probe | `2` | `3` with half-away-from-zero rounding |

The large-integer case illustrates the difference between Double agreement and exact agreement. Both VBA routes returned `9007199254740992` from an exact input of `9007199254740993`. This agreed at the Double boundary and lost one unit at the exact boundary. The rounding row uses the explicit VBA `Round` diagnostic, which follows ties-to-even. [Comparison limits](../tools/vba/README.md#fixed-operation-contract), [Microsoft VBA Round reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function)

## How the app uses the result

The 21-case run led to the optional **Excel VBA Check**. SlideFill generates reports with its Python arithmetic. The check separately compares supported single-column sums, means, and counts in verified report tables and charts. It distinguishes exact equality, equality after conversion to Double, and a difference. It opens only the app's hash-verified macro workbook read-only in a separate instance of installed desktop Excel, following the existing macro policy. The app bundles its own VBA module. Excel supplies the runtime on the host. [Usage](../README.md#compare-with-excel-vba), [Excel AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)

The clean workbook extracted from a later EXE was initially refused under Office policy, both in an isolated copy and beside the EXE. The project author enabled content for that specific workbook and closed it without saving. The packaged bridge and workbook then returned all seven expected native results in a read-only run with matching file hashes. This covered sum, mean, count, the previous seven values' mean, and three growth comparisons. It used the packaged bridge directly because desktop capture failed in that run. A separate source GUI test had exercised the check window. [VBA comparison procedure](../tools/vba/README.md)

## Reviewing the figures and wording

The project author asked for a guided sequence and more accurate insight suggestions. The help now follows the path from input selection and preview to optional verification, insight review, and export. The insight checker checks the comparison period and direction and shows the evidence used for each candidate. In a bilingual replay, the updated checks rejected eight targeted wrong statements and retained four valid ones. Six Korean drafts from a separate packaged-model run agreed with the corresponding VBA comparisons on the synthetic data. With a current value of `99`, those comparisons were `+10.0%` against `90`, `−10.0%` against `110`, and `−2.0%` against the previous seven values' mean of `101`. This was evidence about those numbers, directions, and comparison bases. [Usage workflow](../README.md#usage), [Verification boundaries](verification.md)

Reviewing **Suggest references** also exposed two selection bugs. Choosing **all rows** could silently restore a default domestic-flight filter, yielding `10` for Gimpo where the all-row total was `160`. A different fixture started with three filters and displayed two. Recalculation lost the patient filter and changed `30` to `330`. The fix separates an omitted filter choice from an explicit empty selection, shows every applied filter with scrolling as needed, and clears stale actions while a condition is incomplete. Korean and English GUI checks covered both keeping and clearing the third filter, a five-filter check covered scrolling, and the explicitly selected `330` was verified in a saved PowerPoint. [Core regressions](../test_core.py), [GUI regression harness](../tools/gui_check.py)

These experiments cover the stated fixtures, calculations, and tested environments. A client-deck review still includes the chosen measure, periods, source freshness, units, wording, and final layout. The original concern is easiest to test when those decisions are visible alongside the figures. [Supported layouts and limits](limits.md)
