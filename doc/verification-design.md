**English** | [한국어](verification-design.ko.md)

# Verification design and algorithms

This guide explains the automatic checks in the source tree: why the extra code exists, what mathematics it uses and where its guarantees stop. For the app's states and special-case report notes, start with [the user guide](verification.md). For current unsupported layouts and evidence gaps, see [limits](limits.md).

The formula/text/exact-intermediate hardening described here is **unreleased development source**, not a feature or new validation claim for the published v0.3.0 executable. See [the current contract](reliability-hardening.md).

## Why check more than the calculated number?

A correct calculation can still become an incorrect report: a row can disappear, a chart can retain an old cache, “Edit Data” can contain different numbers, or a workbook can change between preview and saving. Recalculating with the same production function would repeat some of its defects. The guard therefore separates four questions:

1. Does the supported source recipe produce the expected values when evaluated by a separate raw-XLSX reader?
2. Do the used saved formulas and their dependencies agree with the bounded independent formula evaluator?
3. Do the serialized text/table/chart and embedded data match the checked values?
4. Are those still the source bytes and output bytes being published?

This separation helps detect implementation errors. It does not independently decide whether the selected cells answer a business question: source recipes and schema meaning remain part of the trusted specification.

## Automatic path

```text
app preview/generation -> core.run -> fill text/table/chart
  -> capture resolved values, source recipes and paragraph/shape identity
  -> serialize preview -> Session.verify
       -> rawxlsx + recipes/oracle: supported source reconstruction
       -> formulas/text_values: bounded formula caches + numeric text
       -> text/table XML or charts: actual output + embedded XLSX
  -> passed / partial / blocked
save -> same preview source snapshot -> staged output -> verify again
     -> byte/hash checks + source fingerprints -> new output only
```

`core._Run` captures filled text and marked tables/charts, including structured reports and Word-linked means. `Session.verify` accumulates failures and unsupported coverage; any known failure wins over partial coverage. `Session.publish` checks the staged file before publishing without overwriting an existing output. Both the UI and the core save path refuse known failures and partial coverage; ordinary confirmation cannot override either state.

Text-only filling uses the guard too. Changed paragraph content/order is checked after serialization, direct/latest/rank values are source-linked, and `prev/week/avg/peak/low` results are independently reconstructed. Prose/model inputs are checked, but every semantic claim in their wording is not independently proved. Static unmarked template content and arbitrary Office layout remain outside this numeric contract. Other existing warnings and sentence checks still apply. Normal use creates no proof/certificate attachment in the report.

## Mathematical algorithms and their purpose

### Exact decimal addition

The small proved kernel operates on integers. For decimal inputs `x_i = c_i * 10^e_i`, choose `k = max(0, -min(e_i))` and encode `n_i = c_i * 10^(e_i+k)`. Every `n_i` is an integer. The decimal result is reconstructed from `Sum(n_i)` with exponent `-k`, without converting through binary float.

Example: `0.10 + 0.20` becomes `(10 + 20) * 10^-2 = 0.30`. `kernel.sum_decimals` implements the conversion for two finite decimals; `layouts._cost_sum` uses it to compare food plus operations with the source total. A mismatch is reported while the original total is retained. Rate rows are not summed into expenditure.

The Dafny `Sum` contract proves sum correctness and termination for every finite integer sequence. Decimal encoding/decoding is tested Python code, not part of that formal proof. Explicit production sums and means also use rational intermediates and the generated sum; that does not prove every statistic or the choice of which rows belong in an aggregate.

### Exact sequence equality

`Same(xs, ys)` returns true exactly when integer sequences have the same length, order and elements. Table comparison encodes row count, column counts, text lengths and Unicode code points. Those lengths distinguish, for example, cells `['ab', 'c']` from `['a', 'bc']`. Supported line-break representations are normalized before encoding. Chart names/categories use code-point equality.

This checks missing/extra rows, order and text as well as values; equal totals alone would miss some of these errors. The integer equality algorithm is proved. XML extraction, normalization and the encoding bridge are tested/trusted code.

### Rational tolerance comparison

Source reference calculations use `Fraction` to avoid adding reference rounding error. For `a=p/q`, `b=r/s`, and tolerance `epsilon=t/u`, with positive denominators and `t >= 0`, the comparison

```text
abs(a-b) <= epsilon
```

is equivalent to

```text
abs(p*s*u - r*q*u) <= t*q*s.
```

`kernel.equal_rationals` extracts Fraction numerators/positive denominators and calls the generated `RationalWithin` method. Its formal contract proves the cross-product integer computation equivalent to the real-rational inequality above. The real-valued scale is ghost proof code and does not enter the generated app code. Fraction normalization/extraction remains trusted; the proof does not establish that the chosen tolerance is appropriate.

`recipes.near` requires exact equality when both values are integers; other direct/serialized comparisons use two ULPs. Explicit grouped aggregates keep exact rational intermediates through sum/product/mean, formatting and any required container conversion. Their independent audit reconstructs supported decimal-decoding bridges, uses generated `SumBound` to bound the sum of supplied row-decoding errors, and requires the exact rational result or the supported legacy single-conversion representation. A captured rational must match exactly. There is no operation-count ULP multiplier. Chart integers must round-trip a float exactly, and explicit supported formats must retain displayed precision. These conversion/format policies and other statistical paths are tested/trusted; this is not an IEEE hardware theorem. See [the formulas and trust boundaries](binary-verification.md#exact-aggregates-and-source-decoding-error).

### Exact changes and bounded saved-formula checks

Production sums, means, differences and ratios preserve Fractions made from decoded decimal spellings. Relative change is `(now-base)/abs(base)` and has no answer for a zero base. A percentage-point difference uses the declared scale. `text_values.py` independently reconstructs numeric comparison/extreme placeholders from declared source history and checks their formatted results; `automatic._document` independently checks Word-linked period windows and means. Source/date/measure selection still specifies the question rather than proving it is the right business question.

`formulas.py` separately parses a declared Excel grammar and evaluates supported expressions/dependencies using rational arithmetic. It checks used formulas, including formulas in a latest/rank selection domain, and distinguishes stale caches from unsupported expressions. Its grammar, coercion/precision policy and resource budgets are [explicit](reliability-hardening.md#formula-grammar-and-limits). It does not invoke Excel, alter source files or prove equivalence to all Excel floating-point behavior. Unknown cases prevent saving rather than silently passing.

### Independent calculations, coverage and integrity

`rawxlsx` reads ZIP/XML records without importing the app's workbook/calculation functions. `recipes` reconstructs supported direct references, dates, grouped aggregates, ranks and shares using declared cell references. `oracle` independently recognizes the bounded structured schemas and checks record roles, source order, page coverage and mandatory warning contents. `charts` compares series, point counts/indices, gaps, caches, relationships and every referenced embedded-workbook cell against the captured grid.

The development certificate checker also uses exact `Fraction`/`Decimal` reference arithmetic, closed-interval lookup and explicit disjoint atom scopes. `arithmetic.partition_sum` rejects duplicate/overlapping scopes and incompatible measures, units or periods. It does not infer real organization membership from coincidentally matching totals; that broader grouping theorem is outside the proved kernel. These checks are ordinary tested code, not additional Dafny proofs.

SHA256 fingerprints detect input/output byte changes between checks. They are integrity checks, not signatures proving an author's identity or a data source's truth. ZIP/XML parsing, the filesystem and the remaining interval between the last check and publication are trusted boundaries.

Decimal/rational display calls the proved half-up integer decision through `kernel.round_decimal` / `round_rational`; actual Excel display was compared separately. The independent recipe/text checks share the production `numfmt` display bridge. They therefore supply independent numeric reconstruction, not a second independent number-format implementation. See [binary reader, grouping and rounding details](binary-verification.md).

## Code map

| Path | Responsibility |
|---|---|
| [app.py](../app.py) | Translated verification status, preview source reuse and blocked-save UI. |
| [core.py](../core.py) | `ExcelSource.audit_evidence`, `core._Run.fill_marked_shape`, `run`: recipes, capture and mandatory guard/save boundary. |
| [suggest.py](../suggest.py), [layouts.py](../layouts.py) | Report selection/layouts; bounded structured records, warnings and decimal cost sum. |
| [automatic.py](../verification/automatic.py) | `Session.capture`, `verify`, `publish`: in-process coordination and final publication checks. |
| [sources.py](../verification/sources.py), [rawxlsx.py](../verification/rawxlsx.py), [recipes.py](../verification/recipes.py) | Independent raw reader and declared source calculations/comparison policy. |
| [formulas.py](../verification/formulas.py), [text_values.py](../verification/text_values.py) | Bounded saved-formula/dependency audit and independent numeric text calculations. |
| [oracle.py](../verification/oracle.py), [charts.py](../verification/charts.py) | Structured meaning/coverage/warnings and saved chart/embedded-data checks. |
| [binary.py](../verification/binary.py), [compound.py](../verification/compound.py), [htmlsource.py](../verification/htmlsource.py), [numfmt.py](../numfmt.py) | Independent binary/HTML evidence and proved display rounding. |
| [kernel.py](../verification/kernel.py), [verified_kernel.py](../verification/verified_kernel.py) | Typed conversions and actual generated integer code used by the app. |
| [Kernel.dfy](../verification/proofs/Kernel.dfy), [build.py](../verification/proofs/build.py), [receipt.json](../verification/proofs/receipt.json) | Proof contracts, pinned generation and source/generated byte identity. |
| [pipeline.py](../verification/pipeline.py), [check.py](../verification/check.py), [output.py](../verification/output.py), [arithmetic.py](../verification/arithmetic.py) | Developer structured-table certificates and reference/output checks. This CLI's contract does not include generic chart certificates. |
| [test_coverage.py](../verification/test_coverage.py), [test_automatic.py](../verification/test_automatic.py), [test_verification.py](../verification/test_verification.py), [test_formulas.py](../verification/test_formulas.py), [test_reliability.py](../verification/test_reliability.py) | Corruption, partial/error precedence, source drift, save behavior, formulas, precision, text, structured fixtures and finite exhaustive checks. |

## Evidence and changes

The binary/grouping/rounding extension checked on 2026-11-08 passed 675 existing tests, 69 verification tests, 116 GUI checks and nine public-file regressions covering 29 bindings. Four actual Excel-generated XLS/XLSB files passed 12 chart/date/table paths, and two bilingual grouping flows passed confirmation, cancellation and saving. Native PowerPoint 16.0 rendered all 14 final new-path reports with no out-of-bounds shapes. Real Excel display comparison passed five workbook checks and 126 format cases with zero differences. These are snapshot-specific observations, not an unseen-company accuracy score or evidence of a rebuilt executable. See [remaining evidence](limits.md).

The expanded formal evidence is Dafny 4.11.0 reporting nine verification units and zero errors for seven source definitions: the recursive sum specification and six executable contracts (sum, sequence equality, bounded difference, rational cross multiplication, signed half-up rounding and supplied-error accumulation). Three new faulty cross-product implementations with unchanged contracts were rejected; the earlier three integer-kernel mutations remain separate evidence. Proof/compiler/runtime/conversion assumptions and reproduction commands are in [the proof guide](../verification/proofs/README.md).

When extending a source recipe or supported layout, update its declared contract and independent expectation, then add a valid example and a corruption or unsupported-case check. For a behavioral change run `test_core.py` and `python -m verification selftest` from the project environment, and inspect affected UI/Office output where relevant. Preserve failed evidence. A new proof-kernel change also requires regeneration, proof logs, matching hashes/receipt and bridge checks; never hand-edit the generated Python file. Ordinary users perform none of these developer steps.

Original documentation follows the project's existing MIT license. No algorithm, model or third-party material was imported for this guide; existing component notices still apply.
