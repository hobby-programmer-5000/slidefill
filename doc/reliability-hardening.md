**English** | [한국어](reliability-hardening.ko.md)

# Numeric reliability: development source

These changes are **unreleased**. They describe this development source tree, not the published v0.3.0 executable. Earlier Windows/Office checks in the [development history](development-history.md) apply to their recorded artifacts.

The objective is to reduce manual number checking by checking the source, calculation and saved PowerPoint separately. It cannot establish whether a source is true or a selected business measure is the right one.

## What is checked

| Layer | Enforced check | Remaining boundary |
|---|---|---|
| Source | Independent readers reconstruct declared cells, filters, dates and labels; fingerprints detect a changed workbook or Word document. | The declared recipe specifies the intended cells. The checker does not independently choose the business scope. |
| Saved formulas | A separate bounded evaluator checks used formula caches and dependencies against exact arithmetic. Formulas in the declared latest/rank selection domain are included, even when their rows were not selected. | It is not a full Excel engine and does not rewrite/recalculate the workbook. Unsupported expressions remain partial. |
| Arithmetic | Sums, products, means, changes and ratios retain exact rational intermediates from decoded operands; supported display rounding uses the proved integer rounding decision. | The original file's decoding, the chosen algorithm and display formatting remain tested code. Exact arithmetic does not infer weights, units or business meaning. |
| Text | Changed paragraphs are checked after serialization. Direct/latest/rank values are compared with their source evidence. Numeric `prev`, `week`, `avg`, `peak` and `low` placeholders are independently reconstructed, including comparison arithmetic and latest-date tie selection. Calculated/model prose carries its input evidence. | Checking saved wording and its inputs does not prove every prose claim, interpretation or model sentence. |
| Tables and charts | Values, row/point coverage, order, chart caches and referenced embedded-workbook cells are compared with source-linked expectations. | Existing unmarked template content and arbitrary Office layout are outside these numeric checks. |
| Publication | The staged report is rechecked, its bytes and input fingerprints are checked again, and only a new output name is published. | The filesystem/runtime are trusted. These hashes are not signatures or a certificate of source truth. |

**Both a known mismatch and incomplete independent coverage prevent saving.** A partial preview is useful for diagnosis; the ordinary “create anyway” choice cannot bypass it. Correct the input, choose a supported reference or simplify an unsupported formula into independently checkable input. Keep the original workbook when making a separate values-only export; that export's values have no formula history for SlideFill to audit.

## Formula grammar and limits

`verification/formulas.py` does not call the production calculation functions. It supports finite numeric literals; cell/range references, including quoted sheet names and absolute/shared relative references; parentheses; `+ - * / %`; integer powers; numeric comparisons; and `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `ABS`, `INT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `IF`, `TRUE`, `FALSE`. Excel's negation/percent/exponent precedence is explicit. Only the supported numeric/Boolean/coercion cases are covered.

Array/data-table formulas, external links, unknown functions/names, cycles, nonnumeric comparisons and unsupported coercions are partial. XLS/XLSB formula cells are partial even when their cached value can be decoded: the binary formula expression is not independently evaluated. ODS supports only expressions that normalize into the same bounded grammar; repeated ODS formula cells remain unsupported. Blank/literal cells in those formats keep their existing value coverage.

Resource limits are 8,192 formula characters, 2,048 tokens, 64 levels of nesting/dependency depth, 100,000 dependency cells, 65,536 bits per exact result and integer exponents of magnitude at most 1,000. Exceeding a limit is incomplete coverage, never a passed calculation. Only the declared used/selection domain and its formula dependencies are audited; this is not a full-workbook audit.

The check compares exact integers exactly. Noninteger formula caches use a two-ULP comparison policy. An Excel floating-point result can disagree with the exact-arithmetic contract: for example a cancellation that loses an integer unit is rejected, rather than reproduced as correct. Recalculating and saving in Excel remains good input preparation, but is not sufficient to bypass this check.

## Precision and mathematical definitions

The production input decoder already supplies integers/floats. Exact `Fraction` calculations preserve the finite decimal spelling of those decoded operands; they cannot recover digits lost before that boundary. Independent source reconstruction rejects lost integers. A changed noninteger decimal-to-binary-to-short-decimal bridge is covered only when the source has at most 17 significant decimal digits and the change is at most one input ULP; higher/unknown precision is partial. An unchanged exact operand needs no such normalization. This is a coverage rule, not a theorem about every spreadsheet decoder.

For example, the aggregate `10^16 + 1 - 10^16` remains `1`, and `0.1 + 0.2` remains the exact rational `3/10` until formatting or a required numeric container conversion. Relative change is `(now - base) / abs(base)`; a zero base withholds it. Percentage-point changes are differences with the stated scale. Supported decimal display ties round away from zero.

Charts require a binary numeric container. The guard rejects a whole integer that cannot survive that conversion exactly. For an explicit supported number format it also checks that the cache and embedded data render to the required precision. An exact table decimal may therefore be accepted while its chart is refused. General-format output and two-ULP serialization checks do not promise that every original decimal digit is retained in a chart. The shared `numfmt` formatter remains a tested/trusted bridge, not an independently proved Office renderer.

An equal-row mean of rates is different from a rate computed as total numerator / total denominator. For rates `100/100` and `0/900`, those are `50%` and `10%`. Automatic rate-mean doubts expose this choice; the app does not invent missing denominators or weights. Confirm the required measure in **Change contents (Ctrl+R)**.

Word-linked comparisons independently reconstruct the source periods, the before/after window size, arithmetic means and change text. The event period is excluded. Mixed window sizes are stated per row instead of being mislabeled by one common header. The source quotation is checked against the DOCX text, and a changed DOCX blocks publication. This proves the supported calculation/quotation consistency, not that the event caused the numeric change.

## What a passed result supports

A passed result supports this statement: **the declared supported input values, calculations and saved representations passed the implemented checks for this report.** It does not certify source accuracy, intended scope, freshness, units, all prose or all visual layouts. Confirm those decisions and inspect the final client deck before sharing. No model is used to calculate these figures; optional AI wording remains a separate review decision.

The generated Dafny kernel proves its stated integer contracts, not these entire parsers, source recipes, algorithms or application. Finite exhaustive tests, mutation tests and curated native Excel/PowerPoint cases are complementary evidence, not an unseen-company accuracy score. See [verification design](verification-design.md), [numeric decoding and proof boundaries](binary-verification.md) and the [proof reproduction guide](../verification/proofs/README.md).

Original changes retain the project's existing MIT license. No dependency, model, weights or third-party parser source was added; existing component notices remain applicable.
