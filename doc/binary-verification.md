**English** | [한국어](binary-verification.ko.md)

# Binary sources, grouping and rounding

These extensions run in the existing [automatic app guard](verification.md). They need no Office installation or user checker command. Excel is used only to produce independent test fixtures and compare display results during development.

## Source formats

- `verification/compound.py` follows the root Workbook/Book stream through CFB FAT/DIFAT or MiniFAT chains, with bounds and cycle checks. `binary.py` reads supported BIFF2–5/8 XLS cells, shared Unicode strings/continuations, packed numbers, Boolean/error values, saved formula results, merges and the date epoch.
- The independent XLSB reader uses ZIP relationships and binary records for supported numeric/text/Boolean/error cells, formula caches, shared strings, styles, merges and the 1900/1904 epoch. Empty cached strings become blank for the existing calamine-based production adapter.
- HTML or delimited system exports named `.xls` have their own evidence path. Flat HTML tables retain spans and source cell positions, including row/header groups and optional cell/row end tags. Nested tables inside cells are retained as separate sheets with a bracketed sheet marker in the parent cell. Excessive/ambiguous spans and malformed nesting remain unsupported.
- ODS time-of-day durations with exact microsecond precision, including zero, become ISO time text. Negative, day-long and finer-precision durations retain their complete source literal, such as `-PT25H` or `PT0.0000001S`. Their notation and cell positions are checked; they are not converted into numeric elapsed-time measures. Year/month duration syntax and other unsupported forms remain partial. Stored formula results are read; formulas are not recalculated.
- Independent date reconstruction also handles supported month/year, quarter/year, bare numeric years and eight-digit dates, plus grouped numeric text such as `39 925`.

These are bounded readers, not complete file-format implementations. Compound inputs and individual binary ZIP parts have a 128 MiB budget; decoded value cells have a two-million-cell budget. Unsupported BIFF workbook/workspace variants and string/record grammars remain explicitly partial. Bounded BIFF2/3/4 single-sheet streams, BIFF5/95 workbooks and BIFF5/8 RSTRING values are covered as described below. Supported anchored XLSB short records are described below. Existing source formats and rendering restrictions still apply. Unsupported evidence never erases a known mismatch elsewhere.

The format facts come from Microsoft's [CFB header specification](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea), [XLS string specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/173d9f51-e5d3-43da-8de2-be7f22e119b9), [formula cache specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/39a0757a-c7bb-4e85-b144-3e7837b059d7) and [XLSB specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/). ODS value types follow the [OASIS schema](https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/OpenDocument-v1.3-os-part3-schema.html). No third-party parser source or runtime was copied; original code keeps the project's MIT license.

## Explicit grouping, without inferred totals

An XLSX-family sheet with adjacent explicit item/value headings and actual Excel row-outline levels can produce a descriptive grouping table. It retains each source row, its level and its original value, including parent values that do not equal child sums. A plain label such as “Parent” does not create a hierarchy. Blank rows end a block; the bounded grammar does not infer arbitrary pivot structures or business membership.

`layouts._outlines` creates the view; `rawxlsx` and `oracle` independently read levels and check row coverage. The report must retain the explanation that grouping metadata does not certify business membership. Both the in-app guard and the developer certificate checker reject removal of that explanation. Generation still asks the user to confirm the scope.

## Proved half-up rounding in the actual display path

`numfmt._digits` calls `kernel.round_decimal`, which converts the finite Decimal value to an exact fraction `n/d` and uses positive scale `s=10^places`. Python computes the nonnegative floor quotient of `abs(n)*s/d`; the bridge checks its defining integer bounds before calling generated `RoundScaled`.

The formal contract proves that the returned integer `r` has the half-up rounding interval:

```text
n >= 0: (2*r-1)*d <= 2*n*s <  (2*r+1)*d
n <  0: (2*r-1)*d <  2*n*s <= (2*r+1)*d
```

Ties therefore move away from zero. Decimal output is reconstructed from the integer coefficient and decimal exponent. The current kernel, including the aggregate bound below, has nine verified units, zero errors, for seven source definitions and six executable methods. Wrong tie, sign and remainder implementations were rejected. See [proof reproduction](../verification/proofs/README.md).

This proves the rounding decision for the supplied exact Decimal value, conditional on the checked quotient bounds and positive scale/denominator. Decimal/Fraction extraction, Python integer operations, earlier scaling/notation selection, text layout and the general floating-point aggregate error policy remain trusted/tested boundaries. It is not a proof of IEEE-754 hardware or every floating-point aggregation.

## Anchored XLSB short records

Records 12–18 omit the column field and advance one column after the preceding cell in the same row: blank, RK number, error, Boolean, real number, inline text and shared text. A full record must anchor each row. A short blank still advances the column; a full record can skip columns. Truncated payloads, repeated/out-of-order columns and grid overflow are rejected. `readers._expanded_xlsb` expands supported short cells into ordinary records in an in-memory ZIP for the existing calamine adapter. `verification.binary.xlsb` independently decodes the original bytes. Original files are never rewritten. Production also retains error cells and cached errors that calamine otherwise exposes as blank.

These record layouts are a compatibility extension described by the authors' [SheetJS format observations](https://github.com/SheetJS/notes/blob/main/xlsb_short_records/README.md), not a published Microsoft short-record specification. No parser code or vendor test workbook was copied. Two synthetic variants derived from native Excel fixtures, covering short blank/RK/shared-text records and both date epochs, were opened read-only in Excel 16.0 and their values compared. Seven short value types are covered by original synthetic writer/production/independent-reader tests. An unanchored-first-cell experiment failed native opening and became a rejection regression; this does not establish every reader's handling of that record placement.

## Exact aggregates and source-decoding error

Explicit grouped sum, product, mean, count and rate calculations keep the decimal spellings of decoded finite operands as exact Fractions. Products, row sums, group sums and mean division have no intermediate binary rounding. `kernel.sum_rationals` uses a common denominator and the actual generated `Sum` implementation; integers stay integers and nonintegers convert to float once at the public result boundary. For example, `10^16 + 1 - 10^16` is 1, and `0.1 + 0.2` is the final float representation of 0.3. This changes arithmetic, not group membership or which rows the user requested.

The independent recipe reader reconstructs each operand from the original file. Supported decimal/binary/short-decimal bridges are checked separately: exact integer loss is refused; a changed noninteger must have at most 17 significant decimal digits and differ by at most one input ULP. Finer precision is explicitly partial. The threshold is a coverage policy, not a theorem about all decoders.

Let `r_i` be an exact source row sum/product and `d_i` the independently reconstructed decoded row result. With `delta_i = d_i-r_i` and `e_i = abs(delta_i)`, the generated `SumBound` proves, for arbitrary finite integer-scaled sequences:

```text
S = sum(delta_i), E = sum(e_i), E >= 0, -E <= S <= E
```

The bridge establishes equal lengths, nonnegative error bounds and each `-e_i <= delta_i <= e_i` before calling generated code. Rational differences are scaled to a common positive denominator and restored afterward. The app audit checks that `S` equals the difference between the independently reconstructed decoded and source totals. For `N` retained rows, mean/rate division scales the total bound to `E/N`. A final float conversion contributes its exact rational conversion difference. There is no additional row-count ULP allowance. The resolved aggregate must equal the exact integer result or the single independently converted float; even a one-ULP alteration of that resolved value is rejected. Separate chart-serialization comparisons still have their documented two-ULP policy.

The current proof has nine verified units and zero errors. Three implementations with doubled deltas, a negative error budget or a missing last element were rejected under unchanged contracts. This proves the sum of supplied errors, not the source parser, Fraction/common-denominator conversion, decimal-to-binary semantics, every product implementation, hardware arithmetic or the whole application. Other statistics and ratios still have tested/trusted floating paths. Chart-cache and embedded-workbook checks remain active: a result that cannot survive chart serialization is blocked rather than accepted under a larger aggregate tolerance.

Original code and generated contracts retain the existing [MIT license](../LICENSE) (SPDX MIT, no numbered version). Retaining it fits the existing reusable-source project; redistribution must preserve its notice. No new dependency, model, external runtime or license scheme was introduced. Existing third-party notices still apply separately. Native Office fixtures remain ignored local evidence and are not distributed.

## Legacy BIFF and nested HTML checks

The bounded independent decoder now handles BIFF2/3/4 raw single-sheet streams, BIFF5/95 workbooks with declared byte encodings, and BIFF5/8 RSTRING values. The production path keeps using existing xlrd; the independent path does not import it. Legacy raw files are recognized as binary instead of delimited text. BIFF2 integer/number cells, supported string/Boolean/error/formula caches and legacy number-format indexes are checked. BIFF2 special formula-cache forms, BIFF4W/workspace variants, unsupported continuation/string grammars and unknown codepages remain partial. Missing codepages permit ASCII evidence only; no local encoding is guessed. Rich-text characters are preserved; character-level font runs are validated for bounds/order but are not reproduced as report styling. Formula results are cached values, not recalculated.

Native Excel16.0 generated two BIFF5 samples, one per date epoch. Their32typed cells and30explicit-format/text displays, with columns widened to avoid clipping, matched the app after normalizing legacy built-in date key14 to the canonical locale-date form. The other two General numeric displays were observed separately because Excel's display depends on column width. The comparison widens columns only in a read-only, unsaved document. BIFF2/3/4 and BIFF8 RSTRING checks use original synthetic writers; this does not claim execution on historical Excel versions. Format facts were checked against the authors' [OpenOffice file-format documentation](https://www.openoffice.org/sc/excelfileformat.pdf) and Microsoft's [file-format identifiers](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat). No third-party parser code or test data was copied. Original code retains MIT and adds no dependencies.

Each HTML table keeps its own grid in document order. A nested table must occur inside a parent cell. Its sheet is separate, and the parent retains only its own text plus a marker such as `Before [Nested 2] After`; inner numbers are not duplicated into the parent's measure cells. Tail rows, sibling tables, multilevel nesting, spans and optional cell/row closing tags are covered. Production uses an open-frame stack; the independent reader records table references and resolves them separately. Nesting is bounded at32levels,256tables and a shared2million-cell span budget. Malformed nesting, unclosed tables and ambiguous/overlapping spans are not claimed as supported; this is not a browser DOM/CSS layout engine.

The sample outer value46, inner values12/34 and outer tail10 are all retained and independently linked. Value changes, removed nested markers and saved-output corruption remain subject to the automatic guard. Very small tables can still receive no automatic report proposal; explicit template references were used for the recorded12report sample check. Automatic selection rules were not expanded for this task.
