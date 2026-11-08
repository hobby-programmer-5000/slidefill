**English** | [한국어](development-history.ko.md)

# Development and verification history

SlideFill is the user's project. The user sets the idea, requirements, scope and validation decisions. Claude and Codex are AI development assistants; these credits describe the work performed, not an independent certification of correctness.

## Before the eighth measurement

Claude (Anthropic), working with the user in Claude Code, contributed the original design and implementation, public-file measurements through the seventh batch, and the v0.1.0/v0.2.0 releases. Existing history and Claude's credit are retained. The published v0.2.0 notes remain the record of that release.

## Codex participation: eighth measurement through v0.3.0

Codex (OpenAI) joined for the eighth measurement on frozen v0.2.0 code and continued through this release. This includes the measurement itself, subsequent implementation, verification, documentation and release preparation.

- **Eighth measurement:** eight previously unseen public workbooks; six generated reports, two refused layouts. The first passed model-sentence batches contained 18/18 source-consistent sentences, a small and limited observation. Native PowerPoint opened six reports/nine slides. Empty charts, missing years/units and correct sentences being filtered were recorded before fixes. This is a v0.2.0 result, not a new v0.3.0 accuracy score.
- **Corrections and bounded layout support:** fixed the observed chart/year/unit/filter regressions; stopped assuming an unknown fiscal-year end; supported stated multi-level headers, population-band cost tables, repeated organization/job blocks and explicit row outlines within documented grammars. Ambiguous business meanings remain a user's decision.
- **Automatic CPU verification:** linked independent source reconstruction, exact rational aggregates, displayed values, chart caches and embedded workbooks to preview/save. Known mismatches block saving; incomplete coverage is stated. No new local model or app dependency was added. A small Dafny-proved integer/rational/rounding kernel is used by the app; source conversions, UI, general statistics and all business meanings are not thereby proved. See [design and trust boundaries](verification-design.md) and [binary/rounding coverage](binary-verification.md).
- **Build and distribution checks:** 675 core tests and 91 verification tests passed in the actual build, followed by packaged-EXE smoke tests, source/archive privacy checks and component-notice verification. Personal settings, public measurement downloads, private evidence, VM disks and Windows media are excluded from distribution.
- **Environment and rendering:** the exact release EXE passed 19 cases in a fresh offline Windows 10 Pro 19045.3803 VM without Office or separate Python, producing 18 reports. Native PowerPoint 16.0.20326.20232 on the host rendered those 18 reports and 18 host reports (44 slides). Guest slide images matched the visually reviewed host images byte for byte. The installed VM is retained for reuse. VM filming could not be completed reliably, so the README demo uses the same UI source on the Windows 10 host. See [the full environment scope](testing-alternatives.md).

The executable is 18,655,237 bytes, SHA256 `d19d0404311b22f42fa6dc36edc6c384e591030d59d6906179dae6efe373d64f`. Documentation and demo changes do not change this tested runtime snapshot.

## What these results do not establish

The eighth batch became regression data after fixes. There is no new fair unseen-file accuracy measurement after those fixes. Private company files, Windows 11, a second physical PC, an independent human click-through and every chart's native Edit Data interaction are untested. Automated demo operation is not a human usability study. The app can still choose an unwanted measure and unsupported inputs can remain partial. A passed check is limited to its stated contract; it is not a whole-app proof or a guarantee that the source data is true.

Original project code and documentation retain [MIT](../LICENSE); bundled components retain [their own terms](../THIRD-PARTY-NOTICES.md). The demo uses original synthetic data. Neither private Windows media nor third-party source workbooks/templates are distributed with the app.
