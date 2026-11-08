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

## GitHub commit attribution

The README records project roles and history. GitHub links commit authorship through commit metadata and an email associated with an account. Codex-assisted changes in this project include this trailer once, after a blank line:

```text
Co-authored-by: Codex <noreply@openai.com>
```

This is the identity used by [the official Codex source](https://github.com/openai/codex/blob/main/codex-rs/ext/git-attribution/src/world_state.rs). Follow [GitHub's co-author format](https://docs.github.com/en/pull-requests/how-tos/commit-changes/creating-a-commit-with-multiple-authors). Keep the human author and credit only tools that actually participated.

The v0.3.0 commit `280af57` describes Codex assistance in prose but omitted this trailer. This documentation follow-up includes the structured credit and records the policy for future commits. The published v0.3.0 tag and tested executable stay at their original snapshot.

## Release status — 2026-11-09

v0.3.0 was promoted from a pre-release to a regular GitHub release after the existing acceptance checks were complete. The version, release tag, executable and eight assets are the same verified artifacts. The documented input, validation and environment limits continue to apply.

## Unreleased numeric hardening — 2026-11-09

Codex continued with the user on an isolated development branch based on `13dfc27`. This work does not change the published v0.3.0 executable. It adds bounded independent formula-cache/dependency checks, exact rational calculation intermediates, independently reconstructed numeric text and Word means, chart precision/format checks, and refusal to save when independent coverage is incomplete. See [the enforced contract and remaining limits](reliability-hardening.md).

The final source passed **675 core tests and 137 verifier tests** during `build.bat`. A fresh Dafny 4.11.0 run verified the existing nine kernel contracts with zero errors; its generated class matched the packaged kernel. This proves the bounded kernel, not the complete application.

Curated workbooks saved by Excel 16.0.20326.20232 supplied 26 supported formula cases; two unsupported comparison cases were explicitly refused. A cancellation cache that differed from exact arithmetic and deliberately stale caches were refused. Native PowerPoint 16.0.20326.20232 opened/rendered four reports/four slides and opened both chart workbooks. The English/Korean reports, complete formula table and exact large-decimal table were visually inspected. An initial oversized test table produced an overflow warning and clipped; its evidence was retained and the synthetic template margins were corrected before repeating the render. This is not a general layout guarantee.

The UI passed scripted English/Korean checks at 900×600 and 1200×800, including save refusal and window callback cleanup. These regression, mutation and native Office cases are not a new unseen-company accuracy measurement. No fresh VM or Windows 11 run was performed for this development build. Source truth, business scope, weights, freshness, all prose and arbitrary template layout still need the review described in the disclaimer. No new dependency or license scheme was introduced.

The local validation executable is 18,694,459 bytes, SHA256 `3d11c97604efaa8f5c41751e8ed773e254f7fedb84a02bc8b0bc74ff1d342ff8`. Its build smoke and five exact-artifact cases passed: valid generation, stale-cache refusal, unsupported-formula refusal, chart-precision refusal and preservation of an existing output file. It is not a new published release.
