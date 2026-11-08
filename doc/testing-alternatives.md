**English** | [한국어](testing-alternatives.ko.md)

# Alternatives for environment and usage tests

These checks supplement [the existing limits](limits.md). Automated simulation is recorded as simulation; public sample files are not described as private company files. Subsequent exact-build, Office-rendering and clean-Windows results are recorded below.

| Missing evidence | Available substitute | What it establishes |

|---|---|---|

| Approved private business files | An official public business-style sample, with source/hash and permission basis recorded locally | Compatibility with that sample and its schema, not with a particular company's files. |

| Optional person's click-through (not required now) | Run the real Tk UI with synthetic inputs, isolated settings, events and recorded dialog responses; inspect screenshots | Programmed paths and visible states, not human comprehension or unplanned interaction. |

| Windows without Office | A fresh Windows Sandbox or VM containing the tested app and no Office | The tested artifact runs in that clean environment, if the check actually executes. Host-side Office rendering or an ordinary filesystem sandbox is insufficient. |

## Public sample and rights

Microsoft explicitly provides a [Financial Sample workbook for trying data](https://learn.microsoft.com/en-us/power-bi/create-reports/sample-financial-download). This is sample sales/profit data, not a private customer workbook. It was downloaded through that page, fingerprinted and used locally for preview, automatic checking and saving. It is kept outside the source/distribution tree. The article's invitation to try the sample is the basis for this local test; no general redistribution or commercial template rights are asserted. Do not apply another repository's MIT label to this workbook.

The eight earlier public inputs remain regression cases. The additional sample is an additional compatibility observation; it is not a statistically meaningful unseen-company success rate. Original project code/documentation remain MIT, and third-party data rights remain separate.

## Earlier test decision

At an earlier point on2026-11-08, the user accepted existing sample use and automated sample checks for that scope. A separate human trial and Office-free Windows trial were **SKIPPED by user choice** for that earlier check. The later request below reopens the Windows trial; the human trial remains excluded. Real-use issues can be addressed when reported. No Windows feature, VM or alternative office suite was set up for that earlier check.

An Office-free Windows test would check SlideFill's own ability to launch and save without Office installed; it would not target an open-source office suite. File-format conventions and runtime dependencies are separate questions. No clean-Windows result was asserted from that earlier, excluded check.

## Earlier source and old-EXE runtime check (2026-11-08)

A subsequent request tested the application itself on the current Windows host. Current source and the unchanged older EXE were copied into isolated test folders with synthetic settings and sample inputs; no personal settings were copied. Keyboard save commands through their actual windows produced three source reports and one EXE report. Each version's built-in smoke mode also ran valid input, invalid template and valid input again, with the expected exit codes0/1/0. Relaunch recovery was checked in separate processes.

The source smoke processes installed a process-local Python audit hook denying socket connections, subprocess/startfile execution and win32com/comtypes imports. All three source cases completed without those attempted calls or loaded Office automation packages. This exercises the tested report paths under those Python restrictions; it does not emulate every native Windows/Shell call or uninstall Office. The copied older EXE was tested separately and does not contain the latest source changes.

Eight generated PPTX files passed headless ZIP, summary text,7×4table, two chart series/six points each, chart-cache and embedded-workbook comparisons against the original sample. Repeat saving retained existing output bytes. No Office application was used to inspect or render these results during this test. Source/EXE/input hashes were preserved.

Computer Use screen capture timed out and native owned-dialog interaction was unavailable; no new visual or same-process UI error-recovery pass is claimed. The source chooser test process was stopped after verifying its identity, and the normal invalid-input/relaunch checks used the supported smoke entry point. Windows Sandbox was absent and the available Docker/WSL environments were Linux, so that earlier test did not establish a clean Office-free Windows result. No Windows feature, virtual machine or alternative office suite was installed. A separate human trial remains excluded.

## Exact new EXE, Office and clean Windows (2026-11-08–09)

The later request explicitly added Office rendering and a clean Windows trial. The unchanged tested EXE has SHA256 `d19d0404311b22f42fa6dc36edc6c384e591030d59d6906179dae6efe373d64f`. Its build passed 675 core tests and 91 verification selftests and built-EXE smoke checks before these additional checks.

PowerPoint 16.0.20326.20232 opened18host-generated reports and18actual guest-generated reports read-only, exporting22slides from each set. Each set's12tables and8charts, including series values and categories, matched the saved PPTX. All22guest PNGs were byte-identical to the22host PNGs already inspected through16distinct images. Korean text, custom decimals, ODS duration strings and aggregate charts were readable, and no shapes extended outside the slides. The negative-value test chart retained the template's axis-at-zero layout. Original report bytes were preserved.

With the user's approval, an existing Windows 10 Pro installation ISO and the already-installed VirtualBox 7.2.16 created a new VM with 1 CPU(s) and 4 GB RAM with a fresh dynamic disk. The guest reported Windows 10 Pro build 19045.3803, no installed Office or separate Python, and no host-validation override. Network, shared folders and Guest Additions were absent. Only the exact EXE, its license/notices, original synthetic inputs and a Windows PowerShell checker were supplied on a read-only payload ISO. Authentic guest summary/events and reports were recovered from the private VDI/NTFS disk using host-only portable7-Zip26.06. No extraction tool was installed in the guest or added to the app.

All 19 guest cases passed: valid/invalid/valid-relaunch, BIFF2/3/4/5 and Korean BIFF5, BIFF8 rich text, nested HTML parent/child tables, XLSB short cells, ODS durations and exact grouped sum/mean/product. Eighteen reports matched the checked text and chart-cache baselines, and all payload/input hashes stayed unchanged. The guest was shut down after testing. This is evidence for the tested Windows 10 VM and artifact; it does not establish Windows 11, a second physical PC, a human trial or every input/template. No app runtime code or dependency changed, so no rebuild was needed. Existing original MIT and bundled third-party notices remain unchanged; the private Windows installation/media are not app distribution assets.
