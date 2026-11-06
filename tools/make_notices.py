"""Writes THIRD-PARTY-NOTICES.md from the license files of what the exe actually bundles.

Fails if a bundled library has no license text, or if something the notices call excluded got into the build.

Run after build.bat: .venv\\Scripts\\python.exe -X utf8 tools\\make_notices.py
"""

import importlib.metadata as metadata
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOC = os.path.join(ROOT, "build", "SlideFill", "Analysis-00.toc")
COMMON = os.path.join(ROOT, "tools", "notices_common.md")
READERS = os.path.join(ROOT, "tools", "notices_readers.md")  # xlrd and python-calamine, with the Rust crates built into calamine: kept by hand
READER_ROWS = [("xlrd", "Python package, for `.xls`", "BSD-3-Clause, with parts under a BSD-style license (in its license text)"),
               ("python-calamine", "Python package with a compiled module, for `.xlsb` and `.ods`", "MIT; the Rust crates built into it are listed below")]
PYTHON_LICENSE = os.path.join(os.path.dirname(os.path.realpath(sys.base_prefix + os.sep + "python.exe")), "LICENSE.txt")

PACKAGES = [("python-pptx", "MIT"), ("openpyxl", "MIT"), ("et_xmlfile", "MIT"), ("lxml", "BSD-3-Clause, and libxml2/libxslt under MIT"),
            ("pillow", "MIT-CMU, and the libraries listed in its license file"), ("xlsxwriter", "BSD-2-Clause"),
            ("typing_extensions", "PSF-2.0")]
EXCLUDED = ("_ssl", "_hashlib", "_bz2", "_lzma", "_zstd", "libssl", "libcrypto")

EXPAT = """Copyright (c) 1998-2000 Thai Open Source Software Center Ltd and Clark Cooper
Copyright (c) 2001-2025 Expat maintainers

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."""

MPDECIMAL = """Copyright (c) 2008-2024 Stefan Krah. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions
are met:

1. Redistributions of source code must retain the above copyright
   notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY
OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
SUCH DAMAGE."""


def license_text(name):
    dist = metadata.distribution(name)
    texts = []
    for item in dist.files or []:
        base = os.path.basename(str(item)).upper()
        if ".DIST-INFO" in str(item).upper() and base.startswith(("LICEN", "COPYING")):
            with open(item.locate(), encoding="utf-8", errors="replace") as fh:
                texts.append(fh.read().strip())
    if not texts:
        raise SystemExit(f"No license text found for {name}")
    return "\n\n".join(texts), dist.version


def fence(text):
    return "```text\n" + text + "\n```"


def main():
    with open(TOC, encoding="utf-8", errors="replace") as fh:
        toc = fh.read()
    binaries = {os.path.basename(b) for b in re.findall(r"'([^']+\.(?:dll|pyd))'", toc)}
    leaked = sorted(b for b in binaries if b.lower().startswith(EXCLUDED))
    if leaked:
        raise SystemExit(f"Excluded libraries got into the build: {leaked}")

    with open(READERS, encoding="utf-8") as fh:
        readers = fh.read().strip()
    rows, sections = [], []
    for name, files, license_name in READER_ROWS:
        version = metadata.version(name)
        if f"## {name} {version}" not in readers:  # a new version may bring other crates: the hand-kept text must be redone, not left stale
            raise SystemExit(f"{READERS} has no section for {name} {version}")
        rows.append(f"| {name} | {version} | {files} | {license_name} |")
    for name, license_name in PACKAGES:
        text, version = license_text(name)
        rows.append(f"| {name} | {version} | Python package | {license_name} |")
        sections.append(f"## {name} {version}\n\n{fence(text)}")

    with open(PYTHON_LICENSE, encoding="utf-8", errors="replace") as fh:
        python_text = fh.read().strip()
    with open(COMMON, encoding="utf-8") as fh:
        common = fh.read().strip()

    head = f"""# Third-party notices

SlideFill itself is under the MIT License (see `LICENSE`). This file covers what the Windows executable `SlideFill.exe` bundles, each part under its own license.

| Component | Version | Files in `SlideFill.exe` | License |
|---|---|---|---|
| Python (CPython) | {sys.version.split()[0]} | `python314.dll`, `_tkinter.pyd`, `_ctypes.pyd`, `_socket.pyd`, `select.pyd`, `_elementtree.pyd`, `_queue.pyd`, `unicodedata.pyd`, standard library | PSF License Agreement v2 |
| Tcl / Tk | 9.0.4 | `tcl90.dll`, `tcl9tk90.dll` | Tcl/Tk license, BSD-style (in Python's license file) |
| Microsoft Visual C++ Runtime / Universal CRT | as shipped with Python | `VCRUNTIME140.dll`, `ucrtbase.dll`, `api-ms-win-*.dll` | Microsoft Distributable Code (conditions in Python's license file) |
| libffi | 3.x | `libffi-8.dll` | MIT |
| LibTomMath | as shipped with Tcl 9 | `libtommath.dll` | The LibTom license / Unlicense |
| zlib / zlib-ng | 1.3.1.zlib-ng | `zlib1.dll`; zlib-ng is built into `python314.dll` | zlib License |
| Expat | 2.8.2 | built into `pyexpat.pyd` | MIT |
| mpdecimal | 4.0.0 | built into `_decimal.pyd` | BSD-2-Clause |
| Unicode Character Database | 16.0.0 | data inside `unicodedata.pyd` | Unicode License v3 |
{chr(10).join(rows)}
| PyInstaller bootloader | {metadata.version("pyinstaller")} | executable stub | GPL-2.0-or-later **with the PyInstaller bootloader exception**, which allows distributing the resulting executable under any license ([details](https://pyinstaller.org/en/stable/license.html)) |

The build excludes TLS and bzip2, xz and Zstandard compression (see `build.bat`), so OpenSSL, bzip2, xz and Zstandard are **not** in `SlideFill.exe`. `tools/make_notices.py` fails if any of them appears. Python's license file below is reproduced unchanged and may still mention them.

The socket modules are present because the standard library's XML helpers import them, and because the optional insight step talks to Ollama at 127.0.0.1 on the same PC. SlideFill makes no other network connections.

---

## Python {sys.version.split()[0]} (including Tcl/Tk and the Windows binary build conditions)

{fence(python_text)}

---

{common}

---

## Expat

{fence(EXPAT)}

## mpdecimal

{fence(MPDECIMAL)}

---

{readers}

---

"""
    with open(os.path.join(ROOT, "THIRD-PARTY-NOTICES.md"), "w", encoding="utf-8", newline="\n") as fh:
        fh.write(head + "\n\n".join(sections) + "\n")
    print(f"Wrote THIRD-PARTY-NOTICES.md ({len(binaries)} binaries checked, none excluded leaked)")


if __name__ == "__main__":
    main()
