@echo off
rem Builds the single-file executable dist\SlideFill.exe.
rem Unit tests run first, and the finished exe must fill the samples: any failure stops the build.
rem First time only: uv venv --python 3.14 .venv ^&^& uv pip install --python .venv\Scripts\python.exe openpyxl python-pptx xlrd python-calamine pyinstaller
rem
rem The excluded modules are parts the app never uses. Remove an exclude only if the app starts
rem importing it, and update THIRD-PARTY-NOTICES.md at the same time.
rem socket and select must stay: python-pptx imports xml.sax.saxutils, which imports urllib and http.client.
rem The only network call in the app is insight.py, to the local model server on 127.0.0.1.
cd /d "%~dp0"
.venv\Scripts\python.exe -X utf8 test_core.py || exit /b 1
.venv\Scripts\python.exe -X utf8 -m unittest test_native_vba test_native_vba_cases || exit /b 1
.venv\Scripts\python.exe -X utf8 -m verification selftest || exit /b 1
.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name SlideFill ^
  --icon app.ico --add-data "app.ico;." ^
  --add-data "native_vba_host.ps1;." --add-data "assets\slidefill-vba.xlsm;assets" ^
  --exclude-module unittest --exclude-module pydoc --exclude-module test ^
  --exclude-module ssl --exclude-module _ssl --exclude-module _hashlib ^
  --exclude-module multiprocessing --exclude-module asyncio --exclude-module concurrent ^
  --exclude-module bz2 --exclude-module _bz2 --exclude-module lzma --exclude-module _lzma ^
  --exclude-module compression.zstd --exclude-module _zstd ^
  --exclude-module PIL.ImageTk --exclude-module PIL._imagingtk --exclude-module PIL.ImageQt ^
  --exclude-module PIL.AvifImagePlugin --exclude-module PIL._avif ^
  --exclude-module PIL.WebPImagePlugin --exclude-module PIL._webp ^
  --exclude-module PIL.ImageCms --exclude-module PIL._imagingcms ^
  --exclude-module lxml.html --exclude-module lxml.objectify --exclude-module lxml.sax --exclude-module lxml.isoschematron ^
  --exclude-module numpy --exclude-module setuptools --exclude-module pkg_resources ^
  app.py || exit /b 1

rem Run the real exe on the samples from a clean temp folder, with a time limit (a startup error shows a dialog and waits).
powershell -NoProfile -ExecutionPolicy Bypass -File tools\smoke.ps1 || exit /b 1
echo Built dist\SlideFill.exe
