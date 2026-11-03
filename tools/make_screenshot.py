"""Writes doc/images/screenshot.png and screenshot.ko.png from the sample files.

Paths are relative so no personal folder names end up in the pictures.

Run from the project root: .venv\\Scripts\\python.exe -X utf8 tools\\make_screenshot.py
"""

import ctypes
import os
import sys
import tempfile
import time
import tkinter as tk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import app  # noqa: E402
import core  # noqa: E402
from PIL import ImageGrab  # noqa: E402

ctypes.windll.shcore.SetProcessDpiAwareness(1)
scratch = tempfile.mkdtemp(prefix="slidefill_shot_")
app.app_dir = lambda: scratch

for lang, name in (("en", "screenshot.png"), ("ko", "screenshot.ko.png")):
    core.save_settings(os.path.join(scratch, core.SETTINGS_FILE), {
        "template": r"samples\sample_template.pptx", "excel": r"samples\sample_data.xlsx",
        "output": "samples", "lang": lang})
    root = tk.Tk()
    ui = app.App(root)
    root.attributes("-topmost", True)
    limit = time.time() + 20
    while time.time() < limit and (ui.busy or not ui.tree.get_children()):
        root.update()
        time.sleep(0.02)
    root.update()
    time.sleep(0.5)
    root.update()
    box = (root.winfo_rootx(), root.winfo_rooty(),
           root.winfo_rootx() + root.winfo_width(), root.winfo_rooty() + root.winfo_height())
    ImageGrab.grab(box, all_screens=True).save(os.path.join(ROOT, "doc", "images", name))
    root.destroy()
print("Wrote doc/images")
