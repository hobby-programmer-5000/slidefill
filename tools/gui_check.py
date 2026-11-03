"""Runs the real window code through its shortcuts and buttons, using Tk events.

Covers: automatic preview, Ctrl+G, a second report never overwriting the first, the problem dialog (No and Yes),
a protected file's error message, the language switch, the help window, the date of a latest reference fitting
its column, the insight candidate window with a stub in place of the model, and settings surviving a restart.
Dialogs are answered by stubs that record what they were asked. Screenshots go to the given folder.

Run: .venv\\Scripts\\python.exe -X utf8 tools\\gui_check.py <work folder>
"""

import ctypes
import json
import os
import sys
import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import app  # noqa: E402
import core  # noqa: E402
import i18n  # noqa: E402
import insight  # noqa: E402
import openpyxl  # noqa: E402
from pptx import Presentation  # noqa: E402

work =os.path.abspath(sys.argv[1])
os.makedirs(work, exist_ok=True)
app.app_dir = lambda: work
TEMPLATE = os.path.join(ROOT, "samples", "sample_template.pptx")
EXCEL = os.path.join(ROOT, "samples", "sample_data.xlsx")
failures = []
asked = []
answer = [False]


def check(label, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + label + (f"  [{detail}]" if detail and not ok else ""))
    if not ok:
        failures.append(label)


messagebox.askyesno = lambda title, body: (asked.append(("ask", body)), answer[0])[1]
messagebox.showerror = lambda title, body: asked.append(("error", body))


def wait(root, instance, seconds=20):
    """Pump the event loop until the worker thread has finished."""
    root.update()
    limit = time.time() + seconds
    while time.time() < limit:
        root.update()
        if not instance.busy and instance.jobs.empty():
            for _ in range(5):
                root.update()
                time.sleep(0.02)
            if not instance.busy:
                return
        time.sleep(0.02)
    raise TimeoutError("worker did not finish")


def press(root, sequence):
    """A key only reaches a window with the focus, and on a desktop in use something else may hold it."""
    root.focus_force()
    root.update()
    root.event_generate(sequence)


def shot(root, name):
    from PIL import ImageGrab
    root.attributes("-topmost", True)
    root.update()
    time.sleep(0.3)
    root.update()
    box = (root.winfo_rootx(), root.winfo_rooty(),
           root.winfo_rootx() + root.winfo_width(), root.winfo_rooty() + root.winfo_height())
    ImageGrab.grab(box, all_screens=True).save(os.path.join(work, name + ".png"))
    root.attributes("-topmost", False)
    root.update()


def start(lang=None):
    root = tk.Tk()
    instance = app.App(root)
    if lang:
        instance.lang = lang
        instance.apply_language()
    wait(root, instance)
    return root, instance


def reports():
    return sorted(n for n in os.listdir(work) if n.endswith(".pptx") and n.startswith(("sample_template_", "issues_")))


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    settings = os.path.join(work, core.SETTINGS_FILE)
    core.save_settings(settings, {"template": TEMPLATE, "excel": EXCEL, "output": work, "lang": "ko"})

    root, ui = start()
    check("remembered paths preview automatically", len(ui.tree.get_children()) == 9, str(len(ui.tree.get_children())))
    check("summary counts the items", ui.summary.cget("text") == "정상 9 · 확인 0 · 문제 0", ui.summary.cget("text"))
    check("result buttons are off before a report exists", str(ui.open_button.cget("state")) == "disabled")
    shot(root, "preview_ko")

    press(root, "<Control-g>")
    wait(root, ui)
    check("Ctrl+G creates a report without any dialog", len(reports()) == 1 and not asked, f"{reports()} {asked}")
    check("result buttons turn on", str(ui.open_button.cget("state")) == "normal")
    check("status shows the saved path", ui.status[0] == "saved" and os.path.exists(ui.status[1]["path"]))
    text = Presentation(ui.last_output).slides[0].shapes[0].text_frame.text
    check("the report really holds the values", text == "2026년 10월 실적 보고", text)
    shot(root, "generated_ko")

    ui.generate_button.invoke()
    wait(root, ui)
    check("the button makes a second file, not an overwrite", len(reports()) == 2, str(reports()))

    press(root, "<Control-l>")
    root.update()
    check("Ctrl+L switches to English", ui.lang == "en" and ui.generate_button.cget("text").startswith("Generate"))
    check("the list follows the language", ui.tree.item(ui.tree.get_children()[0])["values"][0] == "OK")
    shot(root, "english")
    ui.tree.selection_set(ui.tree.get_children()[0])
    root.update()
    press(root, "<Control-l>")
    root.update()
    check("switching with a row selected leaves no stray text", ui.message.cget("text") == i18n.text("ko", "select_hint"),
          ui.message.cget("text"))
    press(root, "<Control-l>")
    root.update()

    press(root, "<F1>")
    root.update()
    helps = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
    check("F1 opens help", len(helps) == 1)
    shot(helps[0], "help_en")
    helps[0].destroy()
    root.update()

    # A template with a problem: the dialog asks first, and No writes nothing.
    prs = Presentation(TEMPLATE)
    prs.slides[0].shapes.add_textbox(0, 0, 914400, 914400).text_frame.text = "{{NoSuchName}}"
    issues = os.path.join(work, "issues.pptx")
    prs.save(issues)
    ui.vars["template"].set(issues)
    before = len(reports())
    answer[0] = False
    asked.clear()
    press(root, "<Control-g>")
    wait(root, ui)
    check("a problem asks before creating", len(asked) == 1 and asked[0][0] == "ask", str(asked))
    check("answering No writes nothing", len(reports()) == before, str(reports()))
    check("the problem row is listed first", ui.tree.item(ui.tree.get_children()[0])["values"][1] == "NoSuchName")
    shot(root, "problem_en")
    answer[0] = True
    asked.clear()
    press(root, "<Control-g>")
    wait(root, ui)
    check("answering Yes writes the report", len(reports()) == before + 1, str(reports()))
    box_text = [s.text_frame.text for s in Presentation(ui.last_output).slides[0].shapes if s.has_text_frame]
    check("the unresolved placeholder stays visible in the report", "{{NoSuchName}}" in box_text)

    # Latest references: the date is how you spot an old workbook, so the list must show it without clipping.
    daily = openpyxl.Workbook()
    daily.active.title = "일별 수송 인원 현황 자료"
    daily.active.append(["일자", "수송", "승차"])
    for day in range(1, 8):
        daily.active.append([f"2025/11/{day:02d}", day * 100, day * 10])  # in the past, or the check would depend on today
    daily_book = os.path.join(work, "daily.xlsx")
    daily.save(daily_book)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 914400, 914400).text_frame.text = "{{최신:일별 수송 인원 현황 자료!A,B}}"
    frame = slide.shapes.add_table(2, 3, 0, 914400, 914400, 914400)
    frame._element.xpath(".//p:cNvPr")[0].set("descr", "{{표:최신7:일별 수송 인원 현황 자료!A,B,C}}")
    daily_template = os.path.join(work, "daily.pptx")
    prs.save(daily_template)
    ui.vars["template"].set(daily_template)
    ui.vars["excel"].set(daily_book)
    press(root, "<F5>")
    wait(root, ui)
    sources = [str(ui.tree.item(row)["values"][3]) for row in ui.tree.get_children()]
    dates = {"2025-11-07 · ", "2025-11-01 ~ 2025-11-07 · "}
    check("the date comes first in the source column",
          len(sources) == 2 and all(any(s.startswith(d) for d in dates) for s in sources), str(sources))
    font = tkfont.nametofont("TkDefaultFont")
    room = ui.tree.column("source", "width")
    check("the widest date range fits the column", font.measure("2025-11-01 ~ 2025-11-07 · ") + ui.px(8) <= room,
          f"{font.measure('2025-11-01 ~ 2025-11-07 · ')} of {room}")
    root.geometry(f"{ui.px(720)}x{ui.px(460)}")
    root.update()
    room = ui.tree.column("source", "width")
    check("it still fits at the smallest window", font.measure("2025-11-01 ~ 2025-11-07 · ") + ui.px(8) <= room,
          f"{font.measure('2025-11-01 ~ 2025-11-07 · ')} of {room}")
    shot(root, "latest_small_window")
    root.geometry(f"{ui.px(1000)}x{ui.px(660)}")
    root.update()

    # Insight candidates, with the model replaced by a stub: nothing reaches the report until a person applies it.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 914400, 914400).text_frame.text = "{{통찰:일별 수송 인원 현황 자료!A,B}}"
    insight_template = os.path.join(work, "insight.pptx")
    prs.save(insight_template)
    ui.vars["template"].set(insight_template)
    press(root, "<F5>")
    wait(root, ui)
    check("the insight button turns on for an insight placeholder", str(ui.insight_button.cget("state")) == "normal")
    first = ui.tree.get_children()[0]
    check("an unchosen insight is listed as Check", ui.tree.item(first)["values"][0] == i18n.text(ui.lang, "status_warn"))
    calculated = str(ui.tree.item(first)["values"][2])
    check("the calculated sentence stands in", calculated.startswith("11월 7일") or calculated.startswith("Nov 7"), calculated)

    def stub_model(path, body, timeout):
        reply = [{"statement": "수송이 700으로 올랐습니다.", "fact_ids": ["value"], "hypothesis": "확인 필요: 행사 여부"},
                 {"statement": "연말 연휴로 인해 늘었습니다.", "fact_ids": ["value"], "hypothesis": ""}]
        return {"response": json.dumps({"insights": reply}, ensure_ascii=False)}

    def pump(window, seconds=10):
        limit = time.time() + seconds
        while window.working and time.time() < limit:
            root.update()
            time.sleep(0.02)
        for _ in range(10):
            root.update()
            time.sleep(0.02)

    real_post = insight.post
    insight.post = stub_model
    press(root, "<Control-i>")
    root.update()
    window = ui.insight_window
    check("Ctrl+I opens the candidate window", window is not None and window.window.winfo_exists())
    pump(window)
    check("only the sentence that passes the checks is offered", list(window.listbox.get(0, "end")) == ["수송이 700으로 올랐습니다."],
          str(window.listbox.get(0, "end")))
    check("the left-out count is shown", "2" in window.status.cget("text") or "3" in window.status.cget("text"), window.status.cget("text"))
    shot(window.window, "insight_window")
    check("nothing is chosen until the person acts", window.chosen.get("1.0", "end").strip() == "")
    window.listbox.selection_set(0)
    window.listbox.event_generate("<<ListboxSelect>>")
    root.update()
    basis = window.basis.get("1.0", "end")
    check("selecting a candidate shows it in full, with its basis and the ideas to check",
          "수송이 700으로 올랐습니다." in basis and "확인 필요: 행사 여부" in basis and "수송" in basis, basis)
    check("the basis text can be scrolled", bool(window.basis.cget("yscrollcommand")))
    window.add()
    check("the person can edit the sentence", window.chosen.get("1.0", "end").strip() == "수송이 700으로 올랐습니다.")
    window.chosen.insert("end", " 직접 고침.")
    window.apply_button.invoke()
    wait(root, ui)
    first = ui.tree.get_children()[0]
    check("the chosen sentence is in the list as OK", str(ui.tree.item(first)["values"][2]) == "수송이 700으로 올랐습니다. 직접 고침."
          and ui.tree.item(first)["values"][0] == i18n.text(ui.lang, "status_ok"), str(ui.tree.item(first)["values"]))
    press(root, "<Control-g>")
    wait(root, ui)
    check("the report holds exactly that sentence", Presentation(ui.last_output).slides[0].shapes[0].text_frame.text == "수송이 700으로 올랐습니다. 직접 고침.")
    check("the choice is saved beside the settings", "직접 고침" in open(ui.insights_file, encoding="utf-8").read())

    def stub_down(path, body, timeout):
        raise insight.InsightError("ollama_missing")

    insight.post = stub_down
    press(root, "<Control-i>")
    root.update()
    window = ui.insight_window
    pump(window)
    check("no Ollama is explained in the window", window.status.cget("text") == i18n.text(ui.lang, "insight_err_ollama_missing"), window.status.cget("text"))
    check("the saved sentence comes back into the box", "직접 고침" in window.chosen.get("1.0", "end"))
    window.window.destroy()
    root.update()
    insight.post = real_post
    ui.insights.clear()
    os.remove(ui.insights_file)
    ui.vars["excel"].set(EXCEL)

    # A mistyped reference: its message has a {key} of its own, and the list must still draw.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 914400, 914400).text_frame.text = "{{최신:일별 수송 인원 현황 자료!1X}}"
    typo_template = os.path.join(work, "typo.pptx")
    prs.save(typo_template)
    ui.vars["excel"].set(daily_book)
    ui.vars["template"].set(typo_template)
    press(root, "<F5>")
    wait(root, ui)
    shown = str(ui.tree.item(ui.tree.get_children()[0])["values"][2])
    check("a mistyped reference is listed with its message", "1X" in shown and ui.report.count("error") == 1, shown)

    # The helper: from a workbook alone to a filled report, with nothing typed.
    ui.vars["excel"].set(daily_book)
    ui.vars["template"].set("")
    press(root, "<Control-r>")
    root.update()
    helper = ui.helper_window
    check("Ctrl+R opens the helper", helper is not None and helper.window.winfo_exists())
    for _ in range(600):
        root.update()
        time.sleep(0.05)
        if not helper.working and helper.found:
            break
    check("the helper offers what works, with real results", any(item.code == "latest" and item.result == "700" for item in helper.found), str([(i.code, i.result) for i in helper.found]))
    check("the first row is selected and shows its reference", helper.key_var.get().startswith("{{"), helper.key_var.get())
    shot(helper.window, "helper")
    helper.copy()
    root.update()
    check("copying puts the reference on the clipboard", root.clipboard_get() == helper.key_var.get(), root.clipboard_get())
    helper.boxes["value"].current(1)
    helper.changed("value")
    for _ in range(200):
        root.update()
        time.sleep(0.05)
        if not helper.working:
            break
    check("choosing another figure recalculates", any(item.code == "latest" and item.result == "70" for item in helper.found), str([(i.code, i.result) for i in helper.found][:2]))
    ui.vars["output"].set(work)
    helper.make_starter()
    root.update()
    made = ui.vars["template"].get()
    check("a starter template is saved and becomes the template", made.endswith("_template.pptx") and os.path.exists(made), made)
    wait(root, ui)
    check("the starter fills with no problems", ui.report is not None and ui.report.count("error") == 0 and ui.report.items, str(ui.report and [(i.key, i.status) for i in ui.report.items]))
    helper.window.destroy()
    root.update()
    ui.vars["excel"].set("")
    press(root, "<Control-r>")
    root.update()
    check("with no workbook the helper says so", ui.status[0] == "helper_no_excel", str(ui.status))
    ui.vars["excel"].set(EXCEL)

    # A protected file: a message, not a crash, and the app keeps working.
    protected = os.path.join(work, "protected.xlsx")
    with open(protected, "wb") as fh:
        fh.write(core.OLE_MAGIC + b"\0" * 64)
    ui.vars["template"].set(TEMPLATE)
    ui.vars["excel"].set(protected)
    asked.clear()
    press(root, "<F5>")
    wait(root, ui)
    check("a protected file shows a message", len(asked) == 1 and "DRM" in asked[0][1], str(asked))
    ui.vars["excel"].set(EXCEL)
    press(root, "<F5>")
    wait(root, ui)
    check("the app recovers after an error", len(ui.tree.get_children()) == 9)

    root.destroy()
    root, ui = start()
    check("language survives a restart", ui.lang == "en")
    check("paths survive a restart", ui.vars["excel"].get() == EXCEL and ui.vars["template"].get() == TEMPLATE)
    shot(root, "restart_en")
    root.destroy()

    print("RESULT: " + ("FAIL " + str(len(failures)) if failures else "PASS"))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
