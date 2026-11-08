"""Runs the real window code through its shortcuts and buttons, using Tk events.

Covers: automatic preview, Ctrl+G, a second report never overwriting the first, the problem dialog (No and Yes),
a protected file's error message, the language switch, the help window, the date of a latest reference fitting
its column, the insight candidate window with a stub in place of the model, and settings surviving a restart.
Dialogs are answered by stubs that record what they were asked. Screenshots go to the given folder.

Run: .venv\\Scripts\\python.exe -X utf8 tools\\gui_check.py <work folder>
"""

import ctypes
import ctypes.wintypes
import json
import os
import sys
import threading
import time
import tkinter as tk
import zipfile
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
messagebox.showinfo = lambda title, body: asked.append(("info", body))


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


def shot(window, name):
    """PrintWindow draws this window alone. A grab of the screen would take whatever else the person has open in front of it."""
    from PIL import Image
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    window.update()
    time.sleep(0.2)
    window.update()
    hwnd = user32.GetAncestor(window.winfo_id(), 2)
    outer = ctypes.wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(outer))
    width, height = outer.right - outer.left, outer.bottom - outer.top
    screen = user32.GetDC(hwnd)
    memory = gdi32.CreateCompatibleDC(screen)
    bitmap = gdi32.CreateCompatibleBitmap(screen, width, height)
    gdi32.SelectObject(memory, bitmap)
    user32.PrintWindow(hwnd, memory, 2)  # PW_RENDERFULLCONTENT
    header = (ctypes.c_uint32 * 10)(40, width, ctypes.c_uint32(-height).value, 1 | (32 << 16), 0, 0, 0, 0, 0, 0)
    pixels = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(memory, bitmap, 0, height, pixels, header, 0)
    Image.frombuffer("RGB", (width, height), pixels, "raw", "BGRX", 0, 1).save(os.path.join(work, name + ".png"))
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(memory)
    user32.ReleaseDC(hwnd, screen)


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
    verified = ui.report.verification
    expected_status = i18n.text("ko", "select_hint") + ("\n" + i18n.text("ko", "verification_" + verified["state"], checked=verified["checked"]) if verified else "")
    check("switching with a row selected leaves no stray text", ui.message.cget("text") == expected_status,
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
    written = [row for row in window.listbox.get(0, "end") if not row.startswith(i18n.text(ui.lang, "insight_mark_calculated"))]
    check("calculated sentences are there, marked as such", any(c.calculated for c in window.shown), str(window.listbox.get(0, "end")))
    check("only the model's sentence that passes the checks is offered", [c.statement for c in window.shown if not c.calculated] == ["수송이 700으로 올랐습니다."],
          str(window.listbox.get(0, "end")))
    check("the model's sentence is marked to be checked", any(row.startswith(i18n.text(ui.lang, "insight_mark_model")) for row in written), str(written))
    model_row = next(place for place, c in enumerate(window.shown) if not c.calculated)
    check("the left-out count is shown", "2" in window.status.cget("text") or "3" in window.status.cget("text"), window.status.cget("text"))
    shot(window.window, "insight_window")
    check("nothing is chosen until the person acts", window.chosen.get("1.0", "end").strip() == "")
    window.listbox.selection_set(model_row)
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

    # A workbook alone: no template asked for, and a report comes out named after the workbook.
    def settle(helper):
        for _ in range(600):
            root.update()
            time.sleep(0.05)
            if not helper.working and helper.found:
                break

    ui.own_template.set(False)
    ui.show_template_row()
    root.update()
    check("the template row is hidden unless asked for", not any(w.winfo_ismapped() for w in ui.template_row))
    check("with no template only the buttons that do something are shown", not ui.preview_button.winfo_ismapped() and ui.insight_button.winfo_ismapped() and ui.helper_button.winfo_ismapped())
    ui.vars["excel"].set(daily_book)
    ui.vars["output"].set(work)
    press(root, "<F5>")
    wait(root, ui)
    check("a workbook alone gives a preview", ui.report is not None and ui.report.count("error") == 0 and ui.report.items, str(ui.report))
    check("the status says what was chosen", ui.status[0] == "auto_ready" and "수송" in ui.status[1]["what"], str(ui.status))
    listed = [str(ui.tree.item(row)["values"][1]) for row in ui.tree.get_children()]
    check("the list has no notation in it", all("!" not in name and "{{" not in name for name in listed), str(listed))
    okay = [str(ui.tree.item(row)["values"][1]) for row in ui.tree.get_children() if ui.tree_items[row].status == "ok"]  # anything to check comes before these
    check("the list reads in the report's order, sentences first", okay[0].endswith(i18n.text(ui.lang, "suggest_note")), str(okay))
    origins = [str(ui.tree.item(row)["values"][3]) for row in ui.tree.get_children()]
    check("nor a cell address in where it comes from", origins and all("!" not in origin for origin in origins), str(origins))
    ui.tree.selection_set(ui.tree.get_children()[0])
    root.update()
    check("nor when a row is selected", "!" not in ui.message.cget("text") and "{{" not in ui.message.cget("text"), ui.message.cget("text"))
    check("both figures reach the report", "승차" in ui.auto.what and any(name.startswith("승차 · ") for name in listed), str(listed))
    check("the working template is not beside the workbook", os.path.dirname(ui.auto.path) != work, ui.auto.path)
    shot(root, "workbook_only")
    press(root, "<Control-g>")
    wait(root, ui)
    check("the report is named after the workbook", os.path.basename(ui.last_output).startswith("daily_") and os.path.exists(ui.last_output), ui.last_output)
    words = " ".join(s.text_frame.text for s in Presentation(ui.last_output).slides[0].shapes if s.has_text_frame)
    check("that report holds the newest value and no placeholder", "700" in words and "{{" not in words, words)
    check("the choice to go without a template survives", core.load_settings(settings).get("use_template") is False)

    # Insight candidates with only a workbook: the model is a stub, and nothing reaches the report until a person applies it.
    check("the insight button works with no template", str(ui.insight_button.cget("state")) == "normal")
    gate = threading.Event()

    def slow_model(path, body, timeout):
        gate.wait(10)
        return stub_model(path, body, timeout)

    insight.post = slow_model
    first_note = next(row for row in ui.tree.get_children() if ui.tree_items[row].key.startswith("해설") or ui.tree_items[row].key.startswith("note"))
    ui.tree.selection_set(first_note)
    root.update()
    press(root, "<Control-i>")
    root.update()
    window = ui.insight_window
    for _ in range(5):
        root.update()
        time.sleep(0.02)
    list_top = window.listbox.winfo_rooty()
    check("while it works a moving bar shows right above the list", window.bar.winfo_ismapped() and window.bar.winfo_rooty() < list_top and window.bar.winfo_rooty() + window.bar.winfo_height() <= list_top + 2)
    check("and says how far along it is", "1" in window.status.cget("text") and "3" in window.status.cget("text"), window.status.cget("text"))
    shot(window.window, "insight_working")
    gate.set()
    pump(window)
    check("the bar goes when it is done", not window.bar.winfo_ismapped())
    check("the candidates are about that figure, without notation", [c.statement for c in window.shown if not c.calculated] == ["수송이 700으로 올랐습니다."], str(window.listbox.get(0, "end")))
    shot(window.window, "insight_done")
    window.listbox.selection_set(next(place for place, c in enumerate(window.shown) if not c.calculated))
    window.listbox.event_generate("<<ListboxSelect>>")
    root.update()
    window.add()
    window.apply_button.invoke()
    wait(root, ui)
    press(root, "<Control-g>")
    wait(root, ui)
    words = " ".join(s.text_frame.text for s in Presentation(ui.last_output).slides[0].shapes if s.has_text_frame)
    check("the chosen sentence takes the place of that figure's reading", "수송이 700으로 올랐습니다." in words, words)
    check("the figure's own sentence stays", "수송 700" in words, words)
    check("no placeholder is left", "{{" not in words, words)
    chosen_row = next(ui.tree.item(row)["values"] for row in ui.tree.get_children() if "수송이 700" in str(ui.tree.item(row)["values"][2]))
    check("the chosen sentence is listed by a plain name", chosen_row[0] == i18n.text(ui.lang, "status_ok") and "!" not in str(chosen_row[1]), str(chosen_row))
    insight.post = real_post
    ui.insights.clear()
    ui.auto.stale = True
    os.remove(ui.insights_file)

    # Changing the contents: tick figures and see the slide before anything changes.
    def calm(window):
        for _ in range(300):
            root.update()
            time.sleep(0.03)
            if not window.working:
                break
        for _ in range(5):
            root.update()
            time.sleep(0.02)

    press(root, "<Control-r>")
    root.update()
    window = ui.helper_window
    calm(window)
    calm(window)
    node_of = lambda name: next(node for node, (_choice, value) in window.nodes.items() if value.name == name)  # noqa: E731
    check("the contents window ticks what the report holds now", window.tree.item(node_of("수송"))["text"].startswith("☑") and window.tree.item(node_of("승차"))["text"].startswith("☑"), str(window.picked))
    check("it draws the slide", window.slides and len(window.canvas.find_all()) > 10, str(len(window.canvas.find_all())))
    check("and nothing to use until something changes", str(window.use_button.cget("state")) == "disabled")
    shot(window.window, "contents_same")
    before_report = ui.auto.sections
    window.toggle(node_of("승차"))
    calm(window)
    check("unticking a figure redraws the slide without it", not any("승차" in item.result for item in window.slides[0][0].sentences), str([item.result for item in window.slides[0][0].sentences]))
    check("it says what will change", "1" in window.change.cget("text") and "2" in window.change.cget("text"), window.change.cget("text"))
    check("the report has not changed yet", ui.auto.sections is before_report)
    shot(window.window, "contents_changed")
    window.toggle(node_of("승차"))
    window.toggle(node_of("수송"))
    calm(window)
    window.use_button.invoke()
    wait(root, ui)
    words = [str(ui.tree.item(row)["values"][2]) for row in ui.tree.get_children()]
    check("the ticked figure goes in the report, and the other comes out", any("승차 70" in w for w in words) and not any("수송 700" in w for w in words), str(words))
    press(root, "<Control-g>")
    wait(root, ui)
    made_words = " ".join(s.text_frame.text for s in Presentation(ui.last_output).slides[0].shapes if s.has_text_frame)
    check("what was drawn is what is made", "승차 70" in made_words and "수송 700" not in made_words, made_words)

    # The finer choices are still there for one file.
    ui.open_detail()
    root.update()
    helper = ui.helper_window
    settle(helper)
    check("with no template the helper hides the notation", helper.plain and str(helper.copy_button.winfo_ismapped()) == "0" and helper.tree["displaycolumns"] == ("what", "result"), str(helper.tree["displaycolumns"]))
    shot(helper.window, "helper_plain")
    offered = list(helper.boxes["value"]["values"]) + list(helper.boxes["base"]["values"])
    check("the helper offers columns by their headings", offered and all(" · " not in text for text in offered) and "수송" in offered, str(offered))
    rows = [str(helper.tree.item(row)["values"][0]) for row in helper.tree.get_children()]
    check("with no template the helper lists only what the report holds", i18n.text(ui.lang, "suggest_prev") not in rows and i18n.text(ui.lang, "suggest_note") in rows, str(rows))
    helper.boxes["value"].current(1)
    helper.changed("value")
    for _ in range(200):
        root.update()
        time.sleep(0.05)
        if not helper.working:
            break
    helper.use_button.invoke()
    wait(root, ui)
    words = [str(ui.tree.item(row)["values"][2]) for row in ui.tree.get_children()]
    check("another figure from the helper goes into the report", ui.auto.chosen and any("승차 70" in w for w in words), str(words))
    ui.lang = "en" if ui.lang == "ko" else "ko"
    press(root, "<Control-l>")
    wait(root, ui)
    check("the person's choice survives a language switch", ui.auto.chosen and any("승차 70" in str(ui.tree.item(row)["values"][2]) for row in ui.tree.get_children()))

    # Dropping files on the window, through the real Windows message.
    def drop_on_window(paths):
        names = ("\0".join(paths) + "\0\0").encode("utf-16-le")
        kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32
        kernel32.GlobalAlloc.restype, kernel32.GlobalLock.restype = ctypes.c_void_p, ctypes.c_void_p
        kernel32.GlobalLock.argtypes = kernel32.GlobalUnlock.argtypes = (ctypes.c_void_p,)
        user32.PostMessageW.argtypes = (ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p)
        block = kernel32.GlobalAlloc(0x42, 20 + len(names))  # GHND
        pointer = kernel32.GlobalLock(block)
        ctypes.memmove(pointer, (ctypes.c_uint32 * 5)(20, 0, 0, 0, 1), 20)  # DROPFILES: names follow at 20, wide characters
        ctypes.memmove(pointer + 20, names, len(names))
        kernel32.GlobalUnlock(block)
        user32.PostMessageW(ui.drop.hwnd, 0x0233, block, None)  # WM_DROPFILES
        for _ in range(15):
            root.update()
            time.sleep(0.02)

    second_book = os.path.join(work, "second.xlsx")
    second = openpyxl.Workbook()
    second.active.title = "고객"
    second.active.append(["일자", "방문"])
    for day in range(1, 8):
        second.active.append([f"2025/11/{day:02d}", day * 11])
    second.save(second_book)
    data_csv = os.path.join(work, "third.csv")
    with open(data_csv, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("일자,판매\n" + "".join(f"2025/11/{day:02d},{day * 5}\n" for day in range(1, 8)))
    note_file = os.path.join(work, "note.pdf")
    open(note_file, "w").close()
    ui.own_template.set(False)
    ui.show_template_row()
    ui.more_files = []
    ui.vars["excel"].set("")
    ui.auto = None
    root.update()
    check("with no files the list says what to do", ui.drop_hint.winfo_ismapped())
    shot(root, "drop_empty")
    check("the window accepts dropped files", ui.drop.hwnd is not None)
    drop_on_window([daily_book, second_book, data_csv, note_file])
    wait(root, ui)
    check("dropped files are listed, and another kind of file is not", ui.inputs() == [daily_book, second_book, data_csv], str(ui.inputs()))
    check("the hint goes", not ui.drop_hint.winfo_ismapped())
    check("a report from three files is previewed", ui.report is not None and ui.report.count("error") == 0 and len(ui.report.items) > 6, str(ui.report and [(i.key, i.status) for i in ui.report.items][:4]))
    check("the refused file is named", any("note.pdf" in body for kind, body in asked) or "note.pdf" in ui.message.cget("text"), str(asked))
    shot(root, "three_files")
    drop_on_window([second_book])
    wait(root, ui)
    check("the same file is not added twice", len(ui.inputs()) == 3, str(ui.inputs()))
    press(root, "<Control-g>")
    wait(root, ui)
    many = Presentation(ui.last_output).slides
    check("three files make an overview and three slides", len(many) == 4 and os.path.basename(ui.last_output).startswith("daily_"), f"{len(many)} {ui.last_output}")
    overview_words = " ".join(s.text_frame.text for s in many[0].shapes if s.has_text_frame)
    check("the overview has each file's main sentence", "수송" in overview_words and "방문" in overview_words and "판매" in overview_words, overview_words)
    press(root, "<Control-r>")
    root.update()
    window = ui.helper_window
    calm(window)
    calm(window)
    check("the contents window groups figures by file", window.tree.exists("file0") and window.tree.exists("file2") and len(window.slides) == 4, str(len(window.slides)))
    shot(window.window, "contents_files")
    window.window.destroy()
    insight.post = stub_model
    press(root, "<Control-i>")
    root.update()
    window = ui.insight_window
    check("the insight window offers all the files together first", "(" in window.about.get() and core.ALL_FILES in window.request.key, window.about.get())
    check("what the window shows as the basis holds no instruction to the model", "cross" not in window.basis.get("1.0", "end"), window.basis.get("1.0", "end")[:120])
    pump(window)
    shot(window.window, "insight_files")
    window.window.destroy()
    root.update()
    insight.post = real_post
    # A Word document among the files: its dated entries are lined up with the figures.
    word_file = os.path.join(work, "record.docx")
    with zipfile.ZipFile(word_file, "w") as package:
        package.writestr("word/document.xml", '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
                         '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>기록 문서</w:t></w:r></w:p>'
                         '<w:p><w:r><w:t>2025년 11월 4일 수송 노선을 늘렸습니다.</w:t></w:r></w:p></w:body></w:document>')
    drop_on_window([word_file])
    wait(root, ui)
    check("a Word file can be dropped in", word_file in ui.inputs(), str(ui.inputs()))
    named = [str(ui.tree.item(row)["values"][1]) for row in ui.tree.get_children()]
    check("the document is in the list by a plain name", i18n.text(ui.lang, "suggest_doc_note") in named and not any("@" in name for name in named), str(named))
    press(root, "<Control-g>")
    wait(root, ui)
    with_word = Presentation(ui.last_output).slides
    check("the report has a slide for the document, after the overview", len(with_word) == 5 and "기록 문서" in " ".join(s.text_frame.text for s in with_word[1].shapes if s.has_text_frame), str(len(with_word)))
    press(root, "<Control-r>")
    root.update()
    window = ui.helper_window
    calm(window)
    calm(window)
    check("the contents window draws the document's slide too", len(window.slides) == 5 and window.slides[1][0].wide, str(len(window.slides)))
    window.turn(1)
    root.update()
    shot(window.window, "contents_document")
    window.window.destroy()
    root.update()
    ui.file_list.selection_clear(0, "end")
    ui.file_list.selection_set(3)
    ui.show_files_state()
    ui.remove_button.invoke()
    wait(root, ui)
    check("removing the document leaves the workbooks", ui.inputs() == [daily_book, second_book, data_csv], str(ui.inputs()))
    ui.file_list.selection_clear(0, "end")
    ui.more_files = []
    ui.vars["excel"].set("")
    drop_on_window([word_file])
    wait(root, ui)
    check("a Word file alone asks for a workbook", any("Excel" in body or "엑셀" in body for kind, body in asked if kind == "error"), str(asked[-1:]))
    ui.more_files = [second_book, data_csv]
    ui.vars["excel"].set(daily_book)
    ui.show_files()
    ui.preview()
    wait(root, ui)
    ui.file_list.selection_set(1, 2)
    ui.show_files_state()
    ui.remove_button.invoke()
    wait(root, ui)
    check("removing files leaves the rest", ui.inputs() == [daily_book], str(ui.inputs()))
    check("more files are remembered", core.load_settings(settings).get("excel") == daily_book)
    ui.more_files = []

    other = openpyxl.Workbook()
    other.active.title = "메모"
    other.active.append(["그냥 글자"])
    other_book = os.path.join(work, "memo.xlsx")
    other.save(other_book)
    ui.vars["excel"].set(other_book)
    asked.clear()
    press(root, "<F5>")
    wait(root, ui)
    check("a workbook with nothing to read says so", len(asked) == 1 and asked[0][0] == "error" and "F1" in asked[0][1], str(asked))
    ui.vars["excel"].set(EXCEL)
    ui.preview()
    ui.vars["excel"].set(daily_book)
    ui.preview()
    wait(root, ui)
    wait(root, ui)
    check("choosing a file while another is being read ends on the new one", ui.auto.excel == (daily_book,) and not ui.again, str(ui.auto.excel))

    # The helper with a template in play: the references, to copy into it.
    ui.own_template.set(True)
    ui.show_template_row()
    root.update()
    check("the template row comes back when asked for", all(w.winfo_ismapped() for w in ui.template_row))
    check("and so does its Preview button", ui.preview_button.winfo_ismapped() and ui.insight_button.winfo_ismapped())
    ui.vars["excel"].set(daily_book)
    ui.vars["template"].set(daily_template)
    press(root, "<Control-r>")
    root.update()
    helper = ui.helper_window
    check("Ctrl+R opens the helper", helper is not None and helper.window.winfo_exists() and not helper.plain)
    settle(helper)
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

    # A choice the rules cannot be sure of: nothing is made until a person has looked.
    doubtful = os.path.join(work, "doubtful.xlsx")
    book = openpyxl.Workbook()
    book.active.title = "Example"
    for row in [["지역", "매출"]] + [[f"지역{n}", 100 + n * 37 % 91] for n in range(8)]:
        book.active.append(row)
    book.save(doubtful)
    template_before, output_before = ui.vars["template"].get(), ui.vars["output"].get()
    ui.vars["template"].set("")
    ui.own_template.set(False)
    ui.more_files = []
    ui.vars["excel"].set(doubtful)
    ui.vars["output"].set(work)
    press(root, "<F5>")
    wait(root, ui)
    check("a doubtful choice says so in the status line", ui.status[0] == "auto_ready_doubt", str(ui.status))
    made_before = [name for name in os.listdir(work) if name.startswith("doubtful") and name.endswith(".pptx")]
    press(root, "<Control-g>")
    wait(root, ui)
    window = ui.doubt_window
    check("creating asks first and makes nothing yet", window is not None and window.window.winfo_exists() and [name for name in os.listdir(work) if name.startswith("doubtful") and name.endswith(".pptx")] == made_before)
    check("it says why in plain words", bool(window and window.reasons) and all("_" not in reason for reason in window.reasons), str(window.reasons if window else None))
    shot(window.window, "doubt_en" if ui.lang == "en" else "doubt_ko")
    window.cancel_button.invoke()
    root.update()
    check("cancel makes nothing and asks again next time", ui.auto.doubts and [name for name in os.listdir(work) if name.startswith("doubtful") and name.endswith(".pptx")] == made_before)
    press(root, "<Control-g>")
    wait(root, ui)
    window = ui.doubt_window
    window.keep_button.invoke()
    root.update()
    wait(root, ui)
    check("keeping it makes the report", len([name for name in os.listdir(work) if name.startswith("doubtful") and name.endswith(".pptx")]) == len(made_before) + 1 and ui.status[0] == "saved", str(ui.status))
    kept_words = " ".join(shape.text_frame.text for slide in Presentation(ui.last_output).slides for shape in slide.shapes if shape.has_text_frame)
    check("the report says what was kept unconfirmed", "※" in kept_words, kept_words[:200])
    press(root, "<Control-g>")
    wait(root, ui)
    check("having kept it, it does not ask again", ui.doubt_window is None and ui.status[0] == "saved", str(ui.status))
    ui.vars["template"].set(template_before)
    ui.vars["output"].set(output_before)
    ui.own_template.set(True)
    ui.vars["excel"].set(EXCEL)
    press(root, "<F5>")
    wait(root, ui)

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
