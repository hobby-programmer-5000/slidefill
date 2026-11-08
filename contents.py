"""The contents window: tick what the report should hold, and see the slide it makes. Also a drop target for files. UI only."""

import atexit
import ctypes
import math
import os
import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from fractions import Fraction
from tkinter import ttk

import core
import suggest
from ui_timers import WindowTimers

FONT_FAMILY = "Malgun Gothic"
WM_DROPFILES = 0x0233
GWLP_WNDPROC = -4
SLIDE_INCHES = 13.333, 7.5
SERIES_COLORS = ("#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646")


class DropTarget:
    """Files dropped on the window, through Windows itself. Tk has no drag and drop of its own, and this needs no package. If it cannot be set up, Add files still works."""

    def __init__(self, root, on_files):
        self.root, self.on_files = root, on_files
        self.timers=WindowTimers(root)
        self.dropped = queue.SimpleQueue()
        self.hwnd = None
        try:
            user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
            result, handle = ctypes.c_ssize_t, ctypes.c_void_p
            proc_type = ctypes.WINFUNCTYPE(result, handle, ctypes.c_uint, handle, handle)
            user32.CallWindowProcW.restype, user32.CallWindowProcW.argtypes = result, (handle, handle, ctypes.c_uint, handle, handle)
            user32.SetWindowLongPtrW.restype, user32.SetWindowLongPtrW.argtypes = handle, (handle, ctypes.c_int, handle)
            user32.GetAncestor.restype, user32.GetAncestor.argtypes = handle, (handle, ctypes.c_uint)
            shell32.DragAcceptFiles.argtypes = (handle, ctypes.c_int)
            shell32.DragQueryFileW.restype, shell32.DragQueryFileW.argtypes = ctypes.c_uint, (handle, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint)
            shell32.DragFinish.argtypes = (handle,)
            self.user32, self.shell32 = user32, shell32
            root.update_idletasks()
            self.hwnd = user32.GetAncestor(root.winfo_id(), 2)
            shell32.DragAcceptFiles(self.hwnd, 1)
            self.proc = proc_type(self.handle)  # kept here: Windows calls it for as long as the window lives
            self.old = user32.SetWindowLongPtrW(self.hwnd, GWLP_WNDPROC, ctypes.cast(self.proc, handle))
        except (AttributeError, OSError):
            self.hwnd = None
            return
        root.bind("<Destroy>", self.closing, add="+")
        atexit.register(self.release)
        self.timers.after(100, self.collect)

    def handle(self, hwnd, message, wparam, lparam):
        # Windows calls this from inside Tk's own event handling, where calling Tk again ends the whole process. So the files only wait here.
        if message == WM_DROPFILES:
            count = self.shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
            paths = []
            for place in range(count):
                size = self.shell32.DragQueryFileW(wparam, place, None, 0) + 1
                buffer = ctypes.create_unicode_buffer(size)
                self.shell32.DragQueryFileW(wparam, place, buffer, size)
                paths.append(buffer.value)
            self.shell32.DragFinish(wparam)
            self.dropped.put(paths)
            return 0
        return self.user32.CallWindowProcW(self.old, hwnd, message, wparam, lparam)

    def collect(self):
        while not self.dropped.empty():
            self.on_files(self.dropped.get())
        self.timers.after(100, self.collect)

    def closing(self, event):
        if event.widget is self.root:
            self.release()

    def release(self):
        """Python must not be called for a window that outlives it."""
        if self.hwnd is not None:
            self.user32.SetWindowLongPtrW(self.hwnd, GWLP_WNDPROC, self.old)
            self.hwnd = None


def draw_chart(canvas, left, top, width, height, rows, bars, font, color_text):
    """A picture of the chart: enough to see what it will show, not PowerPoint's own drawing."""
    if not rows:
        return

    def numeric(value):
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, Fraction)):
            return Fraction(value)
        if isinstance(value, float) and math.isfinite(value):
            return Fraction(value)
        return None

    names = [str(cell[1]) for cell in rows[0][1:]]
    labels = [str(row[0][1]) for row in rows[1:]]
    series = [[numeric(row[c + 1][0]) if c + 1 < len(row) else None for row in rows[1:]] for c in range(len(names))]
    known = [value for line in series for value in line if value is not None]
    if not known or not labels:
        return
    high, low = max(max(known), 0), min(min(known), 0)
    span = (high - low) or 1
    plot_top, plot_bottom = top + 4, top + height - font * 3.4
    plot_left, plot_right = left + font * 5, left + width - 4
    if plot_right <= plot_left or plot_bottom <= plot_top:
        return
    # Normalize exactly before converting to screen coordinates. Converting huge
    # finite values first can overflow, and subtracting close floats loses scale.
    y_of = lambda value: plot_bottom - float((value - low) / span) * (plot_bottom - plot_top)  # noqa: E731
    baseline = y_of(0)
    canvas.create_line(plot_left, baseline, plot_right, baseline, fill="#808080")
    canvas.create_line(plot_left, plot_top, plot_left, plot_bottom, fill="#808080")
    for value, y, anchor in ((high, plot_top, "ne"), (low, plot_bottom, "se")):
        label = f"{value:,.0f}" if 10 <= abs(value) < 10**9 else f"{value:.3g}"
        canvas.create_text(plot_left - 3, y, text=label, anchor=anchor, font=(FONT_FAMILY, -font), fill=color_text)
    step = (plot_right - plot_left) / len(labels)
    for line, values in enumerate(series):
        color = SERIES_COLORS[line % len(SERIES_COLORS)]
        points = [None if value is None else (plot_left + step * (place + 0.5), y_of(value)) for place, value in enumerate(values)]
        if bars:
            for point in points:
                if point is not None:
                    x, y = point
                    canvas.create_rectangle(x - step * 0.3, min(y, baseline), x + step * 0.3, max(y, baseline), fill=color, outline="")
        else:
            run = []
            for point in [*points, None]:
                if point is None:
                    if len(run) >= 2:
                        canvas.create_line(*[coordinate for pair in run for coordinate in pair], fill=color, width=2)
                    run = []
                else:
                    run.append(point)
                    x, y = point
                    canvas.create_oval(x - 2, y - 2, x + 2, y + 2, fill=color, outline="")
    canvas.create_text(plot_left, plot_bottom + 2, text=labels[0][:12], anchor="nw", font=(FONT_FAMILY, -font), fill=color_text)
    canvas.create_text(plot_right, plot_bottom + 2, text=labels[-1][:12], anchor="ne", font=(FONT_FAMILY, -font), fill=color_text)
    x = plot_left
    for line, name in enumerate(names):
        canvas.create_rectangle(x, top + height - font * 1.4, x + font, top + height - font * 0.4, fill=SERIES_COLORS[line % len(SERIES_COLORS)], outline="")
        item = canvas.create_text(x + font * 1.3, top + height - font * 0.9, text=name[:18], anchor="w", font=(FONT_FAMILY, -font), fill=color_text)
        x = canvas.bbox(item)[2] + font


def draw_table(canvas, left, top, width, rows, row_height, font, columns=None):
    columns = columns or [width / max(len(rows[0]), 1)] * len(rows[0])
    for r, row in enumerate(rows):
        y = top + r * row_height
        canvas.create_rectangle(left, y, left + width, y + row_height, fill="#4F81BD" if r == 0 else "#D0D8E8" if r % 2 else "#E9EDF4", outline="white")
        x = left
        for column, (_raw, text, _fmt) in zip(columns, row):
            words = str(text if text is not None else "")
            limit = max(int(column / font * 1.6), 3)
            canvas.create_text(x + 3, y + row_height / 2, text=words if len(words) <= limit else words[:limit - 1] + "…", anchor="w",
                               font=(FONT_FAMILY, -font, "bold" if r == 0 else "normal"), fill="white" if r == 0 else "black")
            x += column


def draw_slide(canvas, words, section, table_rows, chart_rows, subtitle=True):
    """The slide as make_report will lay it out, using the same places on the slide, so a choice can be seen before it is made."""
    canvas.delete("all")
    width, height = canvas.winfo_width(), canvas.winfo_height()
    inches = min((width - 8) / SLIDE_INCHES[0], (height - 8) / SLIDE_INCHES[1])
    if inches <= 0:
        return
    origin_x, origin_y = (width - inches * SLIDE_INCHES[0]) / 2, (height - inches * SLIDE_INCHES[1]) / 2
    canvas.create_rectangle(origin_x, origin_y, origin_x + inches * SLIDE_INCHES[0], origin_y + inches * SLIDE_INCHES[1], fill="white", outline="#9AA0A6")
    x_of = lambda at: origin_x + at * inches  # noqa: E731
    y_of = lambda at: origin_y + at * inches  # noqa: E731
    points = lambda size: max(int(size / 72 * inches), 7)  # noqa: E731
    navy, muted, body = "#1F3864", "#5F6368", "#404040"
    text_left = x_of(suggest.LEFT + 0.1)
    canvas.create_text(text_left, y_of(0.35), text=words, anchor="nw", width=(suggest.FULL_WIDTH - 0.2) * inches, font=(FONT_FAMILY, -points(26), "bold"), fill=navy)
    if subtitle and section.heading and section.heading not in words:
        canvas.create_text(text_left, y_of(0.92), text=section.heading, anchor="nw", font=(FONT_FAMILY, -points(13)), fill=muted)
    canvas.create_line(text_left, y_of(suggest.RULE_AT), x_of(suggest.LEFT + suggest.FULL_WIDTH - 0.1), y_of(suggest.RULE_AT), fill=navy, width=2)
    side = bool(table_rows or chart_rows) and not section.wide
    room = suggest.TEXT_WIDTH if side else suggest.FULL_WIDTH
    y = y_of(suggest.CONTENT_TOP)
    for item in suggest._fit(section.sentences, room):
        quiet = item.code in suggest.QUIET_CODES
        if item.label:
            mark = canvas.create_text(text_left, y, text=item.label, anchor="nw", font=(FONT_FAMILY, -points(12)), fill=muted)
            y = canvas.bbox(mark)[3] + 2
        mark = canvas.create_text(text_left, y, text=item.result, anchor="nw", width=(room - 0.2) * inches, font=(FONT_FAMILY, -points(suggest._size(item)), "normal" if quiet else "bold"), fill=body if quiet else navy)
        y = canvas.bbox(mark)[3] + suggest.SENTENCE_GAP * inches
    if section.wide:
        if table_rows:
            draw_table(canvas, text_left, y, (suggest.FULL_WIDTH - 0.2) * inches, table_rows, suggest.WIDE_ROW * inches, points(suggest.WIDE_POINTS), [width_at * inches for width_at in suggest.WIDE_COLUMNS])
        return
    table_height = suggest.TABLE_ROW * len(table_rows) if table_rows else 0
    if chart_rows:
        draw_chart(canvas, x_of(suggest.RIGHT), y_of(suggest.CONTENT_TOP), suggest.RIGHT_WIDTH * inches, (suggest.SLIDE_BOTTOM - suggest.CONTENT_TOP - table_height - (0.15 if table_rows else 0)) * inches,
                   chart_rows, section.chart.code != "chart", points(11), muted)
    if table_rows:
        draw_table(canvas, x_of(suggest.RIGHT), y_of(suggest.SLIDE_BOTTOM - table_height), suggest.RIGHT_WIDTH * inches, table_rows, suggest.TABLE_ROW * inches, points(suggest.TABLE_POINTS))


@dataclass
class SheetChoice:
    """One sheet in the contents window."""
    file: int
    shape: object
    base: object
    found: list
    values: list


class ContentsWindow:
    """Nothing to learn, and nothing changes in the report until the person says so."""

    def __init__(self, app, files, on_detail):
        self.app, self.files, self.on_detail = app, list(files), on_detail
        self.plain = True
        self.key = (tuple(files), True)
        self.source = None
        self.choices = []
        self.picked = set()
        self.first = set()
        self.nodes = {}
        self.slides = []
        self.current = 0
        self.skipped = 0
        self.focus = None
        self.results = queue.Queue()
        self.job = 0
        self.working = False
        self.names = [os.path.splitext(os.path.basename(path))[0] for path in files]
        window = self.window = tk.Toplevel(app.root)
        self.timers=WindowTimers(window)
        window.title(app.t("helper_title_auto"))
        window.geometry(f"{app.px(1080)}x{app.px(680)}")
        window.transient(app.root)
        body = ttk.Frame(window, padding=app.px(12))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(2, weight=1)
        small = app.px(4)
        ttk.Label(body, text=app.t("helper_intro_auto"), wraplength=app.px(1040), justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, app.px(8)))
        ttk.Label(body, text=app.t("contents_figures"), style="Summary.TLabel").grid(row=1, column=0, sticky="w")
        top = ttk.Frame(body)
        top.grid(row=1, column=1, sticky="ew", padx=(app.px(12), 0))
        top.columnconfigure(1, weight=1)
        self.slide_label = ttk.Label(top, style="Summary.TLabel")
        self.slide_label.grid(row=0, column=0, sticky="w")
        self.bar = ttk.Progressbar(top, mode="indeterminate")
        self.bar.grid(row=0, column=1, sticky="ew", padx=app.px(10))
        self.back = ttk.Button(top, text="◀", width=3, command=lambda: self.turn(-1))
        self.back.grid(row=0, column=2)
        self.ahead = ttk.Button(top, text="▶", width=3, command=lambda: self.turn(1))
        self.ahead.grid(row=0, column=3, padx=(small, 0))
        left = ttk.Frame(body)
        left.grid(row=2, column=0, sticky="nsew", pady=(small, 0))
        left.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(left, show="tree", selectmode="none")
        self.tree.column("#0", width=app.px(320), stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<ButtonRelease-1>", self.clicked)
        self.tree.tag_configure("sheet", font=(FONT_FAMILY, 10, "bold"))
        self.canvas = tk.Canvas(body, background="#ECEFF1", highlightthickness=0)
        self.canvas.grid(row=2, column=1, sticky="nsew", padx=(app.px(12), 0), pady=(small, 0))
        self.canvas.bind("<Configure>", lambda _e: self.draw())
        self.change = ttk.Label(body, style="Summary.TLabel", wraplength=app.px(1040), justify="left")
        self.change.grid(row=3, column=0, columnspan=2, sticky="w", pady=(app.px(8), 0))
        foot = ttk.Frame(body)
        foot.grid(row=4, column=0, columnspan=2, sticky="e", pady=(small, 0))
        self.use_button = ttk.Button(foot, text=app.t("helper_use"), style="Accent.TButton", command=self.use, state="disabled")
        self.use_button.pack(side="left")
        self.detail_button = ttk.Button(foot, text=app.t("contents_detail"), command=self.open_detail)
        if len(self.files) == 1:  # the finer choices are about one file's rows
            self.detail_button.pack(side="left", padx=(small, 0))
        ttk.Button(foot, text=app.t("insight_close"), command=window.destroy).pack(side="left", padx=(small, 0))
        window.bind("<Escape>", lambda _e: window.destroy())
        window.bind("<Left>", lambda _e: self.turn(-1))
        window.bind("<Right>", lambda _e: self.turn(1))
        self.timers.after(50, self.poll)
        self.change.configure(text=app.t("helper_working"))
        self.run(self.load, self.loaded)

    # ---- background work: reading a sheet can take a while, and the window should stay alive

    def run(self, work, done):
        self.job += 1
        job, results = self.job, self.results
        self.working = True
        self.bar.grid()
        self.bar.start(20)

        def target():
            try:
                results.put((job, done, work(), None))
            except core.SlideFillError as exc:
                results.put((job, done, None, self.app.t("err_" + exc.code, **exc.params)))
            except Exception as exc:  # a thread has no one else to tell
                results.put((job, done, None, self.app.t("err_unexpected", detail=repr(exc))))

        threading.Thread(target=target, daemon=True).start()

    def poll(self):
        if not self.window.winfo_exists():
            return
        try:
            job, done, result, error = self.results.get_nowait()
        except queue.Empty:
            self.timers.after(60, self.poll)
            return
        if job == self.job:
            self.working = False
            self.bar.stop()
            self.bar.grid_remove()
            if error:
                self.change.configure(text=error)
            else:
                done(result)
        self.timers.after(30, self.poll)

    # ---- reading

    def load(self):
        lang = self.app.lang
        self.source = core.Sources(self.files)
        self.source.thousands = True
        self.names = self.source.names  # workbooks only: the documents are not ticked, they come along
        auto = self.app.auto
        now = {(section.file, section.sheet) for section in auto.sections} if auto and auto.excel == tuple(self.files) else set()
        choices = []
        for number, source in enumerate(self.source.sources):
            for ws in suggest.data_sheets(source):
                if ws.max_row > suggest.SMALL_SHEET and (number, ws.title) not in now and choices:
                    self.skipped += 1
                    continue
                shape, base, found = suggest.first_choice(source, ws.title, lang)
                if shape is not None:
                    choices.append(SheetChoice(number, shape, base, found, list(base.values)))
        return choices

    def loaded(self, choices):
        self.choices = choices
        auto = self.app.auto
        if auto and auto.excel == tuple(self.files):
            for section in auto.sections:
                for figure in section.figures:
                    self.picked.add((section.file, section.sheet, figure.name))
        self.first = set(self.picked)
        many = len(self.files) > 1
        for place, choice in enumerate(choices):
            parent = ""
            if many:
                parent = f"file{choice.file}"
                if not self.tree.exists(parent):
                    self.tree.insert("", "end", iid=parent, text=self.names[choice.file], open=True, tags=("sheet",))
            sheet = f"s{place}"
            has_pick = any((choice.file, choice.shape.sheet, value.name) in self.picked for value in choice.values)
            self.tree.insert(parent, "end", iid=sheet, text=choice.shape.sheet, open=has_pick, tags=() if many else ("sheet",))
            for number, value in enumerate(choice.values):
                node = f"s{place}v{number}"
                self.nodes[node] = (choice, value)
                self.tree.insert(sheet, "end", iid=node, text=self.figure_text(choice, value))
        self.rebuild()

    def figure_text(self, choice, value):
        name = self.app.t("helper_every_row") if value.every_row else self.app.t("helper_counted") if value.counted else value.name or self.app.t("helper_no_name")
        return ("☑ " if (choice.file, choice.shape.sheet, value.name) in self.picked else "☐ ") + name

    # ---- choosing

    def clicked(self, event):
        self.toggle(self.tree.identify_row(event.y))

    def toggle(self, node):
        if node not in self.nodes:
            return
        choice, value = self.nodes[node]
        mark = (choice.file, choice.shape.sheet, value.name)
        if mark in self.picked:
            self.picked.discard(mark)
        else:
            if sum(1 for picked in self.picked if picked[:2] == mark[:2]) >= suggest.MAX_TABLE_FIGURES:
                self.change.configure(text=self.app.t("contents_limit", count=suggest.MAX_TABLE_FIGURES))
                return
            self.picked.add(mark)
        self.tree.item(node, text=self.figure_text(choice, value))
        self.focus = mark[:2]
        self.rebuild()

    def rebuild(self):
        picked = set(self.picked)
        many = len(self.files) > 1
        lang = self.app.lang
        choices = list(self.choices)

        def work():
            sections = []
            for choice in choices:
                figures = [value for value in choice.values if (choice.file, choice.shape.sheet, value.name) in picked]
                if not figures:
                    continue
                section = suggest.section_for(self.source.sources[choice.file], choice.shape, choice.base, figures, choice.found, lang)
                if many:
                    seen = any(other.file == choice.file for other in sections)
                    suggest._for_file(section, choice.file, f"{self.names[choice.file]} · {section.heading}".rstrip(" ·") if seen else self.names[choice.file])
                sections.append(section)
            if many and sections:
                sections = suggest.with_documents(sections, self.source, lang)
            reader = self.source if many else self.source.sources[0]
            return [(section, suggest.preview_rows(reader, section.table), suggest.preview_rows(reader, section.chart)) for section in sections]

        self.run(work, self.rebuilt)

    def rebuilt(self, slides):
        self.slides = slides
        target = next((place for place, (section, _t, _c) in enumerate(slides) if self.focus and (section.file, section.sheet) == self.focus), None)
        self.current = target if target is not None else min(self.current, max(len(slides) - 1, 0))
        self.draw()
        self.show_change()

    def turn(self, step):
        if self.slides:
            self.current = max(0, min(len(self.slides) - 1, self.current + step))
            self.draw()

    def slide_words(self, place):
        section = self.slides[place][0]
        name = self.app.report_name(self.files, [s for s, _t, _c in self.slides if s.file >= 0], self.app.lang)
        return section.title or (name if place == 0 else section.heading or name)

    def draw(self):
        if not self.window.winfo_exists():
            return
        self.back.configure(state="normal" if self.current > 0 else "disabled")
        self.ahead.configure(state="normal" if self.current < len(self.slides) - 1 else "disabled")
        if not self.slides:
            self.canvas.delete("all")
            self.slide_label.configure(text=self.app.t("contents_preview_none"))
            return
        self.slide_label.configure(text=self.app.t("contents_preview", n=self.current + 1, total=len(self.slides)))
        section, table_rows, chart_rows = self.slides[self.current]
        draw_slide(self.canvas, self.slide_words(self.current), section, table_rows, chart_rows, self.current == 0 or bool(section.title))

    def show_change(self):
        sheets = {mark[:2] for mark in self.first}
        before = self.app.t("contents_count", slides=len(sheets) + (1 + len(self.source.documents) if len(self.files) > 1 and sheets else 0), figures=len(self.first))
        after = self.app.t("contents_count", slides=len(self.slides), figures=len(self.picked))
        changed = self.picked != self.first
        if not self.picked:
            words = self.app.t("contents_none")
        elif changed:
            words = self.app.t("contents_change", before=before, after=after)
        else:
            words = self.app.t("contents_same", now=before)
        if self.skipped:
            words += " " + self.app.t("contents_skipped", count=self.skipped)
        self.change.configure(text=words)
        self.use_button.configure(state="normal" if changed and self.picked else "disabled")

    def use(self):
        if self.working or self.app.busy:
            self.change.configure(text=self.app.t("helper_working"))
            return
        sections = [section for section, _t, _c in self.slides]
        what = " / ".join(suggest.describe(choice.shape, *[value for value in choice.values if (choice.file, choice.shape.sheet, value.name) in self.picked])
                          for choice in self.choices if any((choice.file, choice.shape.sheet, value.name) in self.picked for value in choice.values))
        self.app.use_sections(self.files, sections, what)
        self.window.destroy()

    def open_detail(self):
        self.window.destroy()
        self.on_detail()
