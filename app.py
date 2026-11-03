"""SlideFill window. UI only: all file work happens in core.py, on a worker thread."""

import ctypes
import os
import queue
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import traceback
from tkinter import filedialog, messagebox, ttk

import core
import i18n
import insight
import suggest

STATUS_ORDER = {"error": 0, "warn": 1, "ok": 2}
STATUS_COLOR = {"error": "#B3261E", "warn": "#9A6700", "ok": "#1A7F37"}
FONT_FAMILY = "Malgun Gothic"
SETTINGS_KEYS = ("template", "excel", "output", "lang")


def app_dir():
    return os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))


def resource(name):
    return os.path.join(getattr(sys, "_MEIPASS", app_dir()), name)


def system_language():
    """Korean on a Korean Windows, English everywhere else."""
    try:
        return "ko" if ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF == 0x12 else "en"
    except (AttributeError, OSError):
        return "en"


class App:
    def __init__(self, root):
        self.root = root
        self.settings_file = core.settings_path(app_dir())
        self.insights_file = core.settings_path(app_dir(), core.INSIGHTS_FILE)
        self.insights = core.load_settings(self.insights_file)
        self.insight_window = None
        self.helper_window = None
        saved = core.load_settings(self.settings_file)
        self.lang = saved.get("lang") if saved.get("lang") in i18n.STRINGS else system_language()
        self.vars = {key: tk.StringVar(value=str(saved.get(key) or "")) for key in ("template", "excel", "output")}
        self.jobs = queue.Queue()
        self.job_id = 0
        self.busy = False
        self.report = None
        self.status = ("start_hint", {})
        self.last_output = ""
        self.tree_items = {}

        self.scale = root.winfo_fpixels("1i") / 96
        self.build()
        self.apply_language()
        self.bind_keys()
        root.after(50, self.poll)
        if self.ready():
            self.preview()

    def t(self, key, /, **params):  # positional only: some messages have a {key} of their own
        return i18n.text(self.lang, key, **params)

    def px(self, value):
        return int(value * self.scale)

    # ------------------------------------------------------------ layout

    def build(self):
        root = self.root
        root.geometry(f"{self.px(1000)}x{self.px(660)}")
        root.minsize(self.px(720), self.px(460))
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            tkfont.nametofont(name).configure(family=FONT_FAMILY, size=10)
        line = tkfont.nametofont("TkDefaultFont").metrics("linespace")
        style = ttk.Style(root)
        style.configure("Treeview", rowheight=line + self.px(8))
        style.configure("Title.TLabel", font=(FONT_FAMILY, 15, "bold"))
        style.configure("Muted.TLabel", foreground="#5F6368")
        style.configure("Summary.TLabel", font=(FONT_FAMILY, 10, "bold"))
        style.configure("Accent.TButton", font=(FONT_FAMILY, 10, "bold"))

        pad = self.px(16)
        main = ttk.Frame(root, padding=pad)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        head = ttk.Frame(main)
        head.grid(row=0, column=0, sticky="ew", pady=(0, self.px(12)))
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text="SlideFill", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.tagline = ttk.Label(head, style="Muted.TLabel")
        self.tagline.grid(row=1, column=0, sticky="w")
        self.help_button = ttk.Button(head, command=self.show_help)
        self.help_button.grid(row=0, column=1, rowspan=2, padx=(0, self.px(6)))
        self.lang_button = ttk.Button(head, command=self.toggle_language)
        self.lang_button.grid(row=0, column=2, rowspan=2)

        files = ttk.Frame(main)
        files.grid(row=1, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)
        self.file_labels, self.browse_buttons = {}, {}
        commands = {"template": self.pick_template, "excel": self.pick_excel, "output": self.pick_output}
        for row, key in enumerate(("template", "excel", "output")):
            label = ttk.Label(files)
            label.grid(row=row, column=0, sticky="w", padx=(0, self.px(10)), pady=self.px(3))
            entry = ttk.Entry(files, textvariable=self.vars[key])
            entry.grid(row=row, column=1, sticky="ew", pady=self.px(3))
            entry.bind("<Return>", lambda _e: self.preview())
            entry.bind("<FocusOut>", lambda _e: self.save_settings())
            button = ttk.Button(files, command=commands[key])
            button.grid(row=row, column=2, padx=(self.px(6), 0), pady=self.px(3))
            self.file_labels[key], self.browse_buttons[key] = label, button

        actions = ttk.Frame(main)
        actions.grid(row=2, column=0, sticky="ew", pady=self.px(12))
        actions.columnconfigure(4, weight=1)
        self.generate_button = ttk.Button(actions, style="Accent.TButton", command=self.generate)
        self.generate_button.grid(row=0, column=0, ipady=self.px(3))
        self.preview_button = ttk.Button(actions, command=self.preview)
        self.preview_button.grid(row=0, column=1, padx=self.px(6), ipady=self.px(3))
        self.insight_button = ttk.Button(actions, command=self.open_insight, state="disabled")
        self.insight_button.grid(row=0, column=2, ipady=self.px(3))
        self.helper_button = ttk.Button(actions, command=self.open_helper)
        self.helper_button.grid(row=0, column=3, padx=(self.px(6), 0), ipady=self.px(3))
        self.summary = ttk.Label(actions, style="Summary.TLabel", anchor="e")
        self.summary.grid(row=0, column=4, sticky="ew")

        table = ttk.Frame(main)
        table.grid(row=3, column=0, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        columns = ("status", "item", "value", "source", "where")
        self.tree = ttk.Treeview(table, columns=columns, show="headings", selectmode="browse")
        widths = {"status": 70, "item": 210, "value": 240, "source": 240, "where": 160}
        for col in columns:
            self.tree.column(col, width=self.px(widths[col]), stretch=col in ("item", "value"),
                             anchor="center" if col == "status" else "w")
        for status, color in STATUS_COLOR.items():
            self.tree.tag_configure(status, foreground=color if status != "ok" else "")
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.show_detail())

        foot = ttk.Frame(main)
        foot.grid(row=4, column=0, sticky="ew", pady=(self.px(10), 0))
        foot.columnconfigure(0, weight=1)
        self.message = ttk.Label(foot, wraplength=self.px(680), justify="left")
        self.message.grid(row=0, column=0, sticky="w")
        self.open_button = ttk.Button(foot, command=lambda: self.open_path(self.last_output))
        self.folder_button = ttk.Button(
            foot, command=lambda: self.open_path(os.path.dirname(self.last_output)))
        self.open_button.grid(row=0, column=1, padx=(self.px(6), 0))
        self.folder_button.grid(row=0, column=2, padx=(self.px(6), 0))
        self.update_result_buttons()
        foot.bind("<Configure>", lambda e: self.message.configure(wraplength=max(e.width - self.px(240), 200)))

    def bind_keys(self):
        bindings = {"<F5>": self.preview, "<Control-g>": self.generate, "<Control-G>": self.generate,
                    "<Control-t>": self.pick_template, "<Control-T>": self.pick_template,
                    "<Control-e>": self.pick_excel, "<Control-E>": self.pick_excel,
                    "<Control-l>": self.toggle_language, "<Control-L>": self.toggle_language,
                    "<Control-i>": self.open_insight, "<Control-I>": self.open_insight,
                    "<Control-r>": self.open_helper, "<Control-R>": self.open_helper,
                    "<F1>": self.show_help}
        for sequence, handler in bindings.items():
            self.root.bind(sequence, lambda _e, h=handler: h())

    # ------------------------------------------------------------ language and rendering

    def apply_language(self):
        self.root.title(self.t("window_title"))
        self.tagline.configure(text=self.t("tagline"))
        self.help_button.configure(text=self.t("help"))
        self.lang_button.configure(text=self.t("language"))
        for key, label in self.file_labels.items():
            label.configure(text=self.t(key))
            self.browse_buttons[key].configure(text=self.t("browse"))
        self.generate_button.configure(text=self.t("generate"))
        self.preview_button.configure(text=self.t("preview"))
        self.insight_button.configure(text=self.t("insight_button"))
        self.helper_button.configure(text=self.t("helper_button"))
        self.open_button.configure(text=self.t("open_result"))
        self.folder_button.configure(text=self.t("open_folder"))
        for col, key in (("status", "col_status"), ("item", "col_item"), ("value", "col_value"),
                         ("source", "col_source"), ("where", "col_where")):
            self.tree.heading(col, text=self.t(key), anchor="center" if col == "status" else "w")
        self.render_report()
        if self.status[0] == "detail":  # rebuilding the list clears its selection
            self.status = ("select_hint", {})
        self.render_status()

    def toggle_language(self):
        self.lang = "en" if self.lang == "ko" else "ko"
        self.save_settings()
        self.apply_language()

    def set_status(self, key, **params):
        self.status = (key, params)
        self.render_status()

    def render_status(self):
        key, params = self.status
        self.message.configure(text=self.t(key, **params))

    def issue_text(self, issue):
        return self.t("issue_" + issue.code, **issue.params)

    def where_text(self, wheres):
        return ", ".join(self.t("where_" + kind, n=n) for kind, n in wheres)

    def render_report(self):
        self.tree.delete(*self.tree.get_children())
        self.tree_items = {}
        report = self.report
        self.insight_button.configure(state="normal" if report and report.requests and not self.busy else "disabled")
        if report is None:
            self.summary.configure(text="")
            return
        for item in sorted(report.items, key=lambda i: STATUS_ORDER[i.status]):
            value = self.issue_text(item.issues[0]) if item.issues and (item.status == "error" or not item.value) else item.value
            label = item.key if item.kind == "text" else f"{self.t('kind_' + item.kind)}: {item.key}"
            row = self.tree.insert("", "end", tags=(item.status,), values=(
                self.t("status_" + item.status), label, value, item.source, self.where_text(item.wheres)))
            self.tree_items[row] = item
        self.summary.configure(text=self.t(
            "summary", ok=report.count("ok"), warn=report.count("warn"), error=report.count("error")))

    def show_detail(self):
        selected = self.tree.selection()
        item = self.tree_items.get(selected[0]) if selected else None
        if item is None:
            return
        lines = [self.issue_text(issue) for issue in item.issues]
        if not lines:
            lines = [f"{item.key} → {item.value}" + (f"  ({item.source})" if item.source else "")]
        self.status = ("detail", {})
        self.message.configure(text="\n".join(lines[:4]))

    def update_result_buttons(self):
        state = "normal" if self.last_output and os.path.exists(self.last_output) else "disabled"
        self.open_button.configure(state=state)
        self.folder_button.configure(state=state)

    # ------------------------------------------------------------ picking files

    def pick_template(self):
        path = filedialog.askopenfilename(
            title=self.t("pick_template"), initialdir=self.start_dir("template"),
            filetypes=[(self.t("filetype_pptx"), "*.pptx")])
        if path:
            old_dir = os.path.dirname(self.vars["template"].get())
            path = os.path.normpath(path)
            self.vars["template"].set(path)
            if not self.vars["output"].get() or os.path.normcase(self.vars["output"].get()) == os.path.normcase(old_dir):
                self.vars["output"].set(os.path.dirname(path))
            self.after_pick()

    def pick_excel(self):
        path = filedialog.askopenfilename(
            title=self.t("pick_excel"), initialdir=self.start_dir("excel"),
            filetypes=[(self.t("filetype_excel"), "*.xlsx *.xlsm")])
        if path:
            self.vars["excel"].set(os.path.normpath(path))
            self.after_pick()

    def pick_output(self):
        path = filedialog.askdirectory(title=self.t("pick_output"), initialdir=self.vars["output"].get() or None)
        if path:
            self.vars["output"].set(os.path.normpath(path))
            self.save_settings()

    def start_dir(self, key):
        return os.path.dirname(self.vars[key].get()) or None

    def after_pick(self):
        self.save_settings()
        if self.ready():
            self.preview()

    def ready(self):
        return bool(self.vars["template"].get().strip() and self.vars["excel"].get().strip())

    def save_settings(self):
        data = {key: self.vars[key].get().strip() for key in ("template", "excel", "output")}
        data["lang"] = self.lang
        try:
            core.save_settings(self.settings_file, data)
        except OSError:
            pass  # settings are a convenience; the app works without them

    # ------------------------------------------------------------ work

    def start(self, work, on_done, status_key):
        self.job_id += 1
        job = self.job_id
        self.set_busy(True)
        self.set_status(status_key)

        def target():
            try:
                self.jobs.put((job, on_done, work(), None))
            except core.SlideFillError as exc:
                self.jobs.put((job, on_done, None, ("err_" + exc.code, exc.params)))
            except Exception as exc:  # shown to the user rather than lost in a windowed app
                self.jobs.put((job, on_done, None, ("err_unexpected", {"detail": repr(exc)})))

        threading.Thread(target=target, daemon=True).start()

    def poll(self):
        try:
            while True:
                job, on_done, result, error = self.jobs.get_nowait()
                if job == self.job_id:
                    self.set_busy(False)
                    on_done(result, error)
        except queue.Empty:
            pass
        self.root.after(50, self.poll)

    def set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.generate_button.configure(state=state)
        self.preview_button.configure(state=state)
        usable = not busy and self.report is not None and self.report.requests
        self.insight_button.configure(state="normal" if usable else "disabled")

    def preview(self):
        if self.busy:
            return
        if not self.ready():
            self.set_status("missing_inputs")
            return
        template, excel = self.vars["template"].get().strip(), self.vars["excel"].get().strip()
        self.save_settings()
        insights = dict(self.insights)
        self.start(lambda: core.run(template, excel, insights=insights), self.preview_done, "working")

    def preview_done(self, report, error):
        if error:
            self.report = None
            self.render_report()
            self.show_error(error)
            self.set_status("start_hint")
            return
        self.report = report
        self.render_report()
        self.set_status("select_hint" if report.items else "no_placeholders")

    def generate(self):
        if self.busy:
            return
        if not self.ready():
            self.set_status("missing_inputs")
            return
        template, excel = self.vars["template"].get().strip(), self.vars["excel"].get().strip()
        folder = self.vars["output"].get().strip() or os.path.dirname(template)
        self.save_settings()
        insights = dict(self.insights)

        def work():
            preview = core.run(template, excel, insights=insights)
            problems = preview.count("error")
            if problems:
                return preview, None
            return preview, core.run(template, excel, core.output_path(folder, template), insights)

        self.start(work, lambda result, error: self.generate_checked(result, error, template, excel, folder),
                   "generating")

    def generate_checked(self, result, error, template, excel, folder):
        if error:
            self.show_error(error)
            self.set_status("start_hint")
            return
        preview, made = result
        if made is None:
            self.report = preview
            self.render_report()
            count = preview.count("error")
            if not messagebox.askyesno(self.t("confirm_title"), self.t("confirm_body", count=count)):
                self.set_status("select_hint")
                return
            insights = dict(self.insights)
            self.start(lambda: core.run(template, excel, core.output_path(folder, template), insights),
                       self.generate_done, "generating")
            return
        self.generate_done(made, None)

    def generate_done(self, report, error):
        if error:
            self.show_error(error)
            self.set_status("start_hint")
            return
        self.report = report
        self.last_output = report.output
        self.render_report()
        self.update_result_buttons()
        self.set_status("saved", path=report.output)

    def show_error(self, error):
        key, params = error
        messagebox.showerror(self.t("error_title"), self.t(key, **params))

    def open_path(self, path):
        if path and os.path.exists(path):
            os.startfile(path)

    # ------------------------------------------------------------ insight candidates

    def open_insight(self):
        if self.busy or not (self.report and self.report.requests):
            return
        if self.insight_window and self.insight_window.window.winfo_exists():
            self.insight_window.window.lift()
            return
        selected = self.tree.selection()
        item = self.tree_items.get(selected[0]) if selected else None
        request = self.report.requests.get(item.key) if item else None
        self.insight_window = InsightWindow(self, request or next(iter(self.report.requests.values())))

    def apply_insight(self, request, text):
        """Save the person's sentence for exactly these figures. Clearing it goes back to the calculated sentence."""
        key = request.key.casefold()
        if text:
            self.insights[key] = {"fingerprint": request.fingerprint, "text": text}
        else:
            self.insights.pop(key, None)
        try:
            core.save_settings(self.insights_file, self.insights)
        except OSError:
            pass  # the choice still holds until the app closes
        self.preview()

    def open_helper(self):
        """Works from the workbook alone: a template is what it helps to make."""
        excel = self.vars["excel"].get().strip()
        if not excel:
            self.set_status("helper_no_excel")
            return
        if self.helper_window and self.helper_window.window.winfo_exists():
            self.helper_window.window.lift()
            return
        self.helper_window = HelperWindow(self, excel)

    def use_template(self, path):
        """The starter template has just been made from this workbook, so a report can follow at once."""
        self.vars["template"].set(path)
        if not self.vars["output"].get():
            self.vars["output"].set(os.path.dirname(path))
        self.after_pick()

    def show_help(self):
        window = tk.Toplevel(self.root)
        window.title(self.t("help_title"))
        window.geometry(f"{self.px(720)}x{self.px(600)}")
        window.transient(self.root)
        body = ttk.Frame(window, padding=self.px(12))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        text = tk.Text(body, wrap="word", font=(FONT_FAMILY, 10), relief="flat", spacing1=self.px(2), spacing3=self.px(2))
        text.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(body, orient="vertical", command=text.yview)
        bar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=bar.set)
        text.insert("1.0", self.t("help_body"))
        text.configure(state="disabled")
        window.bind("<Escape>", lambda _e: window.destroy())
        window.bind("<F1>", lambda _e: window.destroy())
        window.focus_set()


class InsightWindow:
    """Shows the calculated results and the model's candidates. Nothing reaches the report until the person applies it."""

    def __init__(self, app, request):
        self.app, self.request = app, request
        self.results = queue.Queue()
        self.next_seed = 0
        self.working = False
        self.shown = []
        self.hypotheses = []
        window = self.window = tk.Toplevel(app.root)
        window.title(app.t("insight_title"))
        window.geometry(f"{app.px(820)}x{app.px(680)}")
        window.transient(app.root)
        body = ttk.Frame(window, padding=app.px(12))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=2)
        body.rowconfigure(7, weight=1)
        small = app.px(4)

        ttk.Label(body, text=f"{request.key}   ({request.source})", style="Summary.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text=app.t("insight_warning"), style="Muted.TLabel", wraplength=app.px(780), justify="left").grid(
            row=1, column=0, sticky="w", pady=(small, small))
        ttk.Label(body, text=app.t("insight_candidates")).grid(row=2, column=0, sticky="w")
        self.listbox = tk.Listbox(body, font=(FONT_FAMILY, 10), activestyle="none", exportselection=False)
        self.listbox.grid(row=3, column=0, sticky="nsew", pady=(small, small))
        self.listbox.bind("<<ListboxSelect>>", lambda _e: self.show_basis())
        self.listbox.bind("<Double-Button-1>", lambda _e: self.add())

        holder = ttk.Frame(body)
        holder.grid(row=4, column=0, sticky="ew")
        holder.columnconfigure(0, weight=1)
        self.basis = tk.Text(holder, height=8, wrap="word", font=(FONT_FAMILY, 9), relief="flat", state="disabled", background=window.cget("background"))
        self.basis.grid(row=0, column=0, sticky="ew")
        bar = ttk.Scrollbar(holder, orient="vertical", command=self.basis.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.basis.configure(yscrollcommand=bar.set)

        row = ttk.Frame(body)
        row.grid(row=5, column=0, sticky="ew", pady=(small, small))
        self.add_button = ttk.Button(row, text=app.t("insight_add"), command=self.add, state="disabled")
        self.add_button.pack(side="left")
        self.again_button = ttk.Button(row, text=app.t("insight_again"), command=self.make)
        self.again_button.pack(side="left", padx=small)

        ttk.Label(body, text=app.t("insight_chosen")).grid(row=6, column=0, sticky="w")
        self.chosen = tk.Text(body, height=4, wrap="word", font=(FONT_FAMILY, 10))
        self.chosen.grid(row=7, column=0, sticky="nsew", pady=(small, small))
        saved = app.insights.get(request.key.casefold())
        if saved and saved.get("fingerprint") == request.fingerprint:
            self.chosen.insert("1.0", saved.get("text", ""))

        self.status = ttk.Label(body, wraplength=app.px(780), justify="left")
        self.status.grid(row=8, column=0, sticky="w")

        foot = ttk.Frame(body)
        foot.grid(row=9, column=0, sticky="e", pady=(small, 0))
        self.apply_button = ttk.Button(foot, text=app.t("insight_apply"), style="Accent.TButton", command=self.apply)
        self.apply_button.pack(side="left", padx=small)
        ttk.Button(foot, text=app.t("insight_close"), command=window.destroy).pack(side="left")
        window.bind("<Escape>", lambda _e: window.destroy())
        window.after(50, self.poll)
        self.show_facts()
        self.make()

    def show_facts(self):
        self.set_basis("\n".join(fact.text for fact in self.request.facts))

    def set_basis(self, text):
        self.basis.configure(state="normal")
        self.basis.delete("1.0", "end")
        self.basis.insert("1.0", text)
        self.basis.configure(state="disabled")

    def make(self):
        if self.working:
            return
        self.working = True
        self.again_button.configure(state="disabled")
        self.status.configure(text=self.app.t("insight_working", model=insight.MODELS[self.request.lang]))
        seed, self.next_seed = self.next_seed, self.next_seed + insight.RUNS
        request, results = self.request, self.results

        def work():
            try:
                results.put(("done", insight.make_candidates(request.facts, request.lang, first_seed=seed)))
            except insight.InsightError as exc:
                results.put(("error", exc))
            except Exception as exc:  # a thread has no one else to tell
                results.put(("error", insight.InsightError("other", detail=repr(exc))))

        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        if not self.window.winfo_exists():
            return
        try:
            kind, value = self.results.get_nowait()
        except queue.Empty:
            self.window.after(100, self.poll)
            return
        self.working = False
        self.again_button.configure(state="normal")
        if kind == "error":
            self.status.configure(text=self.app.t("insight_err_" + value.code, **value.params))
        else:
            self.show_candidates(value)
        self.window.after(100, self.poll)

    def show_candidates(self, result):
        known = {c.statement for c in self.shown}
        for candidate in result.candidates:
            if candidate.statement not in known:
                self.shown.append(candidate)
                self.listbox.insert("end", candidate.statement)
        for hypothesis in result.hypotheses:
            if hypothesis not in self.hypotheses:
                self.hypotheses.append(hypothesis)
        key = "insight_done" if self.shown else "insight_none"
        self.status.configure(text=self.app.t(key, count=len(self.shown), rejected=result.rejected, model=result.model))
        if self.hypotheses:
            self.set_basis(self.app.t("insight_hypotheses") + "\n" + "\n".join(self.hypotheses))

    def show_basis(self):
        picked = self.listbox.curselection()
        self.add_button.configure(state="normal" if picked else "disabled")
        if not picked:
            return
        known = {fact.id: fact.text for fact in self.request.facts}
        candidate = self.shown[picked[0]]
        cited = [known[name] for name in candidate.fact_ids if name in known]
        lines = [candidate.statement, "", self.app.t("insight_basis")] + cited  # the list cuts off a long sentence
        if self.hypotheses:
            lines +=["", self.app.t("insight_hypotheses")] + self.hypotheses
        self.set_basis("\n".join(lines))

    def add(self):
        picked = self.listbox.curselection()
        if not picked:
            return
        current = self.chosen.get("1.0", "end").strip()
        self.chosen.delete("1.0", "end")
        self.chosen.insert("1.0", (current + " " if current else "") + self.shown[picked[0]].statement)

    def apply(self):
        self.app.apply_insight(self.request, self.chosen.get("1.0", "end").strip())
        self.window.destroy()


class HelperWindow:
    """Offers references that work on the workbook, each with its real result. Nobody has to learn how to write them."""

    HOWS = ("plain", "sum", "mean", "count", "rate")

    def __init__(self, app, excel):
        self.app, self.excel = app, excel
        self.source = None
        self.shape = None
        self.found = []
        self.results = queue.Queue()
        self.job = 0
        self.working = False
        self.sheets = []
        self.bases = []
        window = self.window = tk.Toplevel(app.root)
        window.title(app.t("helper_title"))
        window.geometry(f"{app.px(960)}x{app.px(640)}")
        window.transient(app.root)
        body = ttk.Frame(window, padding=app.px(12))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)
        small = app.px(4)

        ttk.Label(body, text=app.t("helper_intro"), style="Muted.TLabel", wraplength=app.px(920), justify="left").grid(row=0, column=0, sticky="w")

        picks = ttk.Frame(body)
        picks.grid(row=1, column=0, sticky="ew", pady=(app.px(8), small))
        for column in (1, 3, 5):
            picks.columnconfigure(column, weight=1)
        self.boxes = {}
        layout = (("sheet", 0, 0), ("base", 0, 2), ("value", 0, 4), ("how", 1, 0), ("group1", 1, 2), ("word1", 1, 4), ("group2", 2, 2), ("word2", 2, 4))
        for name, row, column in layout:
            ttk.Label(picks, text=app.t("helper_" + name)).grid(row=row, column=column, sticky="w", padx=(0 if column == 0 else app.px(10), small), pady=2)
            box = ttk.Combobox(picks, state="readonly", width=24)
            box.grid(row=row, column=column + 1, sticky="ew", pady=2)
            box.bind("<<ComboboxSelected>>", lambda _e, n=name: self.changed(n))
            self.boxes[name] = box

        self.shape_line = ttk.Label(body, style="Summary.TLabel", wraplength=app.px(920), justify="left")
        self.shape_line.grid(row=2, column=0, sticky="w", pady=(small, small))

        table = ttk.Frame(body)
        table.grid(row=3, column=0, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        columns = ("what", "result", "place")
        self.tree = ttk.Treeview(table, columns=columns, show="headings", selectmode="browse")
        for column, width in (("what", 180), ("result", 520), ("place", 180)):
            self.tree.heading(column, text=app.t("helper_col_" + column), anchor="w")
            self.tree.column(column, width=app.px(width), stretch=column == "result", anchor="w")
        for status, color in STATUS_COLOR.items():
            self.tree.tag_configure(status, foreground=color if status != "ok" else "")
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.show_row())
        self.tree.bind("<Double-Button-1>", lambda _e: self.copy())

        self.key_var = tk.StringVar()
        ttk.Entry(body, textvariable=self.key_var, state="readonly").grid(row=4, column=0, sticky="ew", pady=(app.px(8), small))
        self.detail = ttk.Label(body, wraplength=app.px(920), justify="left")
        self.detail.grid(row=5, column=0, sticky="w")

        foot = ttk.Frame(body)
        foot.grid(row=6, column=0, sticky="e", pady=(small, 0))
        self.copy_button = ttk.Button(foot, text=app.t("helper_copy"), command=self.copy, state="disabled")
        self.copy_button.pack(side="left")
        self.starter_button = ttk.Button(foot, text=app.t("helper_starter"), style="Accent.TButton", command=self.make_starter, state="disabled")
        self.starter_button.pack(side="left", padx=small)
        ttk.Button(foot, text=app.t("insight_close"), command=window.destroy).pack(side="left")
        window.bind("<Escape>", lambda _e: window.destroy())
        window.bind("<Control-c>", lambda _e: self.copy())
        window.after(50, self.poll)
        self.run(self.load, self.loaded)

    # ---- background work: a large workbook takes a while, and the window should stay alive

    def run(self, work, done):
        self.job += 1
        job = self.job
        self.working = True
        self.detail.configure(text=self.app.t("helper_working"))
        for box in self.boxes.values():
            box.configure(state="disabled")
        results = self.results

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
            self.window.after(100, self.poll)
            return
        if job == self.job:
            self.working = False
            for box in self.boxes.values():
                box.configure(state="readonly")
            if error:
                self.detail.configure(text=error)
            else:
                done(result)
        self.window.after(100, self.poll)

    # ---- reading

    def load(self):
        self.source = core.ExcelSource(self.excel)
        sheets = suggest.data_sheets(self.source)
        first, base, found = suggest.first_choice(self.source, lang=self.app.lang)
        return sheets, first, base, found

    def loaded(self, result):
        sheets, shape, base, found = result
        self.sheets = [ws.title for ws in sheets]
        self.boxes["sheet"]["values"] = self.sheets
        if shape is None:
            self.detail.configure(text=self.app.t("helper_nothing"))
            return
        self.boxes["sheet"].set(shape.sheet)
        self.show_shape(shape, base, found)

    def show_shape(self, shape, base, found):
        """Fill the choices from what was read, with the same defaults the first suggestions used."""
        self.shape = shape
        self.bases = shape.bases
        self.boxes["base"]["values"] = [self.base_text(b) for b in self.bases]
        self.boxes["base"].current(self.bases.index(base))
        self.fill_base(base)
        self.show_found(found)

    def base_text(self, base):
        return f"{base.spec} · {base.name or self.app.t('helper_no_name')}"

    def value_text(self, value):
        if value.every_row:
            return self.app.t("helper_every_row")
        if value.counted:
            return self.app.t("helper_counted")
        return f"{value.spec} · {value.name}"

    def fill_base(self, base):
        value = suggest.default_value(base)
        self.boxes["value"]["values"] = [self.value_text(v) for v in base.values]
        if value is not None:
            self.boxes["value"].current(base.values.index(value))
        how = suggest.default_how(base, value)
        self.boxes["how"]["values"] = [self.app.t("helper_how_" + name) for name in self.HOWS]
        self.boxes["how"].current(self.HOWS.index(how))
        none = self.app.t("helper_none")
        names = [none] + [f"{g.spec} · {g.name}" for g in base.groups]
        for number, group_box in (("1", "group1"), ("2", "group2")):
            self.boxes[group_box]["values"] = names
            self.boxes[group_box].current(0)
            self.boxes["word" + number]["values"] = []
            self.boxes["word" + number].set("")
        for number, (spec, text) in enumerate(suggest.default_conditions(self.source, self.shape, base, value, how)[:2], start=1):
            at = next(i for i, g in enumerate(base.groups) if g.spec == spec)
            self.boxes[f"group{number}"].current(at + 1)
            self.boxes[f"word{number}"]["values"] = base.groups[at].values
            self.boxes[f"word{number}"].set(text)
        line = self.app.t("helper_shape_dates" if base.kind == "date" else "helper_shape_names", sheet=self.shape.sheet, spec=base.spec, first=base.first, last=base.last, count=base.count)
        if self.shape.glob and base.kind == "date":
            line += " " + self.app.t("helper_shape_joined", pattern=self.shape.glob)
        for code, params in base.notes:
            line += " " + self.app.t("helper_" + code, **params)
        self.shape_line.configure(text=line)

    def base(self):
        return self.bases[self.boxes["base"].current()]

    def conditions(self):
        base, chosen = self.base(), []
        for number in ("1", "2"):
            at = self.boxes["group" + number].current()
            text = self.boxes["word" + number].get()
            if at > 0 and text:
                chosen.append((base.groups[at - 1].spec, text))
        return chosen

    def changed(self, name):
        if self.working:
            return
        if name == "sheet":
            title = self.boxes["sheet"].get()
            self.run(lambda: self.read_sheet(title), lambda result: self.show_shape(*result))
            return
        if name == "base":
            self.fill_base(self.base())
        if name in ("group1", "group2"):
            at = self.boxes[name].current()
            number = name[-1]
            base = self.base()
            self.boxes["word" + number]["values"] = base.groups[at - 1].values if at > 0 else []
            self.boxes["word" + number].set("")
            if at > 0:
                return  # wait for the value, which is what makes the condition
        self.recount()

    def read_sheet(self, title):
        shape = suggest.read_shape(self.source, title)
        if not shape.bases:
            raise core.SlideFillError("helper_empty_sheet", sheet=title)
        base = shape.bases[0]
        return shape, base, suggest.suggestions(self.source, shape, base, lang=self.app.lang)

    def recount(self):
        base, shape = self.base(), self.shape
        if not base.values:
            return
        value = base.values[self.boxes["value"].current()]
        how = self.HOWS[self.boxes["how"].current()]
        conditions = self.conditions()
        if how == "rate" and not conditions:
            self.detail.configure(text=self.app.t("helper_rate_needs_condition"))
            return
        lang = self.app.lang
        self.run(lambda: suggest.suggestions(self.source, shape, base, value, how, conditions, lang), self.show_found)

    # ---- showing

    def show_found(self, found):
        self.found = found
        self.tree.delete(*self.tree.get_children())
        for number, item in enumerate(found):
            result = item.result if item.status != "error" else self.app.issue_text(item.issues[0])
            place = self.app.t("helper_place_" + item.place)
            self.tree.insert("", "end", iid=str(number), tags=(item.status,), values=(self.app.t("suggest_" + item.code), result, place))
        works = any(item.status != "error" for item in found)
        self.starter_button.configure(state="normal" if works else "disabled")
        self.copy_button.configure(state="disabled")
        self.key_var.set("")
        self.detail.configure(text=self.app.t("helper_pick_row") if works else self.app.t("helper_nothing"))
        if found and works:
            first = next(i for i, item in enumerate(found) if item.status != "error")
            self.tree.selection_set(str(first))

    def picked(self):
        chosen = self.tree.selection()
        return self.found[int(chosen[0])] if chosen else None

    def show_row(self):
        item = self.picked()
        if item is None:
            return
        self.key_var.set(item.typed)
        self.copy_button.configure(state="normal" if item.status != "error" else "disabled")
        lines = [self.app.t("helper_where_" + item.place)]
        if item.place == "text":
            lines.append(item.result)
        lines += [self.app.issue_text(issue) for issue in item.issues[:3]]
        self.detail.configure(text="\n".join(lines))

    def copy(self):
        item = self.picked()
        if item is None or item.status == "error":
            return
        self.window.clipboard_clear()
        self.window.clipboard_append(item.typed)
        self.window.update()
        self.detail.configure(text=self.app.t("helper_copied", reference=item.typed) + "\n" + self.app.t("helper_where_" + item.place))

    def make_starter(self):
        folder = self.app.vars["output"].get().strip() or os.path.dirname(self.excel)
        path = suggest.starter_path(folder, self.excel)
        title = os.path.splitext(os.path.basename(self.excel))[0]
        try:
            suggest.make_starter(path, title, self.found)
        except OSError as exc:
            self.detail.configure(text=self.app.t("err_cannot_save", path=path, detail=str(exc)))
            return
        self.app.use_template(path)
        self.detail.configure(text=self.app.t("helper_starter_saved", path=path))


def smoke_insight(template, excel, output):
    """Check for the frozen exe with a real local model: the second stage runs and leaves candidates in a JSON file."""
    try:
        report = core.run(template, excel)
        results = {}
        for request in report.requests.values():
            made = insight.make_candidates(request.facts, request.lang)
            results[request.key] = {"lang": request.lang, "model": made.model, "rejected": made.rejected, "hypotheses": made.hypotheses,
                                    "candidates": [c.statement for c in made.candidates], "fallback": request.fallback}
        core.save_settings(output, results)
        return 0 if results and all(r["candidates"] for r in results.values()) else 1
    except Exception:
        with open(output + ".log", "w", encoding="utf-8") as fh:
            fh.write(traceback.format_exc())
        return 1


def smoke(template, excel, output):
    """Build check for the frozen exe: GUI toolkit starts, and a real fill and save work. Exit code 0 means fine."""
    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        report = core.run(template, excel, output)
        problems = [(i.key, i.status, [x.code for x in i.issues]) for i in report.items if i.status != "ok"]
        if os.path.exists(output) and report.items and not problems:
            return 0
        detail = f"items={len(report.items)} problems={problems}"
    except Exception:
        detail = traceback.format_exc()
    with open(output + ".log", "w", encoding="utf-8") as fh:  # a windowed exe has no console to print to
        fh.write(detail)
    return 1


def main():
    if len(sys.argv) == 5 and sys.argv[1] == "--smoke":
        sys.exit(smoke(*sys.argv[2:5]))
    if len(sys.argv) == 5 and sys.argv[1] == "--smoke-insight":
        sys.exit(smoke_insight(*sys.argv[2:5]))
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    root = tk.Tk()
    try:
        root.iconbitmap(resource("app.ico"))
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
