"""SlideFill window. UI only: all file work happens in core.py, on a worker thread."""

import atexit
import ctypes
import os
import queue
import shutil
import sys
import tempfile
import threading
import tkinter as tk
import tkinter.font as tkfont
import traceback
from dataclasses import dataclass, field
from tkinter import filedialog, messagebox, ttk

import contents
import core
import i18n
import insight
import readers
import suggest

STATUS_ORDER = {"error": 0, "warn": 1, "ok": 2}
STATUS_COLOR = {"error": "#B3261E", "warn": "#9A6700", "ok": "#1A7F37"}
FONT_FAMILY = "Malgun Gothic"
SETTINGS_KEYS = ("template", "excel", "output", "lang")
DOUBTFUL = object()  # what a run returns in place of a report when a person should look first
FILE_PATTERN = " ".join("*" + ext for ext in readers.ALL_TYPES + readers.DOCUMENT_TYPES)


def is_document(path):
    return os.path.splitext(path)[1].lower() in readers.DOCUMENT_TYPES


def is_one_workbook(files):
    return len(files) == 1 and not is_document(files[0])


def first_workbook(files):
    return next((path for path in files if not is_document(path)), files[0])


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


@dataclass
class AutoReport:
    """The report SlideFill laid out itself, for when the person brought only workbooks."""
    excel: tuple  # every file it reads
    stamp: float
    lang: str
    path: str  # a working file in a temporary folder; the person only ever sees the report made from it
    what: str
    labels: dict = field(default_factory=dict)
    chosen: bool = False  # the person picked these contents, so they stay until another workbook is chosen
    figures_left: int = 0
    sheets_left: int = 0
    title: str = ""
    sections: list = field(default_factory=list)
    asks: dict = field(default_factory=dict)  # what the model can be asked, by reference: (request, figure)
    stale: bool = True  # the working file needs making again
    files_left: list = field(default_factory=list)
    doubts: list = field(default_factory=list)  # why a person should look at what was chosen, until they have


class DoubtWindow:
    """Says in plain words why the choice may be wrong and shows what it gave, then leaves the decision to the person. Nothing has been made yet."""

    def __init__(self, app, auto):
        self.app = app
        window = self.window = tk.Toplevel(app.root)
        window.title(app.t("doubt_title"))
        window.transient(app.root)
        body = ttk.Frame(window, padding=app.px(16))
        body.pack(fill="both", expand=True)
        width = app.px(620)
        ttk.Label(body, text=app.t("doubt_intro"), wraplength=width, justify="left").pack(anchor="w")
        ttk.Label(body, text=app.t("doubt_chosen", what=auto.what), wraplength=width, justify="left", style="Summary.TLabel").pack(anchor="w", pady=(app.px(10), 0))
        self.reasons = []
        for code, params in auto.doubts:
            words = app.t("doubt_" + code, figure=params.get("figure", ""))
            if params.get("file"):
                words = f"{params['file']}: {words}"
            self.reasons.append(words)
            ttk.Label(body, text="• " + words, wraplength=width, justify="left").pack(anchor="w", pady=(app.px(6), 0))
        foot = ttk.Frame(body)
        foot.pack(anchor="e", pady=(app.px(16), 0))
        self.change_button = ttk.Button(foot, text=app.t("doubt_change"), style="Accent.TButton", command=lambda: self.answer("change"))
        self.change_button.pack(side="left")
        self.keep_button = ttk.Button(foot, text=app.t("doubt_keep"), command=lambda: self.answer("keep"))
        self.keep_button.pack(side="left", padx=(app.px(4), 0))
        self.cancel_button = ttk.Button(foot, text=app.t("doubt_cancel"), command=lambda: self.answer(None))
        self.cancel_button.pack(side="left", padx=(app.px(4), 0))
        window.bind("<Escape>", lambda _e: self.answer(None))
        window.protocol("WM_DELETE_WINDOW", lambda: self.answer(None))
        window.update_idletasks()
        window.geometry(f"+{app.root.winfo_rootx() + app.px(40)}+{app.root.winfo_rooty() + app.px(40)}")
        window.grab_set()
        self.change_button.focus_set()

    def answer(self, what):
        self.window.grab_release()
        self.window.destroy()
        self.app.doubts_answered(what)


class App:
    def __init__(self, root):
        self.root = root
        self.settings_file = core.settings_path(app_dir())
        self.insights_file = core.settings_path(app_dir(), core.INSIGHTS_FILE)
        self.insights = core.load_settings(self.insights_file)
        self.insight_window = None
        self.helper_window = None
        self.doubt_window = None
        saved = core.load_settings(self.settings_file)
        self.lang = saved.get("lang") if saved.get("lang") in i18n.STRINGS else system_language()
        self.vars = {key: tk.StringVar(value=str(saved.get(key) or "")) for key in ("template", "excel", "output")}
        self.more_files = [str(path) for path in saved.get("more_excels") or []]
        self.own_template = tk.BooleanVar(value=bool(saved.get("use_template", saved.get("template"))))
        self.auto = None
        self.work_dir = None
        self.jobs = queue.Queue()
        self.job_id = 0
        self.busy = False
        self.again = False
        self.report = None
        self.status = ("start_hint", {})
        self.last_output = ""
        self.tree_items = {}

        self.scale = root.winfo_fpixels("1i") / 96
        self.build()
        self.apply_language()
        self.bind_keys()
        self.vars["excel"].trace_add("write", lambda *_a: self.show_files())
        self.show_files()
        self.drop = contents.DropTarget(root, self.add_files)
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
        self.file_labels["excel"] = ttk.Label(files)
        self.file_labels["excel"].grid(row=0, column=0, sticky="nw", padx=(0, self.px(10)), pady=self.px(3))
        holder = ttk.Frame(files)
        holder.grid(row=0, column=1, sticky="ew", pady=self.px(3))
        holder.columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(holder, height=3, activestyle="none", exportselection=False, font=(FONT_FAMILY, 10), selectmode="extended")
        self.file_list.grid(row=0, column=0, sticky="ew")
        self.file_list.bind("<<ListboxSelect>>", lambda _e: self.show_files_state())
        self.file_list.bind("<Delete>", lambda _e: self.remove_files())
        self.drop_hint = ttk.Label(self.file_list, style="Muted.TLabel", background="white", justify="center")
        side = ttk.Frame(files)
        side.grid(row=0, column=2, sticky="n", padx=(self.px(6), 0), pady=self.px(3))
        self.browse_buttons["excel"] = ttk.Button(side, command=self.pick_excel)
        self.browse_buttons["excel"].pack(fill="x")
        self.remove_button = ttk.Button(side, command=self.remove_files, state="disabled")
        self.remove_button.pack(fill="x", pady=(self.px(3), 0))
        commands = {"template": self.pick_template, "output": self.pick_output}
        for row, key in ((1, "output"), (3, "template")):
            label = ttk.Label(files)
            label.grid(row=row, column=0, sticky="w", padx=(0, self.px(10)), pady=self.px(3))
            entry = ttk.Entry(files, textvariable=self.vars[key])
            entry.grid(row=row, column=1, sticky="ew", pady=self.px(3))
            entry.bind("<Return>", lambda _e: self.preview())
            entry.bind("<FocusOut>", lambda _e: self.save_settings())
            button = ttk.Button(files, command=commands[key])
            button.grid(row=row, column=2, padx=(self.px(6), 0), pady=self.px(3))
            self.file_labels[key], self.browse_buttons[key] = label, button
            if key == "template":
                self.template_row = (label, entry, button)
        self.template_check = ttk.Checkbutton(files, variable=self.own_template, command=self.toggle_template)
        self.template_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=self.px(3))

        actions = ttk.Frame(main)
        actions.grid(row=2, column=0, sticky="ew", pady=self.px(12))
        actions.columnconfigure(4, weight=1)
        self.generate_button = ttk.Button(actions, style="Accent.TButton", command=self.generate)
        self.generate_button.grid(row=0, column=0, ipady=self.px(3))
        self.preview_button = ttk.Button(actions, command=self.preview)
        self.preview_button.grid(row=0, column=1, padx=(self.px(6), 0), ipady=self.px(3))
        self.insight_button = ttk.Button(actions, command=self.open_insight, state="disabled")
        self.insight_button.grid(row=0, column=2, padx=(self.px(6), 0), ipady=self.px(3))
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
            self.browse_buttons[key].configure(text=self.t("add_files" if key == "excel" else "browse"))
        self.remove_button.configure(text=self.t("remove_files"))
        self.drop_hint.configure(text=self.t("drop_hint"))
        self.template_check.configure(text=self.t("own_template"))
        self.generate_button.configure(text=self.t("generate"))
        self.preview_button.configure(text=self.t("preview"))
        self.insight_button.configure(text=self.t("insight_button"))
        self.show_template_row()
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
        if self.report and self.auto and not self.auto.chosen and not self.template_path():
            self.preview()  # the sentences SlideFill laid out are in the language they were made in

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
        self.insight_button.configure(state="normal" if self.insight_requests() and not self.busy else "disabled")
        if report is None:
            self.summary.configure(text="")
            return
        order = {key: place for place, key in enumerate(self.auto.labels)} if self.auto and not self.template_path() else {}  # as the report reads
        for item in sorted(report.items, key=lambda i: (STATUS_ORDER[i.status], order.get(i.key.casefold(), 0))):
            value = self.issue_text(item.issues[0]) if item.issues and (item.status == "error" or not item.value) else item.value
            label, source = self.plain_words(item)
            row = self.tree.insert("", "end", tags=(item.status,), values=(
                self.t("status_" + item.status), label, value, source, self.where_text(item.wheres)))
            self.tree_items[row] = item
        self.summary.configure(text=self.t(
            "summary", ok=report.count("ok"), warn=report.count("warn"), error=report.count("error")))

    def plain_words(self, item):
        """(what it is, where it comes from). Someone with a template wrote the notation; someone with only a workbook should never meet it."""
        named = self.auto.labels.get(item.key.casefold()) if self.auto and not self.template_path() else None
        if named:
            return " · ".join(part for part in (named[1], self.t("suggest_" + named[0])) if part), item.source.rpartition("!")[0] or item.source
        return (item.key if item.kind == "text" else f"{self.t('kind_' + item.kind)}: {item.key}"), item.source

    def show_detail(self):
        selected = self.tree.selection()
        item = self.tree_items.get(selected[0]) if selected else None
        if item is None:
            return
        lines = [self.issue_text(issue) for issue in item.issues]
        if not lines:
            label, source = self.plain_words(item)
            lines = [f"{label} → {item.value}" + (f"  ({source})" if source else "")]
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
            self.own_template.set(True)
            self.show_template_row()
            self.after_pick()

    def pick_excel(self):
        paths = filedialog.askopenfilenames(
            title=self.t("pick_excel"), initialdir=self.start_dir("excel"),
            filetypes=[(self.t("filetype_excel"), FILE_PATTERN)])
        if paths:
            self.add_files(paths)

    def add_files(self, paths):
        """From the file dialog or from a drop. A file of another kind is said out loud, not ignored."""
        known = {os.path.normcase(path) for path in self.inputs()}
        added, refused = 0, []
        for path in paths:
            path = os.path.normpath(path)
            if os.path.splitext(path)[1].lower() not in readers.ALL_TYPES + readers.DOCUMENT_TYPES or os.path.isdir(path):
                refused.append(os.path.basename(path))
            elif os.path.normcase(path) not in known:
                known.add(os.path.normcase(path))
                added += 1
                if self.vars["excel"].get().strip():
                    self.more_files.append(path)
                else:
                    self.vars["excel"].set(path)
                if not self.vars["output"].get():
                    self.vars["output"].set(os.path.dirname(path))
        self.show_files()
        if added:
            self.after_pick()
        if refused:
            self.show_refused(refused, added)

    def show_refused(self, refused, added):
        if added:
            messagebox.showinfo(self.t("error_title"), self.t("drop_refused", names=", ".join(refused)))
        else:
            self.set_status("drop_refused", names=", ".join(refused))

    def remove_files(self):
        chosen = set(self.file_list.curselection())
        kept = [path for place, path in enumerate(self.inputs()) if place not in chosen]
        self.more_files = kept[1:]
        self.vars["excel"].set(kept[0] if kept else "")
        self.show_files()
        self.save_settings()
        if kept:
            self.preview()
        else:
            self.report = None
            self.auto = None
            self.render_report()
            self.set_status("start_hint")

    def inputs(self):
        first = self.vars["excel"].get().strip()
        return [first] + self.more_files if first else list(self.more_files)

    def show_files(self):
        self.file_list.delete(0, "end")
        for path in self.inputs():
            self.file_list.insert("end", f"{os.path.basename(path)}    ({os.path.dirname(path)})")
        if self.inputs():
            self.drop_hint.place_forget()
        else:
            self.drop_hint.place(relx=0.5, rely=0.5, anchor="center")
        self.show_files_state()

    def show_files_state(self):
        self.remove_button.configure(state="normal" if self.file_list.curselection() else "disabled")

    def toggle_template(self):
        if self.own_template.get() and not self.vars["template"].get().strip():
            self.pick_template()
            if not self.vars["template"].get().strip():  # the person changed their mind in the file dialog
                self.own_template.set(False)
            return
        self.show_template_row()
        self.after_pick()

    def show_template_row(self):
        """A template is for those who have one. Everyone else should not have to wonder about it."""
        for widget in self.template_row:
            if self.own_template.get():
                widget.grid()
            else:
                widget.grid_remove()
        if self.own_template.get():  # with no template, choosing a workbook already previews it
            self.preview_button.grid()
        else:
            self.preview_button.grid_remove()
        self.helper_button.configure(text=self.t("helper_button" if self.own_template.get() else "helper_button_auto"))

    def template_path(self):
        """The person's own template, or nothing when SlideFill is to lay out the report itself."""
        return self.vars["template"].get().strip() if self.own_template.get() else ""

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
        return bool(self.inputs())

    def save_settings(self):
        data = {key: self.vars[key].get().strip() for key in ("template", "excel", "output")}
        data["more_excels"] = self.more_files
        data["lang"] = self.lang
        data["use_template"] = self.own_template.get()
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
        if self.again and not self.busy:  # a file was chosen while another was still being read
            self.again = False
            self.preview()
        self.root.after(50, self.poll)

    def set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.generate_button.configure(state=state)
        self.preview_button.configure(state=state)
        self.insight_button.configure(state="normal" if not busy and self.insight_requests() else "disabled")

    @staticmethod
    def read_files(files):
        return core.ExcelSource(files[0]) if is_one_workbook(files) else core.Sources(files)

    def auto_template(self, files, lang):
        """(working template, workbooks already read or None). Runs on the worker: choosing what to show means reading them."""
        known = self.auto
        files = tuple(files)
        stamp = max((os.path.getmtime(path) for path in files if os.path.exists(path)), default=0)
        if known and known.excel == files and (known.chosen or (known.stamp, known.lang) == (stamp, lang)):
            if not known.stale:
                return known.path, None
            source = self.read_files(files)
            self.make_auto(known, source)
            return known.path, source
        source = self.read_files(files)
        plan = suggest.plan_report(source, lang) if is_one_workbook(files) else suggest.plan_files(source, lang)
        if plan is None:
            raise core.SlideFillError("auto_nothing")
        self.auto = AutoReport(files, stamp, lang, self.auto_path(files, lang), plan.what, figures_left=plan.figures_left, sheets_left=plan.sheets_left,
                               title=self.report_name(files, plan.sections, lang), sections=plan.sections, files_left=plan.files_left, doubts=plan.doubts)
        self.make_auto(self.auto, source)
        return self.auto.path, source

    @staticmethod
    def report_name(files, sections, lang):
        first = suggest.title_of(first_workbook(files), sections)
        return first if len(files) == 1 else i18n.text(lang, "report_many", first=first, more=len(files) - 1)

    def make_auto(self, auto, source):
        """The working file, with any sentence the person chose in place of the calculated one."""
        auto.asks = suggest.ask_insights(source, auto.sections, auto.title, auto.lang)
        sections = suggest.choose_insights(auto.sections, auto.asks, dict(self.insights))
        suggest.make_report(auto.path, auto.title, sections)
        auto.labels = suggest.labels(item for section in sections for item in section.items)
        auto.stale = False

    def auto_path(self, files, lang):
        if self.work_dir is None:
            self.work_dir = tempfile.mkdtemp(prefix="SlideFill-")
            atexit.register(shutil.rmtree, self.work_dir, ignore_errors=True)
        stem = self.report_title(first_workbook(files))
        if len(files) > 1:
            stem = i18n.text(lang, "report_many_file", first=stem, more=len(files) - 1)
        return os.path.join(self.work_dir, stem + ".pptx")  # the report is named after this file

    @staticmethod
    def report_title(excel):
        return os.path.splitext(os.path.basename(excel))[0]

    def use_contents(self, excel, found, what, words=""):
        """The person chose what the report shows, in the helper."""
        sections = [suggest.one_section(found, words)]
        files = (excel,)
        self.auto = AutoReport(files, 0, self.lang, self.auto_path(files, self.lang), what, chosen=True, title=suggest.title_of(excel, sections), sections=sections)
        self.preview()

    def use_sections(self, files, sections, what):
        """The person chose what the report shows, in the contents window."""
        files = tuple(files)
        real = [section for section in sections if section.file >= 0]
        self.auto = AutoReport(files, 0, self.lang, self.auto_path(files, self.lang), what, chosen=True, title=self.report_name(files, real, self.lang), sections=sections)
        self.preview()

    def preview(self):
        if self.busy:
            self.again = True
            return
        if not self.ready():
            self.set_status("missing_inputs")
            return
        template, files, lang = self.template_path(), self.inputs(), self.lang
        excel = first_workbook(files) if template or is_one_workbook(files) else files
        self.save_settings()
        insights = dict(self.insights)

        def work():
            path, source = (template, None) if template else self.auto_template(files, lang)
            return core.run(path, excel, insights=insights, source=source, thousands=not template)

        self.start(work, self.preview_done, "working" if template else "helper_working")

    def preview_done(self, report, error):
        if error:
            self.report = None
            self.render_report()
            self.show_error(error)
            self.set_status("start_hint")
            return
        self.report = report
        self.render_report()
        if not self.template_path():
            auto = self.auto
            if auto.doubts:
                self.set_status("auto_ready_doubt", what=auto.what, count=len(auto.doubts))
            elif auto.files_left:
                self.set_status("auto_ready_unread", what=auto.what, names=", ".join(auto.files_left))
            elif auto.figures_left or auto.sheets_left:
                self.set_status("auto_ready_left", what=auto.what, figures=auto.figures_left, sheets=auto.sheets_left)
            else:
                self.set_status("auto_ready", what=auto.what)
            return
        if len(self.inputs()) > 1:
            self.set_status("template_first_file", name=os.path.basename(self.inputs()[0]))
            return
        self.set_status("select_hint" if report.items else "no_placeholders")

    def generate(self):
        if self.busy:
            return
        if not self.ready():
            self.set_status("missing_inputs")
            return
        template, files, lang = self.template_path(), self.inputs(), self.lang
        excel = first_workbook(files) if template or is_one_workbook(files) else files
        folder = self.vars["output"].get().strip() or os.path.dirname(template or files[0])
        self.save_settings()
        insights = dict(self.insights)

        def work():
            path, source = (template, None) if template else self.auto_template(files, lang)
            preview = core.run(path, excel, insights=insights, source=source, thousands=not template)
            if not template and self.auto.doubts:
                return path, preview, DOUBTFUL
            problems = preview.count("error")
            if problems:
                return path, preview, None
            return path, preview, core.run(path, excel, core.output_path(folder, path), insights, thousands=not template)

        self.start(work, lambda result, error: self.generate_checked(result, error, excel, folder, not template), "generating")

    def generate_checked(self, result, error, excel, folder, thousands):
        if error:
            self.show_error(error)
            self.set_status("start_hint")
            return
        template, preview, made = result
        if made is DOUBTFUL:
            self.report = preview
            self.render_report()
            self.ask_about_doubts()
            return
        if made is None:
            self.report = preview
            self.render_report()
            count = preview.count("error")
            if not messagebox.askyesno(self.t("confirm_title"), self.t("confirm_body", count=count)):
                self.set_status("select_hint")
                return
            insights = dict(self.insights)
            self.start(lambda: core.run(template, excel, core.output_path(folder, template), insights, thousands=thousands),
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

    def ask_about_doubts(self):
        """The rules could not be sure of what they chose, so a person looks before it becomes a report."""
        auto = self.auto
        self.set_status("auto_ready_doubt", what=auto.what, count=len(auto.doubts))
        self.doubt_window = DoubtWindow(self, auto)

    def doubts_answered(self, answer):
        """change: choose the figures. keep: this is what was meant, so make it and do not ask again."""
        self.doubt_window = None
        if answer == "change":
            self.open_helper()
        elif answer == "keep":
            auto = self.auto
            notes = [(params.get("file", ""), i18n.text(auto.lang, "doubt_report_note", reason=i18n.text(auto.lang, "doubt_" + code, figure=params.get("figure", "")))) for code, params in auto.doubts]
            auto.sections = suggest.with_doubt_notes(auto.sections, notes)
            auto.stale = True
            auto.doubts = []
            self.generate()

    def show_error(self, error):
        key, params = error
        messagebox.showerror(self.t("error_title"), self.t(key, **params))

    def open_path(self, path):
        if path and os.path.exists(path):
            os.startfile(path)

    # ------------------------------------------------------------ insight candidates

    def insight_requests(self):
        """What the model can be asked about. A template asks by its 통찰 placeholders; a report made from the workbook alone can ask about each figure."""
        if self.report is None:
            return {}
        if self.template_path() or not self.auto:
            return self.report.requests
        return {request.key: request for request, _figure in self.auto.asks.values()}

    def open_insight(self):
        requests = self.insight_requests()
        if self.busy or not requests:
            return
        if self.insight_window and self.insight_window.window.winfo_exists():
            self.insight_window.window.lift()
            return
        selected = self.tree.selection()
        item = self.tree_items.get(selected[0]) if selected else None
        if self.template_path() or not self.auto:
            request = requests.get(item.key) if item else None
            self.insight_window = InsightWindow(self, request or next(iter(requests.values())))
            return
        asks = self.auto.asks
        choices = [(self.insight_label(request, figure), request) for request, figure in asks.values()]
        chosen = asks.get(item.key.partition(":")[2].casefold() if item else "") or asks.get(core.ALL_FILES) or next(iter(asks.values()))
        self.insight_window = InsightWindow(self, chosen[0], self.insight_label(*chosen), choices)

    def insight_label(self, request, figure):
        if core.ALL_FILES in request.key:
            return self.t("insight_all_files", files=request.source)
        place = request.source.rpartition("!")[0] or request.source
        return " · ".join(part for part in (figure, place) if part)

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
        if self.auto and not self.template_path():
            self.auto.stale = True
        self.preview()

    def open_helper(self):
        """With no template it changes what the report holds; with one it helps to mark it."""
        files = self.inputs()
        if not files:
            self.set_status("helper_no_excel")
            return
        plain = not self.template_path()
        key = (tuple(files), plain)
        if self.helper_window and self.helper_window.window.winfo_exists():
            if self.helper_window.key == key:
                self.helper_window.window.lift()
                return
            self.helper_window.window.destroy()  # it was reading other workbooks, or speaking to someone with a template
        self.helper_window = contents.ContentsWindow(self, files, self.open_detail) if plain else HelperWindow(self, first_workbook(files))

    def open_detail(self):
        """The finer choices, such as which rows count, for one file."""
        self.helper_window = HelperWindow(self, first_workbook(self.inputs()), True)

    def use_template(self, path):
        """The starter template has just been made from this workbook, so a report can follow at once."""
        self.vars["template"].set(path)
        if not self.vars["output"].get():
            self.vars["output"].set(os.path.dirname(path))
        self.own_template.set(True)
        self.show_template_row()
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

    def __init__(self, app, request, title=None, choices=None):
        self.app, self.request = app, request
        self.choices = choices or []
        self.round = 0
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

        if len(self.choices) > 1:  # which figure the sentence is about, or all the files together
            about = ttk.Frame(body)
            about.grid(row=0, column=0, sticky="ew")
            about.columnconfigure(1, weight=1)
            ttk.Label(about, text=app.t("insight_about"), style="Summary.TLabel").grid(row=0, column=0, padx=(0, app.px(8)))
            self.about = ttk.Combobox(about, state="readonly", values=[label for label, _request in self.choices])
            self.about.grid(row=0, column=1, sticky="ew")
            self.about.current(next(i for i, (_label, other) in enumerate(self.choices) if other is request))
            self.about.bind("<<ComboboxSelected>>", lambda _e: self.switch())
        else:
            ttk.Label(body, text=title or f"{request.key}   ({request.source})", style="Summary.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text=app.t("insight_warning"), style="Muted.TLabel", wraplength=app.px(780), justify="left").grid(
            row=1, column=0, sticky="w", pady=(small, small))
        # What is happening goes right above the list it is filling: the eye is on the list, not the bottom of the window.
        progress = ttk.Frame(body)
        progress.grid(row=2, column=0, sticky="ew")
        progress.columnconfigure(2, weight=1)
        ttk.Label(progress, text=app.t("insight_candidates"), style="Summary.TLabel").grid(row=0, column=0, sticky="w", padx=(0, app.px(10)))
        self.status = ttk.Label(progress, wraplength=app.px(600), justify="left")
        self.status.grid(row=0, column=1, sticky="w", padx=(0, app.px(10)))
        self.bar = ttk.Progressbar(progress, mode="indeterminate")
        self.bar.grid(row=0, column=2, sticky="ew")
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
        self.load_saved()

        foot = ttk.Frame(body)
        foot.grid(row=9, column=0, sticky="e", pady=(small, 0))
        self.apply_button = ttk.Button(foot, text=app.t("insight_apply"), style="Accent.TButton", command=self.apply)
        self.apply_button.pack(side="left", padx=small)
        ttk.Button(foot, text=app.t("insight_close"), command=window.destroy).pack(side="left")
        window.bind("<Escape>", lambda _e: window.destroy())
        window.after(50, self.poll)
        self.show_facts()
        self.show_calculated()
        self.make()

    def load_saved(self):
        self.chosen.delete("1.0", "end")
        saved = self.app.insights.get(self.request.key.casefold())
        if saved and saved.get("fingerprint") == self.request.fingerprint:
            self.chosen.insert("1.0", saved.get("text", ""))

    def switch(self):
        """Another figure: start over for it. An answer still on its way for the last one is ignored when it comes."""
        self.request = self.choices[self.about.current()][1]
        self.shown, self.hypotheses = [], []
        self.listbox.delete(0, "end")
        self.add_button.configure(state="disabled")
        self.load_saved()
        self.show_facts()
        self.show_calculated()
        self.working = False
        self.make()

    def show_facts(self):
        self.set_basis("\n".join(fact.text for fact in self.request.facts if fact.id != "files" and not fact.id.endswith("subject")))  # those are instructions to the model, not results

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
        self.bar.grid()
        self.bar.start(20)
        self.show_progress(1, insight.RUNS)
        seed, self.next_seed = self.next_seed, self.next_seed + insight.RUNS
        self.round += 1
        request, results, turn = self.request, self.results, self.round

        def work():
            try:
                results.put((turn, "done", insight.make_candidates(request.facts, request.lang, first_seed=seed, on_run=lambda done, runs, found: results.put((turn, "run", (done, runs, found))))))
            except insight.InsightError as exc:
                results.put((turn, "error", exc))
            except Exception as exc:  # a thread has no one else to tell
                results.put((turn, "error", insight.InsightError("other", detail=repr(exc))))

        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        if not self.window.winfo_exists():
            return
        try:
            turn, kind, value = self.results.get_nowait()
        except queue.Empty:
            self.window.after(100, self.poll)
            return
        if turn != self.round:
            self.window.after(10, self.poll)
            return
        if kind == "run":
            done, runs, found = value
            self.add_candidates(found)
            if done < runs:
                self.show_progress(done + 1, runs)
            self.window.after(100, self.poll)
            return
        self.working = False
        self.bar.stop()
        self.bar.grid_remove()
        self.again_button.configure(state="normal")
        if kind == "error":
            self.status.configure(text=self.app.t("insight_err_" + value.code, **value.params))
        else:
            self.show_candidates(value)
        self.window.after(100, self.poll)

    def show_calculated(self):
        """There from the first moment, needing no model: the program's own sentences, marked apart from what a model writes."""
        self.add_candidates(insight.Candidates("", insight.calculated_candidates(self.request.facts, self.request.lang)))

    def show_progress(self, run, runs):
        self.status.configure(text=self.app.t("insight_working", run=run, runs=runs, count=sum(not c.calculated for c in self.shown)))

    def add_candidates(self, result):
        """They go up as they come. A long wait is easier with something to read."""
        known = {c.statement for c in self.shown}
        for candidate in result.candidates:
            if candidate.statement not in known:
                known.add(candidate.statement)
                self.shown.append(candidate)
                self.listbox.insert("end", self.app.t("insight_mark_calculated" if candidate.calculated else "insight_mark_model") + candidate.statement)
                if candidate.calculated:
                    self.listbox.itemconfigure("end", foreground="#1F3864")
        for hypothesis in result.hypotheses:
            if hypothesis not in self.hypotheses:
                self.hypotheses.append(hypothesis)

    def show_candidates(self, result):
        self.add_candidates(result)
        written = sum(not c.calculated for c in self.shown)
        key = "insight_done" if written else "insight_none"
        self.status.configure(text=self.app.t(key, count=written, rejected=result.rejected, model=result.model))
        if self.hypotheses and not self.listbox.curselection():
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

    def __init__(self, app, excel, plain=False):
        self.app, self.excel = app, excel
        self.plain = plain  # no template in play: choosing what the report shows, with no notation in sight
        self.key = ((excel,), plain)
        self.source = None
        self.shape = None
        self.found = []
        self.results = queue.Queue()
        self.job = 0
        self.working = False
        self.sheets = []
        self.bases = []
        window = self.window = tk.Toplevel(app.root)
        window.title(app.t("helper_title_auto" if plain else "helper_title"))
        window.geometry(f"{app.px(960)}x{app.px(640)}")
        window.transient(app.root)
        body = ttk.Frame(window, padding=app.px(12))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)
        small = app.px(4)

        ttk.Label(body, text=app.t("helper_intro_auto" if plain else "helper_intro"), style="Muted.TLabel", wraplength=app.px(920), justify="left").grid(row=0, column=0, sticky="w")

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
        if plain:
            self.tree.configure(displaycolumns=("what", "result"))
        for status, color in STATUS_COLOR.items():
            self.tree.tag_configure(status, foreground=color if status != "ok" else "")
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.show_row())
        self.tree.bind("<Double-Button-1>", lambda _e: self.copy())

        self.key_var = tk.StringVar()
        if not plain:
            ttk.Entry(body, textvariable=self.key_var, state="readonly").grid(row=4, column=0, sticky="ew", pady=(app.px(8), small))
        self.detail = ttk.Label(body, wraplength=app.px(920), justify="left")
        self.detail.grid(row=5, column=0, sticky="w", pady=(app.px(8) if plain else 0, 0))

        foot = ttk.Frame(body)
        foot.grid(row=6, column=0, sticky="e", pady=(small, 0))
        self.copy_button = ttk.Button(foot, text=app.t("helper_copy"), command=self.copy, state="disabled")
        self.use_button = ttk.Button(foot, text=app.t("helper_use"), style="Accent.TButton", command=self.use_contents, state="disabled")
        self.starter_button = ttk.Button(foot, text=app.t("helper_starter_auto" if plain else "helper_starter"), style="TButton" if plain else "Accent.TButton", command=self.make_starter, state="disabled")
        if plain:
            self.use_button.pack(side="left")
        else:
            self.copy_button.pack(side="left")
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
        self.boxes["base"]["values"] = [self.choice(b.spec, b.name, [o.name for o in self.bases]) for b in self.bases]
        self.boxes["base"].current(self.bases.index(base))
        self.fill_base(base)
        self.show_found(found)

    def choice(self, spec, name, names):
        """How a column is offered: by its heading alone, with its letter only when the heading cannot tell it apart. A template's author needs the letter."""
        shown = name or self.app.t("helper_no_name")
        if not self.plain:
            return f"{spec} · {shown}"
        return shown if name and names.count(name) == 1 else f"{shown} ({spec})"

    def value_text(self, value, base):
        if value.every_row:
            return self.app.t("helper_every_row")
        if value.counted:
            return self.app.t("helper_counted")
        return self.choice(value.spec, value.name, [v.name for v in base.values])

    def fill_base(self, base):
        value = suggest.default_value(base)
        self.boxes["value"]["values"] = [self.value_text(v, base) for v in base.values]
        if value is not None:
            self.boxes["value"].current(base.values.index(value))
        how = suggest.default_how(base, value)
        self.boxes["how"]["values"] = [self.app.t("helper_how_" + name) for name in self.HOWS]
        self.boxes["how"].current(self.HOWS.index(how))
        none = self.app.t("helper_none")
        names = [none] + [self.choice(g.spec, g.name, [o.name for o in base.groups]) for g in base.groups]
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
        line = self.app.t("helper_shape_dates" if base.kind == "date" else "helper_shape_names", sheet=self.shape.sheet, spec=base.name if self.plain and base.name else base.spec, first=base.first, last=base.last, count=base.count)
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
        shown = suggest.in_report(found) if self.plain else found  # with no template to paste into, only what the report will hold
        for number, item in enumerate(found):
            if not any(item is other for other in shown) and any(i.status != "error" for i in found):
                continue
            result = item.result if item.status != "error" else self.app.issue_text(item.issues[0])
            place = self.app.t("helper_place_" + item.place)
            self.tree.insert("", "end", iid=str(number), tags=(item.status,), values=(self.app.t("suggest_" + item.code), result, place))
        works = any(item.status != "error" for item in found)
        self.starter_button.configure(state="normal" if works else "disabled")
        self.use_button.configure(state="normal" if works else "disabled")
        self.copy_button.configure(state="disabled")
        self.key_var.set("")
        self.detail.configure(text=self.app.t("select_hint" if self.plain else "helper_pick_row") if works else self.app.t("helper_nothing"))
        if found and works:
            self.tree.selection_set(self.tree.get_children()[0])

    def picked(self):
        chosen = self.tree.selection()
        return self.found[int(chosen[0])] if chosen else None

    def show_row(self):
        item = self.picked()
        if item is None:
            return
        self.key_var.set(item.typed)
        self.copy_button.configure(state="normal" if item.status != "error" else "disabled")
        lines = [] if self.plain else [self.app.t("helper_where_" + item.place)]
        if item.place == "text":
            lines.append(item.result)
        lines += [self.app.issue_text(issue) for issue in item.issues[:3]]
        self.detail.configure(text="\n".join(lines))

    def copy(self):
        item = self.picked()
        if self.plain or item is None or item.status == "error":
            return
        self.window.clipboard_clear()
        self.window.clipboard_append(item.typed)
        self.window.update()
        self.detail.configure(text=self.app.t("helper_copied", reference=item.typed) + "\n" + self.app.t("helper_where_" + item.place))

    def use_contents(self):
        if self.app.busy:
            self.detail.configure(text=self.app.t("helper_working"))
            return
        value = self.base().values[self.boxes["value"].current()] if self.base().values else None
        self.app.use_contents(self.excel, self.found, suggest.describe(self.shape, value), suggest.heading(self.source, self.shape))
        self.window.destroy()

    def make_starter(self):
        folder = self.app.vars["output"].get().strip() or os.path.dirname(self.excel)
        path = suggest.starter_path(folder, self.excel)
        title = os.path.splitext(os.path.basename(self.excel))[0]
        try:
            suggest.make_starter(path, title, self.found, suggest.heading(self.source, self.shape))
        except OSError as exc:
            self.detail.configure(text=self.app.t("err_cannot_save", path=path, detail=str(exc)))
            return
        self.app.use_template(path)
        self.detail.configure(text=self.app.t("helper_starter_saved", path=path))
        self.use_button.configure(state="disabled")  # the report now follows that template


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
