"""
polymorphize_gui — the desktop front end.

Version 6.7.

One window, three tabs, and one rule: a failed endpoint is impossible to
miss. Every failure appears three times over, in the banner across the top of
the window, as a red row in the results table, and in the detail pane with the
cell reference and the correction spelled out.
"""

from __future__ import annotations

import os
import queue
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import polymorphize_generate as gen
import polymorphize_template as tmpl
import polymorphize_validate as val

__version__ = "6.7"

# --------------------------------------------------------------------------- #
# Drag and drop
# --------------------------------------------------------------------------- #

#: ``tkinterdnd2`` provides the Tk drag-and-drop bindings. It is optional: the
#: window must open and every field must remain usable through Browse whether
#: or not it is installed, because an optional convenience should never be able
#: to stop the application starting.
try:                                                           # pragma: no cover
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND = True
except Exception:                                              # noqa: BLE001
    DND_FILES = None
    TkinterDnD = None
    DND = False


def _dropped_paths(data):
    """Tk hands dropped files over as a brace-quoted list.

    ``{C:/A folder/book.xlsx} /tmp/other.xlsx`` is two paths, the first braced
    because it contains a space. Tk does the quoting, so the parsing has to be
    done here rather than by splitting on whitespace.
    """
    out, buf, depth = [], "", 0
    for ch in str(data):
        if ch == "{":
            depth += 1
            if depth == 1:
                continue
        if ch == "}":
            depth -= 1
            if depth == 0:
                out.append(buf)
                buf = ""
                continue
        if ch.isspace() and depth == 0:
            if buf:
                out.append(buf)
                buf = ""
            continue
        buf += ch
    if buf:
        out.append(buf)
    return [p for p in (s.strip().strip('"') for s in out) if p]

# Colours. Muted enough to read all day, distinct enough to scan.
INK = "#1a1a1a"
MUTED = "#5a6472"
LINE = "#c9d2e0"
BG = "#f4f6fa"
PANEL = "#ffffff"
OK_BG = "#e8f4ea"
OK_FG = "#1e6b34"
FAIL_BG = "#fdeaea"
FAIL_FG = "#a11b1b"
WARN_BG = "#fdf6e3"
WARN_FG = "#8a6100"
SKIP_FG = "#8b95a5"
ACCENT = "#1f3864"

STATUS_LABEL = {"ok": "Will generate", "failed": "FAILED",
                "skipped": "Support sheet"}


class App(ttk.Frame):
    """The main window."""

    def __init__(self, master):
        super().__init__(master, padding=0)
        self.master.title("SOR Polymorphizer %s" % __version__)
        self.master.geometry("1180x860")
        self.master.minsize(980, 700)

        self.v_workbook = tk.StringVar()
        self.v_sor = tk.StringVar()
        self.v_outdir = tk.StringVar()
        self.v_level = tk.StringVar(value="lenient")
        self.v_format = tk.StringVar(value="yaml")
        #: The classifier's verdict on the chosen workbook.
        self.v_verdict = tk.StringVar(value="")
        self.v_output_mode = tk.StringVar(value="merged")
        self.v_api_version = tk.StringVar(value="1.0.0")
        # On by default. The whole point of writing it from the run is that
        # the document and the specification cannot drift apart, and that only
        # holds if it is written every time.
        self.v_export = tk.BooleanVar(value=True)
        self.v_status = tk.StringVar(value="Choose a mapping workbook to begin.")

        self.reports = []
        self.by_sheet = {}
        self._queue = queue.Queue()
        self._busy = False

        self._style()
        self._build()
        self.pack(fill="both", expand=True)
        self.after(120, self._drain)

    # ------------------------------------------------------------------ style

    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        self.master.configure(bg=BG)
        s.configure(".", background=BG, foreground=INK,
                    font=("Segoe UI", 10))
        s.configure("TFrame", background=BG)
        s.configure("Panel.TFrame", background=PANEL, relief="flat")
        s.configure("TLabel", background=BG, foreground=INK)
        s.configure("Muted.TLabel", background=BG, foreground=MUTED)
        s.configure("PanelMuted.TLabel", background=PANEL, foreground=MUTED)
        s.configure("Panel.TLabel", background=PANEL, foreground=INK)
        s.configure("H1.TLabel", background=BG, foreground=ACCENT,
                    font=("Segoe UI Semibold", 15))
        s.configure("H2.TLabel", background=BG, foreground=ACCENT,
                    font=("Segoe UI Semibold", 11))
        s.configure("PanelH2.TLabel", background=PANEL, foreground=ACCENT,
                    font=("Segoe UI Semibold", 11))
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", padding=(18, 9),
                    font=("Segoe UI", 10))
        s.map("TNotebook.Tab", background=[("selected", PANEL)])
        s.configure("TButton", padding=(14, 7))
        s.configure("Primary.TButton", padding=(18, 8),
                    font=("Segoe UI Semibold", 10))
        s.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    rowheight=25, borderwidth=1)
        s.configure("Treeview.Heading", font=("Segoe UI Semibold", 9),
                    padding=(6, 6))
        s.configure("TEntry", padding=5)
        s.configure("Drop.TEntry", padding=5, fieldbackground=OK_BG)
        s.configure("Ok.TLabel", background=OK_BG, foreground=OK_FG,
                    font=("Segoe UI Semibold", 10), padding=(12, 9))
        s.configure("Fail.TLabel", background=FAIL_BG, foreground=FAIL_FG,
                    font=("Segoe UI Semibold", 10), padding=(12, 9))
        s.configure("Neutral.TLabel", background=BG, foreground=MUTED,
                    padding=(12, 9))

    # ------------------------------------------------------------------ build

    def _build(self):
        head = ttk.Frame(self, padding=(20, 16, 20, 8))
        head.pack(fill="x")
        ttk.Label(head, text="SOR Polymorphizer", style="H1.TLabel").pack(anchor="w")
        ttk.Label(head, style="Muted.TLabel",
                  text="One OpenAPI specification from the whole mapping "
                       "workbook. The workbook is authoritative.").pack(
                           anchor="w", pady=(2, 0))

        self.banner = ttk.Label(self, textvariable=self.v_status,
                                style="Neutral.TLabel", anchor="w")
        self.banner.pack(fill="x", padx=20, pady=(6, 10))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        self.nb = nb
        self.tab_run = ttk.Frame(nb, padding=16)
        self.tab_findings = ttk.Frame(nb, padding=16)
        self.tab_log = ttk.Frame(nb, padding=16)
        nb.add(self.tab_run, text="Workbook")
        nb.add(self.tab_findings, text="Findings")
        nb.add(self.tab_log, text="Log")

        self._build_run(self.tab_run)
        self._build_findings(self.tab_findings)
        self._build_log(self.tab_log)

    def _file_row(self, parent, row, label, var, hint, *, folder=False,
                  save=False, kind="workbook"):
        if DND:
            hint += ("  Drag a folder here, or use Browse." if folder
                     else "  Drag files here, or use Browse.")
        # The label and its hint share one row spanning every column. Putting
        # the label in column 0 lets the entry's column width clip it.
        head = ttk.Frame(parent)
        head.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 2))
        ttk.Label(head, text=label).pack(side="left")
        ttk.Label(head, text="   " + hint, style="Muted.TLabel",
                  font=("Segoe UI", 9)).pack(side="left")
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row + 1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        self._make_drop_target(entry, var, folder=folder, kind=kind)

        def browse():
            if folder:
                path = filedialog.askdirectory(title=label)
            elif kind == "sor":
                chosen = filedialog.askopenfilenames(
                    title=label,
                    filetypes=[("OpenAPI or Swagger", "*.yaml *.yml *.json"),
                               ("All files", "*.*")])
                path = "; ".join(chosen) if chosen else ""
            elif save:
                path = filedialog.asksaveasfilename(
                    title=label, defaultextension=".xlsx",
                    filetypes=[("Excel workbook", "*.xlsx")])
            else:
                path = filedialog.askopenfilename(
                    title=label,
                    filetypes=[("Excel workbook", "*.xlsx *.xlsm"),
                               ("All files", "*.*")])
            if path:
                var.set(path)
                if kind == "workbook" and not folder and not save:
                    self._on_workbook_chosen(path)

        ttk.Button(parent, text="Browse", command=browse).grid(
            row=row + 1, column=2, sticky="e")
        return entry

    def _make_drop_target(self, widget, var, *, folder=False, kind="workbook"):
        """Accept a dropped file or folder on one field.

        A workbook dropped on the output field, or a folder dropped on the
        workbook field, is corrected rather than refused: the intent is
        obvious, and refusing it would be pedantry.
        """
        if not DND:
            return
        widget.drop_target_register(DND_FILES)

        def on_enter(_event):
            widget.configure(style="Drop.TEntry")

        def on_leave(_event):
            widget.configure(style="TEntry")

        def on_drop(event):
            on_leave(event)
            paths = _dropped_paths(event.data)
            if not paths:
                return
            path = paths[0]
            if len(paths) > 1:
                self._say("%d items were dropped; using %s"
                          % (len(paths), os.path.basename(path)))
            if folder:
                if os.path.isfile(path):
                    path = os.path.dirname(path)
                self.v_outdir.set(path)
            elif kind == "sor":
                # Several SOR specifications are normal, so a drop adds to
                # what is there rather than replacing it.
                existing = [p for p in
                            (x.strip() for x in var.get().split(";")) if p]
                for p in paths:
                    if p not in existing:
                        existing.append(p)
                var.set("; ".join(existing))
            else:
                if os.path.isdir(path):
                    books = sorted(f for f in os.listdir(path)
                                   if f.lower().endswith((".xlsx", ".xlsm")))
                    if len(books) != 1:
                        messagebox.showinfo(
                            "Which workbook?",
                            "That folder holds %d workbooks. Drop the one you "
                            "want, or use Browse." % len(books))
                        return
                    path = os.path.join(path, books[0])
                if not path.lower().endswith((".xlsx", ".xlsm")):
                    messagebox.showinfo(
                        "Not a workbook",
                        "%s is not an .xlsx or .xlsm file."
                        % os.path.basename(path))
                    return
                self.v_workbook.set(path)
                self._on_workbook_chosen(path)

        widget.dnd_bind("<<DropEnter>>", on_enter)
        widget.dnd_bind("<<DropLeave>>", on_leave)
        widget.dnd_bind("<<Drop>>", on_drop)

    def _build_run(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        form = ttk.Frame(parent)
        form.grid(row=0, column=0, columnspan=2, sticky="ew")
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=0)
        self._file_row(form, 0, "Mapping workbook",
                       self.v_workbook,
                       "A Level format workbook or a field mapping document. "
                       "The format is detected when you choose the file.")
        # The verdict on the chosen file. Classification only: it says what the
        # file is and how much of it, never whether it is good.
        self.lbl_format = ttk.Label(form, textvariable=self.v_verdict,
                                    style="Muted.TLabel",
                                    font=("Segoe UI", 9), anchor="w")
        self.lbl_format.grid(row=2, column=0, columnspan=3, sticky="ew",
                             pady=(2, 6))
        self._file_row(form, 3, "System of Record specification",
                       self.v_sor,
                       "The SOR swagger, or a folder of them. Separate several "
                       "with a semicolon. Optional, and without it every SOR "
                       "field name is taken on trust.",
                       kind="sor")
        self._file_row(form, 6, "Output folder", self.v_outdir,
                       "Where the generated specification is written.",
                       folder=True)

        opts = ttk.Frame(parent)
        opts.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Label(opts, text="Checking level", style="H2.TLabel").grid(
            row=0, column=0, sticky="w", columnspan=3)
        ttk.Radiobutton(opts, text="Lenient  —  only what stops an endpoint "
                                   "generating",
                        variable=self.v_level, value="lenient").grid(
            row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Radiobutton(opts, text="Strict  —  the full template contract, "
                                   "including descriptions and examples",
                        variable=self.v_level, value="strict").grid(
            row=2, column=0, sticky="w")

        mode = ttk.Frame(parent)
        mode.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        ttk.Label(mode, text="Output", style="H2.TLabel").grid(
            row=0, column=0, sticky="w", columnspan=3)
        ttk.Radiobutton(mode, text="One specification for the whole workbook",
                        variable=self.v_output_mode, value="merged").grid(
            row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Radiobutton(mode, text="One specification per sheet, for debugging",
                        variable=self.v_output_mode, value="split").grid(
            row=2, column=0, sticky="w")
        ttk.Label(mode, style="Muted.TLabel", font=("Segoe UI", 9),
                  text="Shared groups are specialised across operations only in "
                       "the merged output: a separate file has its own "
                       "namespace.").grid(
            row=3, column=0, sticky="w", columnspan=3, pady=(2, 0))

        doc = ttk.Frame(parent)
        doc.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        ttk.Label(doc, text="Documentation", style="H2.TLabel").grid(
            row=0, column=0, sticky="w", columnspan=3)
        ttk.Checkbutton(
            doc, variable=self.v_export,
            text="Also write the field mapping document, in the API COE "
                 "format").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(doc, style="Muted.TLabel", font=("Segoe UI", 9),
                  text="One sheet per SOR endpoint, with three columns saying "
                       "what reached the interface and what did not. It is "
                       "itself a valid input, so corrections can be made in "
                       "it and fed straight back in.").grid(
            row=2, column=0, sticky="w", columnspan=3, pady=(2, 0))

        line2 = ttk.Frame(parent)
        line2.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        ttk.Label(line2, text="Output format").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(line2, text="YAML", variable=self.v_format,
                        value="yaml").grid(row=0, column=1, padx=(10, 4))
        ttk.Radiobutton(line2, text="JSON", variable=self.v_format,
                        value="json").grid(row=0, column=2, padx=(0, 24))
        ttk.Label(line2, text="API version").grid(row=0, column=3, sticky="w")
        ttk.Entry(line2, textvariable=self.v_api_version, width=12).grid(
            row=0, column=4, padx=(10, 0))

        btns = ttk.Frame(parent)
        btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(18, 12))
        self.b_check = ttk.Button(btns, text="Check the workbook",
                                  command=self._check)
        self.b_check.pack(side="left")
        self.b_generate = ttk.Button(btns, text="Generate specifications",
                                     style="Primary.TButton", command=self._generate)
        self.b_generate.pack(side="left", padx=(10, 0))
        ttk.Button(btns, text="Write a blank template…",
                   command=self._template).pack(side="right")

        ttk.Label(parent, text="Sheets", style="H2.TLabel").grid(
            row=7, column=0, sticky="w", columnspan=2)
        ttk.Label(parent, style="Muted.TLabel", font=("Segoe UI", 9),
                  text="Select a row to see its findings on the Findings tab.").grid(
            row=8, column=0, sticky="w", columnspan=2, pady=(0, 6))

        wrap = ttk.Frame(parent)
        wrap.grid(row=9, column=0, columnspan=2, sticky="nsew")
        parent.rowconfigure(9, weight=1)
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)

        cols = ("status", "pattern", "endpoints", "errors", "warnings")
        self.tree = ttk.Treeview(wrap, columns=cols, show="tree headings",
                                 selectmode="browse")
        self.tree.heading("#0", text="Sheet")
        self.tree.column("#0", width=280, anchor="w")
        for name, title, width, anchor in (
                ("status", "Status", 120, "w"),
                ("pattern", "Pattern", 80, "center"),
                ("endpoints", "SOR endpoints", 110, "center"),
                ("errors", "Errors", 70, "center"),
                ("warnings", "Warnings", 80, "center")):
            self.tree.heading(name, text=title)
            self.tree.column(name, width=width, anchor=anchor)
        self.tree.tag_configure("failed", background=FAIL_BG, foreground=FAIL_FG)
        self.tree.tag_configure("ok", background=PANEL)
        self.tree.tag_configure("warned", background=WARN_BG)
        self.tree.tag_configure("skipped", foreground=SKIP_FG)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<<TreeviewSelect>>", self._on_sheet_select)
        self.tree.bind("<Double-1>", lambda _e: self.nb.select(self.tab_findings))

    def _build_findings(self, parent):
        parent.columnconfigure(0, weight=1)
        # The detail pane gets the larger share: it holds the correction, and a
        # correction the analyst has to scroll to find is a correction missed.
        parent.rowconfigure(2, weight=2)
        parent.rowconfigure(4, weight=3)

        self.v_findings_for = tk.StringVar(value="No sheet selected.")
        ttk.Label(parent, textvariable=self.v_findings_for,
                  style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(parent, style="Muted.TLabel", font=("Segoe UI", 9),
                  text="Errors stop the endpoint. Warnings do not, but each one "
                       "is worth a look.").grid(row=1, column=0, sticky="w",
                                                pady=(0, 6))

        wrap = ttk.Frame(parent)
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        cols = ("code", "cell", "what")
        self.ftree = ttk.Treeview(wrap, columns=cols, show="tree headings",
                                  selectmode="browse")
        self.ftree.heading("#0", text="Severity")
        self.ftree.column("#0", width=90, anchor="w")
        for name, title, width in (("code", "Code", 70),
                                   ("cell", "Cell", 90),
                                   ("what", "What was found", 720)):
            self.ftree.heading(name, text=title)
            self.ftree.column(name, width=width, anchor="w")
        self.ftree.tag_configure("error", background=FAIL_BG, foreground=FAIL_FG)
        self.ftree.tag_configure("warning", background=WARN_BG, foreground=WARN_FG)
        self.ftree.grid(row=0, column=0, sticky="nsew")
        fsb = ttk.Scrollbar(wrap, orient="vertical", command=self.ftree.yview)
        fsb.grid(row=0, column=1, sticky="ns")
        self.ftree.configure(yscrollcommand=fsb.set)
        self.ftree.bind("<<TreeviewSelect>>", self._on_finding_select)

        ttk.Label(parent, text="Detail", style="H2.TLabel").grid(
            row=3, column=0, sticky="w", pady=(14, 4))
        dwrap = ttk.Frame(parent)
        dwrap.grid(row=4, column=0, sticky="nsew")
        dwrap.columnconfigure(0, weight=1)
        dwrap.rowconfigure(0, weight=1)
        self.detail = tk.Text(dwrap, height=12, wrap="word", relief="flat",
                              background=PANEL, foreground=INK,
                              highlightthickness=1, highlightbackground=LINE,
                              font=("Segoe UI", 10), padx=12, pady=10)
        self.detail.grid(row=0, column=0, sticky="nsew")
        dsb = ttk.Scrollbar(dwrap, orient="vertical", command=self.detail.yview)
        dsb.grid(row=0, column=1, sticky="ns")
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.tag_configure("label", foreground=ACCENT,
                                  font=("Segoe UI Semibold", 10))
        self.detail.tag_configure("code", foreground=MUTED,
                                  font=("Consolas", 10))
        self.detail.configure(state="disabled")

    def _build_log(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        ttk.Label(parent, text="Run log", style="H2.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6))
        self.log = tk.Text(parent, wrap="none", relief="flat", background=PANEL,
                           foreground=INK, highlightthickness=1,
                           highlightbackground=LINE, font=("Consolas", 10),
                           padx=10, pady=8)
        self.log.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(parent, orient="vertical", command=self.log.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb.set, state="disabled")

    # ------------------------------------------------------------- plumbing

    def _say(self, line):
        self._queue.put(("log", line))

    def _drain(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log":
                    self.log.configure(state="normal")
                    self.log.insert("end", str(payload) + "\n")
                    self.log.see("end")
                    self.log.configure(state="disabled")
                elif kind == "reports":
                    self._show_reports(*payload)
                elif kind == "busy":
                    self._set_busy(payload)
                elif kind == "error":
                    messagebox.showerror("Something went wrong", payload)
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.b_check.configure(state=state)
        self.b_generate.configure(state=state)
        self.master.configure(cursor="watch" if busy else "")

    def _banner_set(self, style, message):
        self.banner.configure(style=style)
        self.v_status.set(message)

    def _on_workbook_chosen(self, path):
        if not self.v_outdir.get():
            self.v_outdir.set(os.path.join(os.path.dirname(path), "openapi"))
        self._classify(path)

    def _classify(self, path):
        """Say what the chosen file is, and let that drive the interface.

        Bounded and read-only, so this stays well under a third of a second
        even on an 8.7 MB workbook. Opening such a file the ordinary way to
        read its header takes about 76 seconds and would look like a hang.
        """
        try:
            import polymorphize_classify as cls
            verdict = cls.classify(path)
        except Exception as exc:                               # noqa: BLE001
            self.v_verdict.set("Could not read this file yet: %s" % exc)
            return None

        self.v_verdict.set(verdict.line())
        if not verdict.readable:
            self._banner_set("Neutral.TLabel",
                             "That file cannot be read yet. It may be open in "
                             "Excel.")
            return verdict
        if verdict.fmt == cls.UNKNOWN:
            self._banner_set("Bad.TLabel",
                             "Neither input format: no Level columns and no "
                             "Parameter Type column found.")
            return verdict
        if verdict.needs_specification:
            # A field mapping document describes an interface against a System
            # of Record, so the specification field earns its place here.
            self._banner_set("Neutral.TLabel",
                             "Field mapping document. Supply the System of "
                             "Record specification below, or every field name "
                             "in it is taken on trust.")
        else:
            self._banner_set("Neutral.TLabel",
                             "Ready. Check the workbook, or go straight to "
                             "generating.")
        return verdict

    def _run_off_thread(self, work):
        if self._busy:
            return
        wb = self.v_workbook.get().strip()
        if not wb or not os.path.isfile(wb):
            messagebox.showwarning("No workbook",
                                   "Choose a mapping workbook first.")
            return
        self._queue.put(("busy", True))

        def target():
            try:
                work(wb)
            except Exception:                                  # noqa: BLE001
                self._say(traceback.format_exc())
                self._queue.put(("error",
                                 "The run stopped unexpectedly. The Log tab has "
                                 "the detail."))
            finally:
                self._queue.put(("busy", False))
        threading.Thread(target=target, daemon=True).start()

    # ------------------------------------------------------------- actions

    def _sor_list(self):
        return [p for p in (x.strip() for x in self.v_sor.get().split(";")) if p]

    def _check(self):
        strict = self.v_level.get() == "strict"
        sor = self._sor_list()

        def work(wb):
            self._say("Checking %s (%s)" % (wb, self.v_level.get()))
            if sor:
                self._say("Verifying against %s" % "; ".join(sor))
            else:
                self._say("No SOR specification: the workbook is taken on trust.")
            reports = val.validate_workbook(wb, strict=strict, sor=sor or None)
            for r in reports:
                self._say("  %-8s %-34s %s"
                          % (r.status, r.sheet, r.pattern or ""))
            self._queue.put(("reports", (reports, None)))
        self._run_off_thread(work)

    def _generate(self):
        out_dir = self.v_outdir.get().strip()
        if not out_dir:
            messagebox.showwarning("No output folder",
                                   "Choose a folder for the generated "
                                   "specifications.")
            return
        strict = self.v_level.get() == "strict"
        fmt = self.v_format.get()
        split = self.v_output_mode.get() == "split"
        sor = self._sor_list()
        api_version = self.v_api_version.get().strip() or "1.0.0"
        export = bool(self.v_export.get())

        def work(wb):
            self._say("Checking %s before generating" % wb)
            reports = val.validate_workbook(wb, strict=strict, sor=sor or None)
            failed = [r for r in reports if r.status == "failed"]
            if failed and not self._confirm_partial(failed, len(reports)):
                self._say("Cancelled by the user.")
                self._queue.put(("reports", (reports, None)))
                return
            self._say("Writing to %s" % out_dir)
            out = gen.run(wb, out_dir, strict=strict, version=api_version,
                          fmt=fmt, split=split, sor=sor or None,
                          export=export, log=self._say)
            self._queue.put(("reports", (reports, out)))
        self._run_off_thread(work)

    def _confirm_partial(self, failed, total):
        names = "\n".join("    • %s" % r.sheet for r in failed[:10])
        more = "" if len(failed) <= 10 else "\n    and %d more" % (len(failed) - 10)
        return messagebox.askyesno(
            "%d endpoint%s will not be generated"
            % (len(failed), "" if len(failed) == 1 else "s"),
            "These sheets have errors and no specification will be written "
            "for them:\n\n%s%s\n\nThe remaining endpoints will generate "
            "normally. Continue?" % (names, more),
            icon="warning", default="no")

    def _template(self):
        """Write a blank template, in whichever input format is wanted."""
        fmt = self._ask_template_format()
        if fmt is None:
            return
        mapping = fmt == "mapping"
        default = ("field_mapping_template.xlsx" if mapping
                   else "SOR_mapping_template.xlsx")
        path = filedialog.asksaveasfilename(
            title="Write a blank template",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            if mapping:
                tmpl.write_mapping_template(path)
            else:
                tmpl.write_template(path)
        except Exception as exc:                               # noqa: BLE001
            messagebox.showerror("Could not write the template", str(exc))
            return
        self._say("Wrote %s" % path)
        messagebox.showinfo(
            "Template written",
            "Wrote %s\n\n%s\n\nThe How to use sheet states the contract, "
            "and the Vocabulary sheet holds the lists behind the dropdowns."
            % (path,
               "It holds a worked operation and a blank one to copy, one "
               "sheet per endpoint and per SOR endpoint." if mapping else
               "It holds two worked examples and a blank operation sheet. "
               "Copy the blank sheet once per endpoint."))

    def _ask_template_format(self):
        """Which of the two input formats to write. ``None`` to cancel.

        Asked rather than assumed. The two formats are not interchangeable:
        the Level format can express a runtime variant and the field mapping
        document cannot, so the choice decides what the analyst is able to
        describe, and it is not a preference to be guessed at.
        """
        win = tk.Toplevel(self.master)
        win.title("Which format?")
        win.transient(self.master)
        win.resizable(False, False)
        win.configure(bg=BG)
        choice = tk.StringVar(value="level")
        answer = {}

        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Which input format?",
                  style="H2.TLabel").pack(anchor="w")
        ttk.Radiobutton(frame, variable=choice, value="level",
                        text="Level format").pack(anchor="w", pady=(10, 0))
        ttk.Label(frame, style="Muted.TLabel", font=("Segoe UI", 9),
                  wraplength=460, justify="left",
                  text="Nesting in Level 1 to Level N columns, and one column "
                       "per SOR endpoint on a sheet. The only format that can "
                       "express a runtime variant, so an operation served by "
                       "several SOR endpoints needs this one.").pack(
                           anchor="w", padx=(22, 0))
        ttk.Radiobutton(frame, variable=choice, value="mapping",
                        text="Field mapping document").pack(anchor="w",
                                                            pady=(12, 0))
        ttk.Label(frame, style="Muted.TLabel", font=("Segoe UI", 9),
                  wraplength=460, justify="left",
                  text="Nesting as a dotted path in one column, banner in "
                       "rows 1 to 9, one SOR endpoint per sheet. The format "
                       "the API COE circulates, and the one this tool writes "
                       "back out.").pack(anchor="w", padx=(22, 0))

        row = ttk.Frame(frame)
        row.pack(fill="x", pady=(18, 0))

        def take():
            answer["fmt"] = choice.get()
            win.destroy()

        ttk.Button(row, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(row, text="Write it", style="Primary.TButton",
                   command=take).pack(side="right", padx=(0, 8))
        win.bind("<Return>", lambda _e: take())
        win.bind("<Escape>", lambda _e: win.destroy())
        win.grab_set()
        self.master.wait_window(win)
        return answer.get("fmt")

    # ------------------------------------------------------------- display

    def _show_reports(self, reports, out):
        self.reports = reports
        self.by_sheet = {r.sheet: r for r in reports}
        self.tree.delete(*self.tree.get_children())
        for r in reports:
            tag = r.status if r.status in ("failed", "skipped") else (
                "warned" if r.warnings else "ok")
            eps = (len(r.result.layout.sor_cols)
                   if r.result is not None and r.result.layout is not None else 0)
            self.tree.insert(
                "", "end", iid=r.sheet, text=r.sheet,
                values=(STATUS_LABEL.get(r.status, r.status),
                        r.pattern or "—",
                        eps or "—",
                        len(r.errors) or "", len(r.warnings) or ""),
                tags=(tag,))

        s = val.summarise(reports)
        sor_note = ("" if self._sor_list() else
                    "  No SOR specification supplied, so field names were "
                    "taken on trust.")
        if s["failed"]:
            first = next(r for r in reports if r.status == "failed")
            self._banner_set(
                "Fail.TLabel",
                "%d of %d endpoints FAILED and no specification was written for "
                "them. First: %s. Open the Findings tab for the correction.%s"
                % (s["failed"], s["operations"], first.sheet, sor_note))
            self.tree.selection_set(first.sheet)
            self.tree.see(first.sheet)
        elif s["warnings"]:
            self._banner_set(
                "Ok.TLabel",
                "All %d endpoints will generate. %d warning%s to look at.%s"
                % (s["ok"], s["warnings"], "" if s["warnings"] == 1 else "s",
                   sor_note))
        else:
            self._banner_set("Ok.TLabel",
                             "All %d endpoints will generate with no findings.%s"
                             % (s["ok"], sor_note))
        if out is not None:
            self._say("")
            self._say("Generated %d of %d endpoints."
                      % (len(out.ok), len(out.ok) + len(out.failed)))
            if out.merged is not None and out.merged.errors:
                self._banner_set(
                    "Fail.TLabel",
                    "The endpoints generated but they could not be combined "
                    "into one specification, so none was written. See the Log "
                    "tab.")
            elif out.merged is not None and out.merged.document is not None:
                m = out.merged.stats
                self._banner_set(
                    "Ok.TLabel",
                    "Wrote one specification: %d operations, %d schemas, %d "
                    "groups hoisted, %d specialised across operations.%s%s"
                    % (m["operations"], m["schemas"], m["hoisted_groups"],
                       m["split_groups"],
                       _artefact_note(out),
                       "  %d endpoint(s) FAILED and are not in it."
                       % len(out.failed) if out.failed else ""))
                if out.failed:
                    self.banner.configure(style="Fail.TLabel")
            if out.export_error:
                # The run promised a spreadsheet and did not produce one. That
                # belongs in front of the user, not in the log.
                self.banner.configure(style="Fail.TLabel")
                messagebox.showwarning(
                    "The field mapping document was not written",
                    "The specifications were written and are complete. The "
                    "spreadsheet was not.\n\n%s\n\nIf it was open in Excel, "
                    "close it and generate again. Otherwise send "
                    "FIELD_MAPPING_NOT_WRITTEN.md from the output folder to "
                    "the maintainer." % out.export_error)

    def _on_sheet_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        report = self.by_sheet.get(sel[0])
        if report is None:
            return
        self.v_findings_for.set(
            "%s  —  %s" % (report.sheet,
                                STATUS_LABEL.get(report.status, report.status)))
        self.ftree.delete(*self.ftree.get_children())
        self._detail_set([])
        ordered = sorted(report.findings,
                         key=lambda f: (f.severity != "error", f.code))
        for i, f in enumerate(ordered):
            self.ftree.insert("", "end", iid=str(i),
                              text="Error" if f.severity == "error" else "Warning",
                              values=(f.code, f.cell or "—",
                                      f.what.replace("\n", " ")),
                              tags=(f.severity,))
        self._current = ordered
        if ordered:
            self.ftree.selection_set("0")
        elif report.status == "ok":
            self._detail_set([("", "Nothing to correct on this sheet.")])

    def _on_finding_select(self, _event=None):
        sel = self.ftree.selection()
        if not sel or not getattr(self, "_current", None):
            return
        f = self._current[int(sel[0])]
        # Found and Fix come first, and fit without scrolling. Everything the
        # analyst needs to act is above the fold.
        self._detail_set([
            ("Found", f.what),
            ("Fix", f.fix),
            ("Expected", f.expected),
        ], head="%s   %s   %s" % (f.code, f.severity, f.where or ""))

    def _detail_set(self, pairs, head=None):
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        if head:
            self.detail.insert("end", "%s\n\n" % head, "code")
        for label, body in pairs:
            if label:
                self.detail.insert("end", "%s\n" % label, "label")
            self.detail.insert("end", "%s\n\n" % body)
        self.detail.configure(state="disabled")


def _artefact_note(out):
    """What else the run wrote, for the banner. Absence is stated, not left."""
    if getattr(out, "export_error", ""):
        return ("  The field mapping document was NOT written: %s"
                % out.export_error)
    written = []
    if out.showcase_path:
        written.append("showcase report")
    if out.export_path:
        written.append("field mapping document")
    note = ""
    if written:
        note = "  Wrote the %s." % " and the ".join(written)
    if not out.export_path and out.split:
        note += ("  No field mapping document: it states which specification "
                 "it matches and a split run produces several.")
    return note


def make_root():
    """The Tk root, drag-and-drop aware when the bindings are installed.

    ``TkinterDnD.Tk`` is a plain ``tk.Tk`` with the drop protocol loaded. If
    loading it fails for any reason the ordinary root is used and the fields
    still work through Browse.
    """
    global DND
    if DND:
        try:
            return TkinterDnD.Tk()
        except Exception:                                      # noqa: BLE001
            DND = False
    return tk.Tk()


def main():
    root = make_root()
    App(root)
    if not DND:
        print("Drag and drop is unavailable: install tkinterdnd2 to enable it. "
              "The Browse buttons work either way.")
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
