#!/usr/bin/env python3
"""
polymorphize_gui.py  —  SOR Polymorphizer (Windows GUI), v5.1
=============================================================

Order of actions is numbered on screen:

    STEP 1  Choose files   (input swagger + ONE SOR mapping spreadsheet;
                            optionally one OR MORE SOR API YAML files)
    STEP 2  Analyze        (reports the detected spreadsheet layout, the
                            mapping MATCH RATE, what would be eliminated,
                            what would be kept and flagged, the cascade, the
                            de-duplications and the polymorphism groups)
    STEP 3  Run transform  (approve the summarised changes; nothing is
                            written until you approve)

New in v5.0
-----------
* The spreadsheet layout is DETECTED (header row + columns by label) instead
  of assumed. The detected layout is shown in STEP 2 so a wrong sheet is
  obvious before anything is written.
* An attribute the mapping does not mention is KEPT AND FLAGGED by default.
  Eliminating it now requires ticking 'Eliminate attributes absent from the
  mapping', which is off deliberately.
* Objects emptied by pruning are removed along with the references to them,
  so payload examples never render as {}.
* De-duplication is decided on the ORIGINAL document and re-verified before
  it is applied, so two schemas that differed in the input are never merged.
* A run aborts rather than writing an emptied specification when the mapping
  match rate is below the floor or the attribute loss is above the ceiling.

New in v5.1
-----------
* Text encoding is handled explicitly. The input may be UTF-8, UTF-8 with a
  byte-order mark, or ANSI (cp1252) as saved by a Windows editor; the output
  and both reports are always written as UTF-8. This removes the
  UnicodeDecodeError that a hard-coded UTF-8 read raised on the em dash in the
  CASA description.
* Building a report can no longer discard a completed transform. A report
  failure is logged as a warning against a document that is already on disk.

Drag-and-drop needs the optional 'tkinterdnd2' package (bundled by the build
script); it falls back to Browse-only if absent.
"""
import os
import queue
import re
import threading
import traceback
import webbrowser

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

import polymorphize_core as core

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    _HAS_DND = False

__version__ = "5.1 (layout detection · keep-and-flag · cascade pruning · guards)"
APP_TITLE = f"SOR Polymorphizer  v{__version__}"


def _parse_drops(data):
    return [m[1:-1] if m.startswith("{") and m.endswith("}") else m
            for m in re.findall(r"\{[^}]*\}|\S+", data.strip())]


def _first_drop(data):
    parts = _parse_drops(data)
    return parts[0] if parts else ""


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("980x1010")
        root.minsize(860, 780)
        self.q = queue.Queue()
        self.sor_swaggers = []
        self.last_out = None
        self.last_html = None
        self.analysis = None
        self.analyzed_sig = None
        self._make_styles()
        self._build()
        self.root.after(100, self._drain_log)

    def _make_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("vista")
        except tk.TclError:
            pass
        st.configure("Step.TButton", font=("Segoe UI", 11, "bold"), padding=(10, 8))
        st.configure("Arrow.TLabel", font=("Segoe UI", 16, "bold"), foreground="#065A82")
        st.configure("StepHdr.TLabelframe.Label", font=("Segoe UI", 10, "bold"),
                     foreground="#21295C")

    # ---- single-file row --------------------------------------------------- #
    def _file_row(self, parent, label, r, save=False, kinds=None, invalidates=False):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=5)
        var = tk.StringVar()
        ent = ttk.Entry(parent, textvariable=var, width=62)
        ent.grid(row=r, column=1, sticky="we", padx=6, pady=5)
        ttk.Button(parent, text="Browse…",
                   command=lambda: self._browse(var, save, kinds)).grid(row=r, column=2,
                                                                       padx=6, pady=5)
        self._enable_single_drop(ent, var)
        if invalidates:
            var.trace_add("write", lambda *a: self._invalidate())
        return var

    def _enable_single_drop(self, widget, var):
        if not _HAS_DND or not hasattr(widget, "drop_target_register"):
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e, v=var: v.set(_first_drop(e.data)))
        except tk.TclError:
            pass

    def _browse(self, var, save, kinds):
        kinds = kinds or [("All files", "*.*")]
        if save:
            p = filedialog.asksaveasfilename(
                filetypes=kinds,
                defaultextension=kinds[0][1].split()[0].replace("*", ""))
        else:
            p = filedialog.askopenfilename(filetypes=kinds)
        if p:
            var.set(p)

    # ---- build ------------------------------------------------------------- #
    def _build(self):
        pad = dict(padx=6, pady=4)
        top = "Order:   1  choose files    ➜    2  Analyze    ➜    review boxes    ➜    3  Run  (approve)"
        ttk.Label(self.root, text=top, foreground="#21295C",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(10, 2))
        dhint = ("Drag files onto the boxes, or use Browse / Add."
                 if _HAS_DND else
                 "Install 'tkinterdnd2' for drag-and-drop; use Browse / Add for now.")
        ttk.Label(self.root, text=dhint, foreground="#5A6B78").pack(anchor="w", padx=16)

        yaml_kinds = [("YAML", "*.yaml *.yml"), ("All files", "*.*")]
        xlsx_kinds = [("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")]
        md_kinds = [("Markdown", "*.md"), ("All files", "*.*")]
        html_kinds = [("HTML", "*.html"), ("All files", "*.*")]

        files = ttk.LabelFrame(self.root, text="STEP 1 · Choose files",
                               style="StepHdr.TLabelframe")
        files.pack(fill="x", padx=10, pady=(4, 6))
        files.columnconfigure(1, weight=1)

        self.v_swagger = self._file_row(files, "Swagger / OpenAPI (input):", 0,
                                        kinds=yaml_kinds, invalidates=True)
        self.v_mapping = self._file_row(files, "SOR mapping spreadsheet:", 1,
                                        kinds=xlsx_kinds, invalidates=True)

        ttk.Label(files, text="SOR API YAML file(s) — optional:").grid(
            row=2, column=0, sticky="nw", padx=6, pady=5)
        sorwrap = ttk.Frame(files)
        sorwrap.grid(row=2, column=1, columnspan=2, sticky="we", padx=6, pady=5)
        self.lb_sor = tk.Listbox(sorwrap, height=3, activestyle="none")
        self.lb_sor.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(sorwrap, orient="vertical", command=self.lb_sor.yview)
        sb.pack(side="left", fill="y")
        self.lb_sor.config(yscrollcommand=sb.set)
        sbtns = ttk.Frame(sorwrap)
        sbtns.pack(side="left", padx=6)
        ttk.Button(sbtns, text="Add…", width=9, command=self._add_sor).pack(pady=(0, 3))
        ttk.Button(sbtns, text="Remove", width=9, command=self._remove_sor).pack()
        self._enable_list_drop(self.lb_sor)

        self.v_out = self._file_row(files, "Output swagger:", 3, save=True, kinds=yaml_kinds)
        self.v_report = self._file_row(files, "Markdown report (optional):", 4,
                                       save=True, kinds=md_kinds)
        self.v_html = self._file_row(files, "HTML change report (optional):", 5,
                                     save=True, kinds=html_kinds)

        opts = ttk.LabelFrame(self.root, text="Options")
        opts.pack(fill="x", padx=10, pady=6)
        self.v_drop_absent = tk.BooleanVar(value=False)
        self.v_disc = tk.BooleanVar(value=True)
        self.v_require_sor = tk.BooleanVar(value=False)
        self.v_auto_layout = tk.BooleanVar(value=True)
        self.v_guards = tk.BooleanVar(value=True)
        self.v_assert = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Eliminate attributes ABSENT from the mapping "
                                   "(off by default — 'not found' is not evidence of 'not used')",
                        variable=self.v_drop_absent).grid(row=0, column=0, columnspan=2,
                                                          sticky="w", **pad)
        self.v_drop_absent.trace_add("write", lambda *a: self._invalidate())
        ttk.Checkbutton(opts, text="Detect spreadsheet layout automatically",
                        variable=self.v_auto_layout).grid(row=1, column=0, sticky="w", **pad)
        self.v_auto_layout.trace_add("write", lambda *a: self._invalidate())
        ttk.Checkbutton(opts, text="Add discriminator to base schemas",
                        variable=self.v_disc).grid(row=1, column=1, sticky="w", **pad)
        ttk.Checkbutton(opts, text="Require SOR field to exist in SOR YAML (else eliminate)",
                        variable=self.v_require_sor).grid(row=2, column=0, columnspan=2,
                                                          sticky="w", **pad)
        self.v_require_sor.trace_add("write", lambda *a: self._invalidate())
        ttk.Checkbutton(opts, text="Abort on low match rate / excessive loss",
                        variable=self.v_guards).grid(row=3, column=0, sticky="w", **pad)
        ttk.Checkbutton(opts, text="Abort on failed output assertions",
                        variable=self.v_assert).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(opts, text="Minimum match rate:").grid(row=4, column=0, sticky="e", **pad)
        self.v_min_match = tk.DoubleVar(value=0.25)
        ttk.Entry(opts, textvariable=self.v_min_match, width=8).grid(row=4, column=1,
                                                                    sticky="w", **pad)
        ttk.Label(opts, text="Maximum total loss:").grid(row=5, column=0, sticky="e", **pad)
        self.v_max_loss = tk.DoubleVar(value=0.75)
        ttk.Entry(opts, textvariable=self.v_max_loss, width=8).grid(row=5, column=1,
                                                                   sticky="w", **pad)
        ttk.Label(opts, text="De-dup strictness:").grid(row=6, column=0, sticky="e", **pad)
        self.v_dedupe_strict = tk.StringVar(value="deep")
        ttk.Combobox(opts, textvariable=self.v_dedupe_strict, width=10, state="readonly",
                     values=("deep", "shape")).grid(row=6, column=1, sticky="w", **pad)
        ttk.Label(opts, text="Discriminator property (set by Analyze):").grid(
            row=7, column=0, sticky="e", **pad)
        self.v_disc_prop = tk.StringVar(value="")
        ttk.Entry(opts, textvariable=self.v_disc_prop, width=24).grid(row=7, column=1,
                                                                     sticky="w", **pad)
        ttk.Label(opts, text="Protected schemas (never pruned):").grid(row=8, column=0,
                                                                      sticky="e", **pad)
        self.v_protected = tk.StringVar(value="ErrorResponse")
        ttk.Entry(opts, textvariable=self.v_protected, width=36).grid(row=8, column=1,
                                                                     sticky="w", **pad)
        self.v_protected.trace_add("write", lambda *a: self._invalidate())

        adv = ttk.LabelFrame(self.root,
                             text="Advanced — fixed spreadsheet columns (used only when "
                                  "automatic detection is switched off)")
        adv.pack(fill="x", padx=10, pady=6)
        self.v_lf = tk.IntVar(value=core.DEFAULT_SHEET_CFG["level_first_col"])
        self.v_ll = tk.IntVar(value=core.DEFAULT_SHEET_CFG["level_last_col"])
        self.v_dt = tk.IntVar(value=core.DEFAULT_SHEET_CFG["dtype_col"])
        self.v_sor = tk.IntVar(value=core.DEFAULT_SHEET_CFG["sor_col"])
        self.v_nu = tk.StringVar(value=core.DEFAULT_SHEET_CFG["not_used_text"])
        for i, (lbl, var) in enumerate([("Level first col", self.v_lf),
                                        ("Level last col", self.v_ll),
                                        ("Data Type col", self.v_dt),
                                        ("SOR Field col", self.v_sor)]):
            ttk.Label(adv, text=lbl + ":").grid(row=0, column=i * 2, sticky="e", padx=4, pady=4)
            ttk.Entry(adv, textvariable=var, width=5).grid(row=0, column=i * 2 + 1,
                                                           sticky="w", padx=4, pady=4)
            var.trace_add("write", lambda *a: self._invalidate())
        ttk.Label(adv, text="Extra 'not used' text:").grid(row=1, column=0, sticky="e",
                                                           padx=4, pady=4)
        ttk.Entry(adv, textvariable=self.v_nu, width=24).grid(row=1, column=1, columnspan=3,
                                                              sticky="w", padx=4, pady=4)
        self.v_nu.trace_add("write", lambda *a: self._invalidate())
        ttk.Label(adv, text="(blank, N/A, TBD and '-' always count as 'no SOR field')",
                  foreground="#5A6B78").grid(row=1, column=4, columnspan=4, sticky="w",
                                             padx=4, pady=4)

        act = ttk.LabelFrame(self.root, text="STEP 2 → STEP 3 · Analyze, then Run",
                             style="StepHdr.TLabelframe")
        act.pack(fill="x", padx=10, pady=(8, 4))
        row = ttk.Frame(act)
        row.pack(fill="x", padx=8, pady=8)
        self.analyze_btn = ttk.Button(row, text="STEP 2:  Analyze file", style="Step.TButton",
                                      command=self._on_analyze)
        self.analyze_btn.pack(side="left")
        ttk.Label(row, text="  ➜  ", style="Arrow.TLabel").pack(side="left")
        self.run_btn = ttk.Button(row, text="STEP 3:  Run transform  (approve)",
                                  style="Step.TButton", command=self._on_run, state="disabled")
        self.run_btn.pack(side="left")
        self.status = tk.StringVar(
            value="Start at STEP 1: choose the input swagger and the SOR mapping spreadsheet.")
        ttk.Label(act, textvariable=self.status, foreground="#5A6B78").pack(anchor="w", padx=10,
                                                                           pady=(0, 6))

        cov = ttk.LabelFrame(self.root, text="Mapping coverage   (filled by Analyze — read this "
                                             "before approving a run)")
        cov.pack(fill="x", padx=10, pady=4)
        self.t_cov = tk.Text(cov, height=6, wrap="none")
        self.t_cov.pack(fill="x", padx=6, pady=6)
        self.t_cov.tag_config("good", foreground="#1E7A34")
        self.t_cov.tag_config("warn", foreground="#9A6A00")
        self.t_cov.tag_config("bad", foreground="#B3261E")

        groups = ttk.LabelFrame(self.root, text="Review · proposed polymorphism groups   "
                                                "(filled by Analyze · edit to override)")
        groups.pack(fill="x", padx=10, pady=4)
        self.t_groups = tk.Text(groups, height=3, wrap="none")
        self.t_groups.pack(fill="x", padx=6, pady=6)

        dd = ttk.LabelFrame(self.root, text="Review · proposed inline de-duplication   "
                                            "(filled by Analyze · edit to override)")
        dd.pack(fill="x", padx=10, pady=4)
        self.t_dedupe = tk.Text(dd, height=3, wrap="none")
        self.t_dedupe.pack(fill="x", padx=6, pady=6)

        util = ttk.Frame(self.root)
        util.pack(fill="x", padx=10, pady=(2, 2))
        self.open_html_btn = ttk.Button(util, text="Open change report",
                                        command=self._open_html, state="disabled")
        self.open_html_btn.pack(side="left")
        self.open_dir_btn = ttk.Button(util, text="Open output folder",
                                       command=self._open_dir, state="disabled")
        self.open_dir_btn.pack(side="left", padx=6)
        ttk.Button(util, text="Clear log",
                   command=lambda: self.log_box.delete("1.0", "end")).pack(side="left")

        logf = ttk.LabelFrame(self.root, text="Changes / log")
        logf.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log_box = scrolledtext.ScrolledText(logf, height=9, wrap="word",
                                                 font=("Consolas", 10))
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)
        self.log_box.tag_config("del", foreground="#B3261E")
        self.log_box.tag_config("add", foreground="#1E7A34")
        self.log_box.tag_config("warn", foreground="#9A6A00")
        self.log_box.tag_config("hdr", foreground="#21295C", font=("Consolas", 11, "bold"))

    # ---- SOR YAML list ----------------------------------------------------- #
    def _enable_list_drop(self, widget):
        if not _HAS_DND or not hasattr(widget, "drop_target_register"):
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: self._add_sor_paths(_parse_drops(e.data)))
        except tk.TclError:
            pass

    def _add_sor(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("YAML", "*.yaml *.yml"), ("All files", "*.*")])
        self._add_sor_paths(paths)

    def _add_sor_paths(self, paths):
        added = False
        for p in paths:
            p = p.strip()
            if p and p not in self.sor_swaggers:
                self.sor_swaggers.append(p)
                added = True
        if added:
            self._refresh_sor_list()
            self._invalidate()

    def _remove_sor(self):
        for i in reversed(self.lb_sor.curselection()):
            del self.sor_swaggers[i]
        self._refresh_sor_list()
        self._invalidate()

    def _refresh_sor_list(self):
        self.lb_sor.delete(0, "end")
        for p in self.sor_swaggers:
            self.lb_sor.insert("end", p)

    # ---- state ------------------------------------------------------------- #
    def _sig(self):
        return (self.v_swagger.get().strip(), self.v_mapping.get().strip(),
                tuple(self.sor_swaggers), self.v_require_sor.get(),
                self.v_drop_absent.get(), self.v_auto_layout.get(),
                self.v_protected.get().strip(), self.v_dedupe_strict.get(),
                self.v_lf.get(), self.v_ll.get(), self.v_dt.get(), self.v_sor.get(),
                self.v_nu.get())

    def _invalidate(self, *_):
        self.analysis = None
        self.analyzed_sig = None
        if hasattr(self, "run_btn"):
            self.run_btn.config(state="disabled")
            self.status.set("Inputs changed — click STEP 2 (Analyze) before running.")

    def _sheet_cfg(self):
        return {"level_first_col": self.v_lf.get(), "level_last_col": self.v_ll.get(),
                "dtype_col": self.v_dt.get(), "sor_col": self.v_sor.get(),
                "not_used_text": self.v_nu.get(),
                "auto_layout": bool(self.v_auto_layout.get())}

    def _protected(self):
        return tuple(s.strip() for s in self.v_protected.get().split(",") if s.strip())

    def _on_unmatched(self):
        return "drop" if self.v_drop_absent.get() else "keep"

    # ---- logging ----------------------------------------------------------- #
    def _log(self, msg, tag=None):
        self.q.put((msg, tag))

    def _drain_log(self):
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and len(item) == 2 and str(item[0]).startswith("__"):
                    self._handle_signal(item)
                    continue
                msg, tag = item if isinstance(item, tuple) else (item, None)
                self.log_box.insert("end", str(msg) + "\n", (tag,) if tag else ())
                self.log_box.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    def _handle_signal(self, item):
        kind, payload = item
        if kind == "__ANALYZED__":
            self._apply_analysis(payload)
        elif kind == "__ANALYZE_ERR__":
            self.analyze_btn.config(state="normal")
            messagebox.showerror("Analyze failed", "See the log for details.")
        elif kind == "__DONE__":
            self.last_out, self.last_html = payload
            self.open_dir_btn.config(state="normal")
            if self.last_html:
                self.open_html_btn.config(state="normal")
            self.run_btn.config(state="normal")
            messagebox.showinfo("Complete", f"Wrote:\n{self.last_out}")
        elif kind == "__ABORTED__":
            self.run_btn.config(state="normal")
            messagebox.showerror("Run aborted — nothing was written", str(payload))
        elif kind == "__ERROR__":
            self.run_btn.config(state="normal")
            messagebox.showerror("Failed", "See the log for details.")

    # ---- open helpers ------------------------------------------------------ #
    def _open_html(self):
        if self.last_html and os.path.isfile(self.last_html):
            webbrowser.open("file://" + os.path.abspath(self.last_html))

    def _open_dir(self):
        if self.last_out:
            folder = os.path.dirname(os.path.abspath(self.last_out))
            try:
                os.startfile(folder)
            except AttributeError:
                webbrowser.open("file://" + folder)

    # ---- STEP 2 : analyze -------------------------------------------------- #
    def _on_analyze(self):
        swagger = self.v_swagger.get().strip()
        mapping = self.v_mapping.get().strip()
        if not swagger or not os.path.isfile(swagger):
            messagebox.showerror("Missing input", "Choose a valid input swagger (STEP 1).")
            return
        if not mapping or not os.path.isfile(mapping):
            messagebox.showerror("Missing input",
                                 "Choose a valid SOR mapping spreadsheet (STEP 1).")
            return
        for y in self.sor_swaggers:
            if not os.path.isfile(y):
                messagebox.showerror("File not found", f"SOR YAML does not exist:\n{y}")
                return
        sig = self._sig()
        self.analyze_btn.config(state="disabled")
        self.status.set("Analyzing…")
        self._log("=" * 72)
        extra = f" · {len(self.sor_swaggers)} SOR YAML(s)" if self.sor_swaggers else ""
        self._log(f"STEP 2 — analyzing {os.path.basename(swagger)}{extra}…")
        sor_swaggers = list(self.sor_swaggers)
        require = self.v_require_sor.get()
        on_unmatched = self._on_unmatched()
        protected = self._protected()
        sheet_cfg = self._sheet_cfg()
        strict = self.v_dedupe_strict.get()

        def worker():
            try:
                res = core.analyze(swagger, mapping, sor_swaggers=sor_swaggers,
                                   require_sor_field=require, on_unmatched=on_unmatched,
                                   protected=protected, sheet_cfg=sheet_cfg,
                                   dedupe_strictness=strict)
                self.q.put(("__ANALYZED__", (sig, res)))
            except Exception:
                self._log("ERROR during analyze:\n" + traceback.format_exc(), "del")
                self.q.put(("__ANALYZE_ERR__", None))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_analysis(self, payload):
        sig, res = payload
        self.analysis = res
        self.analyzed_sig = sig
        diag = res["diagnostics"]
        rate = diag["match_rate"] * 100
        tag = "good" if rate >= 75 else ("warn" if rate >= 40 else "bad")

        self.t_cov.delete("1.0", "end")
        st = res["mapping_stats"]
        self.t_cov.insert("end", f"Mapping match rate: {rate:.1f}%  "
                                 f"({diag['attributes_matched']} of "
                                 f"{diag['attributes_total']} attributes)\n", tag)
        self.t_cov.insert("end", f"Mapping rows indexed: {st['attribute_rows']} from "
                                 f"{st['sheets']} sheet(s) — {st['mapped']} mapped, "
                                 f"{st['unmapped']} marked not-used\n")
        self.t_cov.insert("end", f"Resolution strategies: {diag['strategies']}\n")
        for lay in res["layouts"][:3]:
            cols = ", ".join(f"{k}=col{v}" for k, v in sorted(lay["cols"].items()))
            self.t_cov.insert("end", f"Layout [{lay['sheet']}] mode={lay['mode']} "
                                     f"header_row={lay['header_row']}  {cols}"
                                     + (f"  levels={lay['levels']}" if lay["levels"] else "")
                                     + "\n")
        if len(res["layouts"]) > 3:
            self.t_cov.insert("end", f"… {len(res['layouts']) - 3} further sheet(s) "
                                     f"with the same or similar layout\n")
        if rate < 40:
            self.t_cov.insert("end", "A low match rate means the LOOKUP failed, not that the "
                                     "SOR lacks the fields. Check the layout above.\n", "bad")

        self.t_groups.delete("1.0", "end")
        self.t_groups.insert("1.0", "\n".join(res["group_specs"]))
        self.t_dedupe.delete("1.0", "end")
        self.t_dedupe.insert("1.0", "\n".join(res["dedupe_specs"]))
        self.v_disc_prop.set(res.get("discriminator", ""))

        self._log("Proposal for these files:", "hdr")
        self._log(f"   mapping matched {diag['attributes_matched']}/{diag['attributes_total']} "
                  f"attribute(s) ({rate:.1f}%)",
                  "add" if rate >= 75 else ("warn" if rate >= 40 else "del"))
        self._log(f"   {len(res['eliminations'])} attribute(s) would be eliminated:", "hdr")
        for pathname, reason in res["eliminations"]:
            self._log(f"      − {pathname}   [{reason}]", "del")
        self._log(f"   {len(res['unmatched'])} attribute(s) would be KEPT and flagged "
                  f"(not found in the mapping):", "hdr")
        for pathname, tried in res["unmatched"][:200]:
            self._log(f"      ? {pathname}   [tried: {tried}]", "warn")
        casc = res["cascade"]
        self._log(f"   {len(casc['removed_schemas'])} emptied object(s) and "
                  f"{len(casc['removed_properties'])} reference(s) would be cascade-removed:",
                  "hdr")
        for n in casc["removed_schemas"]:
            self._log(f"      − object {n}", "del")
        if casc["empty_body_schemas"]:
            self._log("   WARNING — operation body schema(s) that would be empty: "
                      + ", ".join(casc["empty_body_schemas"]), "warn")
        self._log(f"   {len(res['group_specs'])} polymorphism group(s) proposed:", "hdr")
        for s in res["group_specs"]:
            self._log(f"      + {s}", "add")
        self._log(f"   {len(res['dedupe_specs'])} inline de-dup(s) proposed "
                  f"(equivalence taken from the ORIGINAL document):", "hdr")
        for s in res["dedupe_specs"]:
            self._log(f"      + {s}", "add")
        if res.get("discriminator"):
            self._log(f"   discriminator determined: {res['discriminator']}", "add")
        if res.get("assertions"):
            self._log(f"   {len(res['assertions'])} output assertion(s) would fail:", "hdr")
            for p in res["assertions"][:20]:
                self._log(f"      ! {p}", "del")
        val = res.get("validation")
        if val is not None:
            self._log(f"   SOR YAML validation: {val['field_count']} SOR fields · "
                      f"{len(val['confirmed'])} confirmed · {len(val['missing'])} NOT found:",
                      "hdr")
            for leaf, field in val["missing"][:100]:
                self._log(f"      ! {leaf} → {field}  (not in SOR YAML)", "warn")
        self._log("Review the coverage box first, then the boxes above, then click STEP 3.")
        status = (f"Analyzed: {rate:.1f}% matched · {len(res['eliminations'])} elim · "
                  f"{len(res['unmatched'])} kept-and-flagged · "
                  f"{len(casc['removed_schemas'])} cascade · "
                  f"{len(res['group_specs'])} groups · {len(res['dedupe_specs'])} de-dup")
        if val is not None:
            status += f" · SOR check: {len(val['missing'])} not found"
        self.status.set(status)
        self.analyze_btn.config(state="normal")
        self.run_btn.config(state="normal")

    # ---- STEP 3 : run ------------------------------------------------------ #
    def _collect_groups(self):
        return [core.parse_group_spec(l.strip())
                for l in self.t_groups.get("1.0", "end").splitlines() if l.strip()]

    def _collect_dedupe(self):
        return [core.parse_dedupe_spec(l.strip())
                for l in self.t_dedupe.get("1.0", "end").splitlines() if l.strip()]

    def _on_run(self):
        if self.analysis is None or self.analyzed_sig != self._sig():
            messagebox.showwarning("Analyze first",
                                   "The inputs changed since the last analysis.\n"
                                   "Click STEP 2 (Analyze), review, then STEP 3 (Run).")
            self._invalidate()
            return
        out = self.v_out.get().strip()
        if not out:
            messagebox.showerror("Missing output", "Choose an output swagger path (STEP 1).")
            return
        try:
            groups = self._collect_groups()
            dedupe = self._collect_dedupe()
        except Exception as ex:
            messagebox.showerror("Bad configuration", str(ex))
            return

        diag = self.analysis["diagnostics"]
        casc = self.analysis["cascade"]
        rate = diag["match_rate"] * 100
        val = self.analysis.get("validation")
        val_line = ""
        if val is not None:
            val_line = (f"   • SOR YAML check: {len(val['missing'])} mapped field(s) not "
                        f"found in the SOR\n")
        warn = ""
        if rate < 40:
            warn = ("\nWARNING: the mapping matched only "
                    f"{rate:.1f}% of the attributes. That normally means the wrong sheet or\n"
                    "the wrong field-name column, not that the SOR is missing the fields.\n")
        if not messagebox.askokcancel(
                "STEP 3 · Approve transform",
                "The following changes will be written:\n\n"
                f"   • Mapping matched {diag['attributes_matched']} of "
                f"{diag['attributes_total']} attributes ({rate:.1f}%)\n"
                f"   • Eliminate {len(self.analysis['eliminations'])} attribute(s) with a "
                f"positive no-SOR signal\n"
                f"   • Keep and flag {len(self.analysis['unmatched'])} attribute(s) absent "
                f"from the mapping\n"
                f"   • Cascade-remove {len(casc['removed_schemas'])} emptied object(s) and "
                f"{len(casc['removed_properties'])} reference(s)\n"
                f"   • Apply {len(groups)} polymorphism group(s)\n"
                f"   • De-duplicate {len(dedupe)} inline object(s)\n"
                f"{val_line}{warn}\n"
                f"Output:\n   {out}\n\nProceed?"):
            self._log("Run cancelled at approval.", "del")
            return

        html = self.v_html.get().strip() or (os.path.splitext(out)[0] + "_changes.html")
        self.v_html.set(html)
        report = self.v_report.get().strip() or None
        add_disc = self.v_disc.get()
        disc_prop = (self.v_disc_prop.get().strip()
                     or (self.analysis or {}).get("discriminator") or "context")
        protected = self._protected()
        sheet_cfg = self._sheet_cfg()
        swagger = self.v_swagger.get().strip()
        mapping = self.v_mapping.get().strip()
        sor_swaggers = list(self.sor_swaggers)
        require = self.v_require_sor.get()
        on_unmatched = self._on_unmatched()
        strict = self.v_dedupe_strict.get()
        try:
            min_match = float(self.v_min_match.get())
            max_loss = float(self.v_max_loss.get())
        except Exception:
            min_match, max_loss = 0.25, 0.75
        guards = bool(self.v_guards.get())
        do_assert = bool(self.v_assert.get())

        self.run_btn.config(state="disabled")
        self.open_html_btn.config(state="disabled")
        self.open_dir_btn.config(state="disabled")
        self._log("=" * 72)
        self._log("STEP 3 — approved. Running transform…")

        def worker():
            try:
                res = core.run(
                    swagger, mapping, out, report, html_report=html,
                    sor_swaggers=sor_swaggers, require_sor_field=require,
                    on_unmatched=on_unmatched, add_discriminator=add_disc,
                    discriminator_prop=disc_prop, poly_groups=groups, dedupe_pairs=dedupe,
                    auto_poly=False, auto_dedupe=False, dedupe_strictness=strict,
                    protected=protected, sheet_cfg=sheet_cfg,
                    min_match_rate=min_match, max_total_loss=max_loss,
                    enforce_guards=guards, enforce_assertions=do_assert, log=self._log)
                d = res["diagnostics"]
                self._log("")
                self._log(f"CHANGES: {len(res['dropped'])} eliminated, "
                          f"{d['attributes_unmatched']} kept-and-flagged, "
                          f"{len(res['cascade']['removed_schemas'])} objects cascade-removed, "
                          f"{len(res['poly_log'])} polymorphism action(s), "
                          f"{len(res['dedupe_log'])} de-dup(s)", "hdr")
                for rec in res["dropped"]:
                    self._log(f"   − {rec['path']}   [{rec['reason']}]", "del")
                for line in res["poly_log"]:
                    self._log(f"   + {line}", "add")
                for line in res["dedupe_log"]:
                    self._log(f"   + de-dup: {line}", "add")
                if res["assertions"]:
                    for p in res["assertions"]:
                        self._log(f"   ! assertion: {p}", "del")
                self._log("\nDone. Open 'change report' for the colour-coded diff.")
                self.q.put(("__DONE__", (out, html)))
            except (core.MappingCoverageError, core.ExcessiveLossError) as e:
                self._log(f"ABORTED (nothing written): {e}", "del")
                self.q.put(("__ABORTED__", e))
            except core.OutputAssertionError as e:
                self._log(f"ASSERTION FAILURE: {e}", "del")
                self.q.put(("__ABORTED__", e))
            except Exception:
                self._log("ERROR:\n" + traceback.format_exc(), "del")
                self.q.put(("__ERROR__", None))

        threading.Thread(target=worker, daemon=True).start()


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
