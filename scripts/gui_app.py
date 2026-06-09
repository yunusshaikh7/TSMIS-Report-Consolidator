"""Main window for the TSMIS Report Consolidator GUI.

Owns the Tk widgets and the queue pump. All file work (PDF parsing, workbook
writing) runs on worker threads (gui_worker); this module only reacts to the
messages they post, on the Tk main thread. The cores stay console-free -- the
GUI is just another driver of the same Events seam used by the .bat flow.
"""
import os
import sys
import threading
import time
from pathlib import Path
from queue import Empty, Queue

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import gui_theme as theme
from gui_theme import DOT, PALETTE
from gui_worker import CheckWorker, ConsolidateWorker

from paths import DATA_ROOT, LOG_DIR, OUTPUT_ROOT
from version import APP_NAME, __version__

# The report list lives in one place (reports.py). Each entry is
# (label, module) where the module provides consolidate(), INPUT_DIR,
# OUT_PATH, INPUT_GLOB, and REPORT_NAME.
from reports import CONSOLIDATE_REPORTS

CONSOLIDATED_DIR = OUTPUT_ROOT / "consolidated"

PAD = 14


def _app_icon_path():
    """Path to the bundled app icon (.ico), or None. Frozen: it's bundled into
    _internal via sys._MEIPASS; in dev it's build/app.ico. Best-effort -- a
    missing icon must never stop the GUI from launching."""
    base = getattr(sys, "_MEIPASS", None)
    candidates = []
    if base:
        candidates.append(Path(base) / "app.ico")
    candidates.append(Path(__file__).resolve().parent.parent / "build" / "app.ico")
    return next((c for c in candidates if c.exists()), None)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        # Window/taskbar icon (best-effort; default= so dialogs inherit it).
        try:
            _ico = _app_icon_path()
            if _ico:
                self.iconbitmap(default=str(_ico))
        except Exception:
            pass

        self.fonts = theme.fonts()
        theme.apply(self)

        self.q = Queue()
        self.task = None                       # None | "consolidate"
        self.cancel_event = threading.Event()

        self.cons_choice = tk.IntVar(value=0)
        self._run_start = None          # monotonic start of the current run (elapsed timer)
        self._timer_job = None          # after() id for the 1 Hz elapsed-time ticker
        self._check_detail = {}         # check key -> latest detail text (shown as a tooltip)
        self._tip = None                # the active tooltip Toplevel, if any

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)         # log row expands

        self._build_header()
        self._build_form()
        self._build_progress()
        self._build_log()
        self._build_footer()

        self._inputs = [self.btn_start, self.input_entry, self.btn_browse,
                        self.btn_recheck, *self.cons_radios]
        self.btn_cancel.state(["disabled"])

        # Pre-fill the input folder with the selected report's default and show
        # how many matching files it holds.
        self._set_default_input_dir()
        self._update_input_feedback()

        # Size to the laid-out content so EVERYTHING (incl. the log) shows at
        # launch, and keep that as the floor so the weighted log row can't be
        # squeezed to zero.
        self.update_idletasks()
        win_w = 640
        win_h = self.winfo_reqheight()
        self.geometry(f"{win_w}x{win_h}")
        self.minsize(580, win_h)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain)
        self.start_checks()                    # run the readiness checks on launch

    # ---- widget construction ------------------------------------------------

    def _build_header(self):
        h = ttk.Frame(self, style="Header.TFrame", padding=(PAD, 12))
        h.grid(row=0, column=0, sticky="ew")
        h.columnconfigure(0, weight=1)

        ttk.Label(h, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(h, text=f"v{__version__}", style="HeaderMuted.TLabel").grid(row=0, column=1, sticky="e")

        status = ttk.Frame(h, style="Header.TFrame")
        status.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.dot = ttk.Label(status, text="●", style="Dot.TLabel", foreground=DOT["unknown"])
        self.dot.grid(row=0, column=0, padx=(0, 6))
        self.status_text = ttk.Label(status, text="Idle", style="Status.TLabel")
        self.status_text.grid(row=0, column=1)

        # Compact readiness strip (output folder + report tools dots).
        strip = ttk.Frame(h, style="Header.TFrame")
        strip.grid(row=1, column=1, sticky="e", pady=(6, 0))
        self._check_items = {}
        for i, (key, label) in enumerate([("output", "Output"), ("tools", "Tools")]):
            dot = ttk.Label(strip, text="●", style="Dot.TLabel", foreground=DOT["unknown"])
            dot.grid(row=0, column=2 * i, padx=(14 if i else 0, 3))
            lab = ttk.Label(strip, text=label, style="HeaderMuted.TLabel")
            lab.grid(row=0, column=2 * i + 1)
            self._check_items[key] = (dot, label)
            self._check_detail[key] = f"{label}: checking…"
            self._attach_tip(dot, key)
            self._attach_tip(lab, key)
        self.btn_recheck = ttk.Button(strip, text="Re-check", width=9, command=self.start_checks)
        self.btn_recheck.grid(row=0, column=4, sticky="e", padx=(14, 0))

    # ---- tiny hover tooltip (for the compact check dots) --------------------

    def _attach_tip(self, widget, key):
        widget.bind("<Enter>", lambda _e, k=key, w=widget: self._show_tip(w, k))
        widget.bind("<Leave>", lambda _e: self._hide_tip())

    def _show_tip(self, widget, key):
        self._hide_tip()
        text = self._check_detail.get(key)
        if not text:
            return
        self._tip = tw = tk.Toplevel(self)
        tw.wm_overrideredirect(True)
        try:
            tw.attributes("-topmost", True)
        except tk.TclError:
            pass
        tw.wm_geometry(f"+{widget.winfo_rootx()}+{widget.winfo_rooty() + widget.winfo_height() + 3}")
        tk.Label(tw, text=text, bg=PALETTE["surface"], fg=PALETTE["text"],
                 relief="solid", borderwidth=1, padx=6, pady=2,
                 font=self.fonts["small"]).pack()

    def _hide_tip(self):
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None

    def _build_form(self):
        f = ttk.Frame(self, padding=(PAD, 10))
        f.grid(row=1, column=0, sticky="ew")
        f.columnconfigure(0, weight=1)
        row = 0

        ttk.Label(f, text="REPORT TO CONSOLIDATE (combine exported files)",
                  style="Section.TLabel").grid(row=row, column=0, sticky="w", pady=(0, 6))
        row += 1
        self.cons_radios = []
        for i, (label, _mod) in enumerate(CONSOLIDATE_REPORTS):
            rb = ttk.Radiobutton(f, text=label, value=i, variable=self.cons_choice,
                                 command=self._on_report_pick)
            rb.grid(row=row, column=0, sticky="w")
            self.cons_radios.append(rb)
            row += 1

        # Input folder: where the per-route exports live. Defaults to this
        # app's output/<report>/ but can point anywhere (e.g. the TSMIS Reports
        # Exporter's output folder on the same PC).
        ttk.Label(f, text="FOLDER WITH THE EXPORTED FILES", style="Section.TLabel").grid(
            row=row, column=0, sticky="w", pady=(10, 2))
        row += 1
        folder = ttk.Frame(f)
        folder.grid(row=row, column=0, sticky="ew")
        folder.columnconfigure(0, weight=1)
        self.input_entry = ttk.Entry(folder)
        self.input_entry.grid(row=0, column=0, sticky="ew")
        self.input_entry.bind("<KeyRelease>", self._update_input_feedback)
        self.btn_browse = ttk.Button(folder, text="Browse…", command=self._browse_input)
        self.btn_browse.grid(row=0, column=1, padx=(8, 0))
        row += 1
        self.input_feedback = ttk.Label(f, text="", style="Muted.TLabel",
                                        wraplength=560, justify="left")
        self.input_feedback.grid(row=row, column=0, sticky="w", pady=(4, 0))
        row += 1

        dest = ttk.Frame(f)
        dest.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        dest.columnconfigure(0, weight=1)
        ttk.Label(dest, text=f"Saved to:  {CONSOLIDATED_DIR}", style="Muted.TLabel",
                  wraplength=460, justify="left").grid(row=0, column=0, sticky="w")
        ttk.Button(dest, text="Open folder",
                   command=self._open_consolidated_folder).grid(row=0, column=1, sticky="e", padx=(8, 0))
        row += 1

        actions = ttk.Frame(f)
        actions.grid(row=row, column=0, sticky="w", pady=(PAD, 0))
        self.btn_start = ttk.Button(actions, text="Start consolidation", style="Accent.TButton",
                                    command=self.start_consolidate)
        self.btn_start.grid(row=0, column=0)
        self.btn_cancel = ttk.Button(actions, text="Cancel", command=self.cancel_current)
        self.btn_cancel.grid(row=0, column=1, padx=(8, 0))

    def _build_progress(self):
        f = ttk.Frame(self, padding=(PAD, 0))
        f.grid(row=2, column=0, sticky="ew")
        f.columnconfigure(0, weight=1)
        self.progress_label = ttk.Label(f, text="Idle", style="TLabel")
        self.progress_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(f, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.elapsed = ttk.Label(f, text="", style="Muted.TLabel")
        self.elapsed.grid(row=2, column=0, sticky="w", pady=(2, 0))

    def _build_log(self):
        f = ttk.Frame(self, padding=(PAD, 6))
        f.grid(row=3, column=0, sticky="nsew")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)
        self.log_widget = ScrolledText(f, height=10, wrap="word",
                                       bg=PALETTE["log_bg"], fg=PALETTE["log_fg"],
                                       relief="solid", borderwidth=1,
                                       font=self.fonts["mono"], padx=8, pady=6)
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        self.log_widget.configure(state="disabled")
        self.log_widget.tag_configure("error", foreground=PALETTE["danger"])
        self.log_widget.tag_configure("ok", foreground=PALETTE["success"])

    def _build_footer(self):
        f = ttk.Frame(self, padding=(PAD, 0, PAD, PAD))
        f.grid(row=4, column=0, sticky="ew")
        f.columnconfigure(0, weight=1)
        ttk.Label(f, text=f"All files are saved under:  {OUTPUT_ROOT}",
                  style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        btns = ttk.Frame(f)
        btns.grid(row=0, column=1, sticky="e")
        ttk.Button(btns, text="Open output folder",
                   command=self._open_output_folder).grid(row=0, column=0)
        ttk.Button(btns, text="Logs", command=self._open_logs_folder).grid(row=0, column=1, padx=(8, 0))

    # ---- small helpers ------------------------------------------------------

    def set_dot(self, state, text):
        self.dot.config(foreground=DOT[state])
        self.status_text.config(text=text)

    def log(self, text):
        tag = ""
        upper = text.upper()
        if "FAIL" in upper or "ERROR" in upper:
            tag = "error"
        elif "Output file" in text or "Files combined" in text or "Parsed:" in text:
            tag = "ok"
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", text + "\n", tag)
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _selected_report(self):
        """(label, module) of the report the radio buttons point at."""
        return CONSOLIDATE_REPORTS[self.cons_choice.get()]

    def _default_input_dirs(self):
        return {str(mod.INPUT_DIR) for _label, mod in CONSOLIDATE_REPORTS}

    def _set_default_input_dir(self):
        _label, mod = self._selected_report()
        self.input_entry.delete(0, "end")
        self.input_entry.insert(0, str(mod.INPUT_DIR))

    def _on_report_pick(self):
        # Follow the report's default folder unless the user typed/browsed a
        # custom one (then leave their choice alone). An empty entry also
        # re-arms the default.
        current = self.input_entry.get().strip()
        if not current or current in self._default_input_dirs():
            self._set_default_input_dir()
        self._update_input_feedback()

    def _update_input_feedback(self, *_):
        """Live hint under the folder entry: how many matching files are there."""
        _label, mod = self._selected_report()
        raw = self.input_entry.get().strip()
        if not raw:
            self.input_feedback.config(
                text="Leave as-is to use this app's output folder, or browse to "
                     "wherever the exported files are.")
            return
        folder = Path(raw)
        if not folder.is_dir():
            self.input_feedback.config(text="That folder doesn't exist (yet).")
            return
        n = len(list(folder.glob(mod.INPUT_GLOB)))
        kind = "PDF" if mod.INPUT_GLOB.endswith(".pdf") else "Excel"
        self.input_feedback.config(
            text=f"{n} {kind} file(s) found ({mod.INPUT_GLOB}).")

    def _browse_input(self):
        _label, mod = self._selected_report()
        start = self.input_entry.get().strip() or str(mod.INPUT_DIR)
        chosen = filedialog.askdirectory(title="Folder with the exported files",
                                         initialdir=start, mustexist=True)
        if chosen:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, chosen)
            self._update_input_feedback()

    # ---- elapsed-run timer (shown beneath the progress bar) ------------------

    @staticmethod
    def _fmt_elapsed(seconds):
        seconds = int(seconds)
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"

    def _start_timer(self):
        self._run_start = time.monotonic()
        self.elapsed.config(text="Elapsed  00:00")
        self._tick_timer()

    def _tick_timer(self):
        if self._run_start is None:
            return
        self.elapsed.config(text=f"Elapsed  {self._fmt_elapsed(time.monotonic() - self._run_start)}")
        self._timer_job = self.after(1000, self._tick_timer)   # re-schedule each second

    def _stop_timer(self):
        """Cancel the ticker and freeze the label on the final elapsed time."""
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        if self._run_start is not None:
            self.elapsed.config(text=f"Elapsed  {self._fmt_elapsed(time.monotonic() - self._run_start)}")
            self._run_start = None

    # ---- startup readiness checks -------------------------------------------

    def start_checks(self):
        if self.task:                          # never probe mid-run
            return
        for key, (dot, short) in self._check_items.items():
            dot.config(foreground=DOT["busy"])
            self._check_detail[key] = f"{short}: checking…"
        self.btn_recheck.state(["disabled"])
        CheckWorker(self.q).start()

    def _set_check(self, key, status, text=None):
        item = self._check_items.get(key)
        if not item:
            return
        dot, _short = item
        dot.config(foreground=DOT.get(status, DOT["unknown"]))
        if text:
            self._check_detail[key] = text          # shown as the hover tooltip
        if status == "bad" and text:
            self.log(f"Warning: {text}")

    # ---- run-state toggling -------------------------------------------------

    def _set_running(self):
        self.task = "consolidate"
        for w in self._inputs:
            w.state(["disabled"])
        self.btn_cancel.state(["!disabled"])
        self.progress_label.config(text="Working…")
        self.progress.start(12)
        self._start_timer()

    def _end_task(self):
        self._stop_timer()
        self.progress.stop()
        self.task = None
        for w in self._inputs:
            w.state(["!disabled"])
        self.btn_cancel.state(["disabled"])
        self.progress_label.config(text="Idle")

    # ---- actions ------------------------------------------------------------

    def start_consolidate(self):
        label, mod = self._selected_report()
        raw = self.input_entry.get().strip()
        input_dir = Path(raw) if raw else None
        out_path = mod.OUT_PATH
        if out_path.exists() and not messagebox.askyesno(
                "Overwrite?",
                f"A consolidated workbook already exists:\n\n{out_path}\n\nOverwrite it?"):
            self.log("Consolidation cancelled (kept existing file).")
            return
        self.cancel_event.clear()
        self.log(f"Starting consolidation: {label}")
        if input_dir is not None:
            self.log(f"Reading from: {input_dir}")
        self._set_running()
        self.set_dot("busy", f"Consolidating {label}…")
        ConsolidateWorker(mod.consolidate, self.q, self.cancel_event,
                          lambda _p: True, input_dir=input_dir).start()

    def cancel_current(self):
        if self.task == "consolidate":
            self.cancel_event.set()
            self.log("Cancel requested…")

    def _open_output_folder(self):
        self._open_folder(OUTPUT_ROOT)

    def _open_consolidated_folder(self):
        self._open_folder(CONSOLIDATED_DIR)

    def _open_logs_folder(self):
        self._open_folder(LOG_DIR)

    def _open_folder(self, folder):
        try:
            folder.mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))           # Windows
        except Exception as e:
            messagebox.showerror("Could not open folder", str(e))

    # ---- queue pump ---------------------------------------------------------

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                self._handle(kind, payload)
        except Empty:
            pass
        self.after(100, self._drain)

    def _handle(self, kind, payload):
        if kind == "log":
            self.log(payload)
        elif kind == "consolidate_done":
            self._finish_consolidate(payload)
        elif kind == "check":
            self._set_check(*payload)
        elif kind == "checks_done":
            if not self.task:
                self.btn_recheck.state(["!disabled"])
        elif kind == "error":
            self.log(f"ERROR: {payload}")
            self.set_dot("bad", "Error")
            messagebox.showerror("Error", f"{payload}\n\nMore details are in the log file.")
            self._end_task()

    def _finish_consolidate(self, result):
        if result.status == "ok":
            for line in result.summary_lines:
                self.log(line)
            self.set_dot("ok", "Done")
        elif result.status == "cancelled":
            self.log(result.message or "Cancelled.")
            self.set_dot("unknown", "Idle")
        else:
            self.log(f"ERROR: {result.message}")
            self.set_dot("bad", "Error")
            messagebox.showerror("Consolidation failed", result.message)
        self._end_task()

    def _on_close(self):
        # Unblock the worker so it can exit cleanly, then close.
        self.cancel_event.set()
        self._stop_timer()                     # cancel the pending ticker before teardown
        self._hide_tip()
        self.destroy()


# Surfaced for the GUI footer / About-style messages; DATA_ROOT matters when the
# packaged app fell back to %LOCALAPPDATA% (read-only install folder).
__all__ = ["App", "DATA_ROOT"]
