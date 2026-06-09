"""Worker threads for the GUI.

All file work (PDF parsing, workbook writing) happens on a worker thread --
never on the Tk main thread. Workers communicate by putting messages on a
queue.Queue (thread-safe); the GUI drains it via root.after(). Workers never
touch Tk widgets.

Message protocol (all are (kind, payload) tuples):
    ("log", str)                       one status line
    ("consolidate_done", ConsolidateResult)
    ("check", (key, status, text))     one readiness-check result; status is 'ok'|'bad'
    ("checks_done", None)              all readiness checks posted
    ("error", str)                     unexpected failure; message is shown to the user
"""
import threading

from events import Events
from paths import OUTPUT_ROOT


class ConsolidateWorker(threading.Thread):
    """Runs one consolidator. Overwrite is resolved by the GUI before start,
    so the injected confirm callback just returns the pre-decided answer."""

    def __init__(self, consolidate_fn, queue, cancel_event, confirm, input_dir=None):
        super().__init__(daemon=True)
        self.consolidate_fn = consolidate_fn
        self.q = queue
        self.cancel = cancel_event
        self.confirm = confirm
        self.input_dir = input_dir          # None = the report's default folder

    def run(self):
        events = Events(
            on_log=lambda t: self.q.put(("log", t)),
            is_cancelled=self.cancel.is_set,
        )
        try:
            result = self.consolidate_fn(events=events, confirm_overwrite=self.confirm,
                                         input_dir=self.input_dir)
            self.q.put(("consolidate_done", result))
        except Exception as e:
            self.q.put(("error", f"{type(e).__name__}: {e}"))


# --- startup readiness checks -------------------------------------------------

def _check_output():
    try:
        OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
        probe = OUTPUT_ROOT / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return ("ok", "Output folder: writable")
    except Exception:
        return ("bad", "Output folder: NOT writable")


def _check_tools():
    try:
        import pdfplumber  # noqa: F401
        import openpyxl    # noqa: F401
        return ("ok", "Report tools (PDF/Excel): ready")
    except Exception as e:
        return ("bad", f"Report tools: missing ({type(e).__name__})")


class CheckWorker(threading.Thread):
    """Runs the launch-time readiness checks off the Tk thread, posting each
    result as ('check', (key, status, text)) and a final ('checks_done', None)."""

    def __init__(self, queue):
        super().__init__(daemon=True)
        self.q = queue

    def run(self):
        for key, fn in (("output", _check_output), ("tools", _check_tools)):
            try:
                status, text = fn()
            except Exception as e:
                status, text = "bad", f"{key}: error ({type(e).__name__})"
            self.q.put(("check", (key, status, text)))
        self.q.put(("checks_done", None))
