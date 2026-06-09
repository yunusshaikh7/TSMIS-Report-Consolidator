"""Lightweight, dependency-free seam between the consolidator cores and
whatever is driving them (the console shim or the GUI).

The cores never print, prompt, or exit: they push status through an Events
sink and return a ConsolidateResult. The console shim wires the callbacks to
print()/input(); the GUI wires them to a queue + widgets.
"""
from dataclasses import dataclass, field
from typing import List


def _noop_log(message):
    pass


def _never():
    return False


class Events:
    """Callbacks the cores use to report progress and check for control input.

    on_log:       human-readable status line (console prints it; GUI appends it
                  to a log pane).
    is_cancelled: return True to stop the run cleanly between files.

    Both default to harmless no-ops, so Events() is a valid silent sink.
    """

    def __init__(self, on_log=None, is_cancelled=None):
        self.on_log = on_log or _noop_log
        self.is_cancelled = is_cancelled or _never


@dataclass
class ConsolidateResult:
    """Outcome of one consolidation run.

    status is one of:
      "ok"        -- workbook written; summary_lines describe the result.
      "cancelled" -- user declined (e.g. the overwrite prompt) or cancelled
                     mid-run; message says why.
      "error"     -- could not complete; message explains it and is safe to
                     show to the user as-is.

    The core fills summary_lines with its own report-specific summary so the
    console shim and the GUI can display results without re-deriving them.
    """
    status: str = "ok"
    message: str = ""
    output_path: str = ""
    summary_lines: List[str] = field(default_factory=list)
