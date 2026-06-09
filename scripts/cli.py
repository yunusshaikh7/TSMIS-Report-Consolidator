"""Console adapter for the consolidator cores.

Wires the cores' Events callbacks to print()/input() and handles the console UX
(the overwrite prompt, exit codes) so the batch files work. The cores themselves
never touch the console -- only this module does.
"""
import sys

from events import Events
from logging_setup import setup_logging


def _confirm_overwrite_console(path):
    """Y/N overwrite prompt for the console flow. EOF (window closed) -> No,
    so double-clicking the BAT and closing it doesn't look like a crash."""
    print()
    print("A consolidated workbook already exists at:")
    print(f"   {path}")
    try:
        ans = input("Overwrite it? [Y/N]: ").strip().lower()
    except EOFError:
        return False
    return ans in ("y", "yes")


def run_consolidate_cli(consolidate_fn):
    """Run one consolidator as a console program. Used by the consolidate_*.py
    entry points and therefore by '2. consolidate (combine reports).bat'.

    The consolidator logs its own progress through Events.on_log and returns a
    ConsolidateResult; this shim renders the outcome and sets the exit code.
    """
    setup_logging()
    result = consolidate_fn(
        events=Events(on_log=print),
        confirm_overwrite=_confirm_overwrite_console,
    )

    if result.status == "cancelled":
        print(result.message or "Cancelled.")
        return
    if result.status == "error":
        print()
        print("=" * 60)
        print(f"ERROR: {result.message}")
        print("=" * 60)
        sys.exit(1)

    print()
    print("=" * 60)
    for line in result.summary_lines:
        print(line)
    print("=" * 60)
