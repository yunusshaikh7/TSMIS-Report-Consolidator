"""Application-wide file logging.

A single rotating log under LOG_DIR captures diagnostics from every entry point
(the GUI and the .bat console flow) so failures can be investigated after the
fact. File-only -- no console/stream handler -- so it never interferes with the
console flow's printed output or the windowed GUI. Call setup_logging() once at
process startup; it is safe to call again.
"""
import logging
from logging.handlers import RotatingFileHandler

from paths import LOG_DIR

LOG_FILE = LOG_DIR / "consolidator.log"
_configured = False


def setup_logging(level=logging.INFO):
    """Configure the root logger's rotating file handler once. Returns the log
    file path."""
    global _configured
    if _configured:
        return LOG_FILE
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(
        LOG_FILE, maxBytes=2_000_000, backupCount=5, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s: %(message)s", "%Y-%m-%d %H:%M:%S"))
    root = logging.getLogger()
    root.setLevel(level)
    root.addHandler(handler)
    _configured = True
    logging.getLogger("consolidator").info("=== logging started -> %s ===", LOG_FILE)
    return LOG_FILE
