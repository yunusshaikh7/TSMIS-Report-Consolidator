"""Centralized look-and-feel for the TSMIS Consolidator GUI.

One place that owns colors, fonts, and ttk styles so the whole app stays
consistent and the theme can be swapped without touching widget code. Built on
the `clam` ttk theme because it honors custom colors (the native Windows themes
ignore most background settings, which makes a branded look impossible).
"""
from tkinter import font as tkfont
from tkinter import ttk

# Professional light palette: dark slate header, white surfaces, blue accent.
PALETTE = {
    "bg":           "#F3F4F6",   # window background (light gray)
    "surface":      "#FFFFFF",   # cards / panels
    "header_bg":    "#1F2937",   # dark slate header band
    "header_fg":    "#FFFFFF",
    "header_muted": "#9CA3AF",
    "text":         "#111827",
    "muted":        "#6B7280",
    "border":       "#E5E7EB",
    "accent":       "#2563EB",   # primary blue
    "accent_dark":  "#1D4ED8",
    "disabled_bg":  "#9CA3AF",
    "success":      "#16A34A",
    "danger":       "#DC2626",
    "warning":      "#D97706",
    "log_bg":       "#FBFBFD",
    "log_fg":       "#111827",
}

# Status dot colors keyed by auth/run state.
DOT = {
    "ok":      PALETTE["success"],
    "bad":     PALETTE["danger"],
    "busy":    PALETTE["warning"],
    "unknown": PALETTE["muted"],
}

_FAMILY = "Segoe UI"
_FAMILY_SEMI = "Segoe UI Semibold"
_MONO = "Consolas"


def fonts():
    """Build the app's named fonts (call once, after a Tk root exists)."""
    return {
        "base":  tkfont.Font(family=_FAMILY, size=10),
        "bold":  tkfont.Font(family=_FAMILY, size=10, weight="bold"),
        "title": tkfont.Font(family=_FAMILY_SEMI, size=15),
        "small": tkfont.Font(family=_FAMILY, size=9),
        "mono":  tkfont.Font(family=_MONO, size=9),
    }


def apply(root):
    """Configure ttk styles on `root` and return the Style object."""
    p = PALETTE
    style = ttk.Style(root)
    style.theme_use("clam")
    root.configure(bg=p["bg"])

    style.configure(".", background=p["bg"], foreground=p["text"],
                    font=(_FAMILY, 10), focuscolor=p["bg"])

    # Frames
    style.configure("TFrame", background=p["bg"])
    style.configure("Surface.TFrame", background=p["surface"])
    style.configure("Header.TFrame", background=p["header_bg"])
    style.configure("Card.TLabelframe", background=p["surface"],
                    bordercolor=p["border"], relief="solid", borderwidth=1)
    style.configure("Card.TLabelframe.Label", background=p["bg"],
                    foreground=p["muted"], font=(_FAMILY, 9))

    # Labels
    style.configure("TLabel", background=p["bg"], foreground=p["text"])
    style.configure("Muted.TLabel", background=p["bg"], foreground=p["muted"],
                    font=(_FAMILY, 9))
    style.configure("Section.TLabel", background=p["bg"], foreground=p["muted"],
                    font=(_FAMILY, 9))
    style.configure("Header.TLabel", background=p["header_bg"], foreground=p["header_fg"])
    style.configure("HeaderMuted.TLabel", background=p["header_bg"],
                    foreground=p["header_muted"], font=(_FAMILY, 9))
    style.configure("Title.TLabel", background=p["header_bg"], foreground=p["header_fg"],
                    font=(_FAMILY_SEMI, 15))
    style.configure("Dot.TLabel", background=p["header_bg"], font=(_FAMILY, 11))
    style.configure("Status.TLabel", background=p["header_bg"], foreground=p["header_fg"],
                    font=(_FAMILY, 10))

    # Buttons
    style.configure("TButton", padding=(12, 6), background=p["surface"],
                    bordercolor=p["border"], relief="flat", foreground=p["text"])
    style.map("TButton",
              background=[("active", "#EEF1F5"), ("disabled", p["bg"])],
              foreground=[("disabled", p["muted"])])

    style.configure("Accent.TButton", padding=(14, 7), background=p["accent"],
                    foreground="#FFFFFF", relief="flat", bordercolor=p["accent"])
    style.map("Accent.TButton",
              background=[("active", p["accent_dark"]), ("disabled", p["disabled_bg"])],
              foreground=[("disabled", "#E5E7EB")])

    # Radiobuttons
    style.configure("TRadiobutton", background=p["bg"], foreground=p["text"],
                    padding=(2, 4))
    style.map("TRadiobutton",
              background=[("active", p["bg"])],
              foreground=[("disabled", p["muted"])])

    # Checkbuttons (report multi-select + fast-mode toggle) -- match radios so the
    # Export tab looks consistent instead of falling back to default clam styling.
    style.configure("TCheckbutton", background=p["bg"], foreground=p["text"],
                    padding=(2, 4))
    style.map("TCheckbutton",
              background=[("active", p["bg"])],
              foreground=[("disabled", p["muted"])])

    # Progress bar
    style.configure("TProgressbar", background=p["accent"], troughcolor=p["border"],
                    bordercolor=p["border"], lightcolor=p["accent"],
                    darkcolor=p["accent"], thickness=10)

    # Notebook (tabs)
    style.configure("TNotebook", background=p["bg"], borderwidth=0, tabmargins=(0, 6, 0, 0))
    style.configure("TNotebook.Tab", padding=(16, 8), font=(_FAMILY, 10),
                    background=p["bg"], foreground=p["muted"], borderwidth=0)
    style.map("TNotebook.Tab",
              background=[("selected", p["surface"])],
              foreground=[("selected", p["accent"]), ("active", p["text"])])

    # Separators
    style.configure("TSeparator", background=p["border"])
    return style
