#!/usr/bin/env python3
"""
Desktop GUI for converting spreadsheets (.xlsx/.xlsm/.xls/.ods) and
documents (.docx/.pptx/.odt/.pdf) to Markdown.

  - Add files with the buttons, a whole folder, or drag & drop (drag & drop
    needs the optional `tkinterdnd2` package)
  - Spreadsheets: pick the sheet, header row, output columns (in order),
    filter rules, sort order and table options. Every file keeps its own
    settings; a newly added file starts from the last ones you used.
  - Documents: converted whole, with options for headers / footers, speaker
    notes, formatting and a file-name heading
  - Live Markdown preview of the highlighted file
  - Convert the selected files or all of them, one .md per file or merged
    into a single .md
  - Light, dark or system (follows Windows) theme

Builds on md_converter_core, shared with md-converter.py.
"""

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from collections import OrderedDict
from copy import deepcopy
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import md_converter_core as core
import ui_theme

try:  # optional drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # not installed (or broken): the app works without it
    DND_FILES = TkinterDnD = None

APP_NAME = "MD Converter"
LEFT_WIDTH = 640            # starting width of the left (files / options) pane
PREVIEW_ROWS = 300          # spreadsheet rows rendered in the preview
PREVIEW_CHARS = 300_000     # document preview is cut after this many chars
HIGHLIGHT_LINES = 4000      # preview lines that get syntax colouring
OP_LABELS = dict(core.FILTER_OPS)
LABEL_TO_OP = {v: k for k, v in OP_LABELS.items()}
NO_SORT = "— no sorting —"  # first sort choice; unlikely to clash with a column name
SORT_ORDERS = ("A → Z", "Z → A")

FILE_DIALOG_TYPES = [
    ("Supported files", " ".join("*" + s for s in core.SUPPORTED_SUFFIXES)),
    ("Spreadsheets", " ".join("*" + s for s in core.SPREADSHEET_SUFFIXES)),
    ("Documents", " ".join("*" + s for s in core.DOCUMENT_SUFFIXES)),
    ("All files", "*.*"),
]

# --------------------------------------------------------------------------- #
#  Look
# --------------------------------------------------------------------------- #
# The colours of the current theme (ui_theme.LIGHT or DARK); widgets read them
# when they are made, and _apply_theme swaps them while the program runs.
C = dict(ui_theme.LIGHT)

DOC_OPTIONS = (
    ("include_header_footer", "Include headers & footers", False,
     "Word and ODT page headers/footers, PowerPoint footer boxes. "
     "For PDF, keeps lines repeated at the top/bottom of every page."),
    ("include_notes", "Include speaker notes", True,
     "PowerPoint notes, added under each slide as a quote."),
    ("keep_formatting", "Keep bold, italic & links", True,
     "Word and PowerPoint. Turn off for plain text."),
    ("title_from_name", "Add the file name as a top heading", False,
     "Starts the Markdown with '# <file name>'. Not used when merging."),
)


# --------------------------------------------------------------------------- #
#  Settings
# --------------------------------------------------------------------------- #
def _settings_path():
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home())
        return base / "MD-Converter" / "settings.json"
    return Path.home() / ".config" / "md-converter" / "settings.json"


def load_settings():
    try:
        data = json.loads(_settings_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data):
    try:
        path = _settings_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass  # settings are a convenience; never fail because of them


@dataclass
class SheetSettings:
    """Everything that decides how one spreadsheet becomes a table."""

    sheet: str = ""
    header_row: int = 1
    columns: list = field(default_factory=list)
    rules: list = field(default_factory=list)  # [{column, op, value}]
    match: str = "all"
    sort_by: str = ""
    descending: bool = False
    title: str = "Requirements"
    show_count: bool = True
    count_label: str = "Total requirements"
    keep_line_breaks: bool = False
    skip_empty_rows: bool = True
    row_numbers: bool = False

    @classmethod
    def from_dict(cls, data):
        if not isinstance(data, dict):
            return None
        known = {f.name for f in fields(cls)}
        try:
            settings = cls(**{k: v for k, v in data.items() if k in known})
            settings.header_row = max(1, int(settings.header_row))
        except (TypeError, ValueError):
            return None
        settings.columns = [str(c) for c in settings.columns or []]
        settings.rules = [dict(r) for r in settings.rules or [] if isinstance(r, dict)]
        return settings

    def table_options(self):
        return dict(
            rules=self.rules, match=self.match, sort_by=self.sort_by or None,
            descending=self.descending, title=self.title,
            show_count=self.show_count, count_label=self.count_label,
            keep_line_breaks=self.keep_line_breaks,
            skip_empty_rows=self.skip_empty_rows, row_numbers=self.row_numbers,
        )


def adapt_settings(settings, columns):
    """Keep only the parts of `settings` that fit `columns` (names matched
    case-insensitively). With no usable output column, fall back to the
    classic requirements layout: ID + Description, with an empty filter on
    Property ready to be filled in."""
    def match(name):
        return core.find_column(columns, name) if name else None

    settings.columns = [c for c in dict.fromkeys(match(c) for c in settings.columns) if c]
    settings.rules = [dict(r, column=match(r.get("column")))
                      for r in settings.rules if match(r.get("column"))]
    settings.sort_by = match(settings.sort_by) or ""
    if not settings.columns:
        settings.columns = [c for c in (match("ID"), match("Description")) if c]
        prop = match("Property")
        if settings.columns and prop and not settings.rules:
            settings.rules = [{"column": prop, "op": "word", "value": ""}]
    return settings


def initial_settings(template, sheets, read):
    """Settings for a spreadsheet opened for the first time, based on the
    last settings used (`template`). `read(sheet, header_row)` loads a
    sheet. Returns (settings, DataFrame).

    The template's header row suits the file it came from; when it does not
    fit this one (the sheet is shorter, or no wanted column is found there),
    header row 1 is used instead."""
    settings = deepcopy(template) if template else SheetSettings()
    settings.sheet = settings.sheet if settings.sheet in sheets else sheets[0]

    def attempt(header_row):
        try:
            df = read(settings.sheet, header_row)
        except Exception as exc:
            return None, None, exc
        trial = deepcopy(settings)
        trial.header_row = header_row
        return adapt_settings(trial, list(df.columns)), df, None

    result, df, error = attempt(settings.header_row)
    if settings.header_row != 1 and (result is None or not result.columns):
        fallback, df1, error1 = attempt(1)
        if fallback is not None and (result is None or fallback.columns):
            result, df, error = fallback, df1, None
    if result is None:
        raise error
    return result, df


@dataclass
class FileEntry:
    path: Path
    kind: str                      # "sheet" or "doc"
    iid: str = ""
    status: str = "Ready"
    sheets: list = None
    settings: SheetSettings = None
    frames: dict = field(default_factory=dict)  # (sheet, header_row) -> DataFrame
    frame_errors: dict = field(default_factory=dict)  # (sheet, header_row) -> message
    pending: set = field(default_factory=set)   # frame keys being read
    stash: dict = field(default_factory=dict)   # (sheet, header_row) -> earlier choices
    intent: dict = None            # the user's last own column/filter/sort choices
    generation: int = 0            # bumped on reload; stale callbacks compare it
    out_name: str = ""             # custom output name for this file
    loading: bool = False
    error: str = ""


# --------------------------------------------------------------------------- #
#  Small widgets
# --------------------------------------------------------------------------- #
class Tooltip:
    def __init__(self, widget, text, delay=550):
        self.widget, self.text, self.delay = widget, text, delay
        self._job = self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None

    def _show(self):
        self._job = None
        if self._tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 6
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._tip = tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{x}+{y}")
        tk.Label(tip, text=self.text, bg=C["tooltip"], fg="white", padx=8, pady=4,
                 justify="left", wraplength=320, font="TkDefaultFont").pack()

    def _hide(self, _event=None):
        self._cancel()
        if self._tip:
            self._tip.destroy()
            self._tip = None


class HintEntry(ttk.Entry):
    """ttk.Entry with a grey hint shown while it is empty and unfocused."""

    def __init__(self, master, hint="", textvariable=None, **kw):
        self.var = textvariable or tk.StringVar()
        super().__init__(master, textvariable=self.var, **kw)
        self._hint = tk.Label(master, text=hint, bg=C["field"], fg=C["faint"],
                              bd=0, padx=0, pady=0, cursor="xterm", font="TkDefaultFont")
        self._hint.bind("<Button-1>", lambda e: self.focus_set())
        self._traces = []
        self.on_change(self._update)
        for seq in ("<FocusIn>", "<FocusOut>", "<Map>", "<Configure>"):
            self.bind(seq, lambda e: self._update(), add="+")
        self.bind("<Unmap>", lambda e: self._hint.place_forget(), add="+")

    def on_change(self, callback):
        """Call `callback()` whenever the text changes. The trace is removed
        when the entry is destroyed, so rebuilt rows do not leak."""
        self._traces.append(self.var.trace_add("write", lambda *a: callback()))

    def destroy(self):
        for trace in self._traces:
            try:
                self.var.trace_remove("write", trace)
            except tk.TclError:
                pass
        self._traces = []
        self._hint.destroy()
        super().destroy()

    def set_hint(self, text):
        self._hint.configure(text=text)
        self._update()

    def get(self):
        return self.var.get()

    def set(self, value):
        self.var.set(value)

    def _update(self):
        try:
            show = (not self.var.get() and self.focus_get() is not self
                    and self.instate(["!disabled"]) and self.winfo_ismapped())
        except (KeyError, tk.TclError):  # focus_get on a destroyed popup
            show = False
        if show:
            self._hint.place(in_=self, x=7, rely=0.5, anchor="w")
        else:
            self._hint.place_forget()


class ScrollFrame(tk.Frame):
    """A vertically scrollable frame; put content in `.inner`."""

    def __init__(self, master, bg, increment=20):
        super().__init__(master, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                                yscrollincrement=increment)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>", lambda e: self._update_region())
        self.canvas.bind("<Configure>", self._on_resize)

    def _on_resize(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        self._update_region()

    def _update_region(self):
        height = max(self.inner.winfo_reqheight(), self.canvas.winfo_height())
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), height))

    def _on_scroll(self, first, last):
        self.vbar.set(first, last)
        fits = float(first) <= 0.0 and float(last) >= 1.0
        self.vbar.state(["disabled"] if fits else ["!disabled"])

    def scroll(self, units):
        if self.inner.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(units, "units")

    def see(self, widget):
        """Scroll so that `widget` (inside .inner) is visible."""
        self.update_idletasks()
        top = widget.winfo_rooty() - self.inner.winfo_rooty()
        bottom = top + widget.winfo_height()
        view_top = self.canvas.canvasy(0)
        view_bottom = view_top + self.canvas.winfo_height()
        total = max(self.inner.winfo_reqheight(), 1)
        if top < view_top:
            self.canvas.yview_moveto(top / total)
        elif bottom > view_bottom:
            self.canvas.yview_moveto((bottom - self.canvas.winfo_height()) / total)


def _enable_dpi_awareness():
    """Sharp text on high-DPI Windows screens instead of blurry scaling."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  App
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title(APP_NAME)
        self.saved = load_settings()
        self.theme_choice = self.saved.get("theme", "system")
        if self.theme_choice not in dict(ui_theme.CHOICES):
            self.theme_choice = "system"
        self.theme = ui_theme.resolve(self.theme_choice)
        ui_theme.apply_palette(C, self.theme)
        self.configure(bg=C["bg"])
        self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        self.files = OrderedDict()          # iid -> FileEntry
        self.focus_iid = None
        self.template = SheetSettings.from_dict(self.saved.get("sheet_template"))
        self.last_dir = self.saved.get("last_dir") or str(Path.home())
        self.last_output = None
        self.busy = False
        self._queue = queue.Queue()
        self._ui_lock = 0
        self._columns = []
        self._avail_items = []
        self._rule_rows = []
        self._preview_job = None
        self._preview_token = 0
        self._preview_md = ""
        self._doc_cache = OrderedDict()
        self._doc_preview_busy = False
        self._doc_preview_again = False
        self._header_job = None
        self._header_entry = None
        self._log_unseen = 0
        self._current_tab = "Preview"
        self._iid_counter = 0

        self._setup_fonts()
        self._setup_style()
        self._build()
        self._set_icon()
        self.dnd = self._setup_dnd()
        self._bind_keys()
        self._restore_window()
        self._refresh_all()
        self._report_missing_packages()
        self._pump_job = self.after(40, self._pump)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        ui_theme.dialog_options(self, C)
        ui_theme.set_title_bar(self, self.theme == "dark")
        self._theme_job = self.after(3000, self._follow_system_theme)
        self.deiconify()

    def _report_missing_packages(self):
        """Say up front which formats cannot be converted, instead of only
        failing file by file (e.g. an .exe built with the wrong Python)."""
        missing = core.missing_packages()
        if not missing:
            return
        names = ", ".join(pkg for pkg, _ in missing)
        formats = ", ".join(fmt for _, fmt in missing)
        if getattr(sys, "frozen", False):
            fix = ("This program was built without them: rebuild it with build-exe.bat "
                   "(see the README).")
        else:
            fix = "Install them with:  pip install " + " ".join(pkg for pkg, _ in missing)
        self._log(f"Missing packages: {names}. These formats cannot be converted: {formats}. {fix}",
                  "error")
        self._set_status(f"Some formats are unavailable (missing {names}). See the Log tab.", "warn")

    def px(self, n):
        return int(round(n * self.scale))

    # ================================================================ style
    def _setup_fonts(self):
        families = set(tkfont.families(self))

        def pick(*names):
            return next((n for n in names if n in families), None)

        ui = pick("Segoe UI", "SF Pro Text", "Helvetica Neue", "Inter", "Cantarell",
                  "Noto Sans", "DejaVu Sans") or "TkDefaultFont"
        mono = pick("Cascadia Mono", "Consolas", "Menlo", "SF Mono", "DejaVu Sans Mono",
                    "Liberation Mono", "Courier New") or "TkFixedFont"
        size = 13 if sys.platform == "darwin" else 10
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                     "TkTooltipFont"):
            try:
                tkfont.nametofont(name).configure(family=ui, size=size)
            except tk.TclError:
                pass
        self.f_base = (ui, size)
        self.f_small = (ui, size - 1)
        self.f_bold = (ui, size, "bold")
        self.f_section = (ui, size + 1, "bold")
        self.f_title = (ui, size + 5, "bold")
        self.f_mono = (mono, size)
        self.f_mono_bold = (mono, size, "bold")

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        px = self.px
        style.configure(
            ".", background=C["bg"], foreground=C["fg"], font=self.f_base,
            bordercolor=C["border"], lightcolor=C["card"], darkcolor=C["border"],
            troughcolor=C["bg"], focuscolor=C["accent"], selectbackground=C["accent"],
            selectforeground=C["on_accent"], insertcolor=C["fg"],
        )
        style.configure("TFrame", background=C["bg"])
        style.configure("Card.TFrame", background=C["card"])
        style.configure("TLabel", background=C["bg"], foreground=C["fg"])
        style.configure("Card.TLabel", background=C["card"])

        # Buttons: the 1px outer border is `bordercolor`; the bevel colours
        # match the face so the button looks flat.
        flat = dict(lightcolor=C["card"], darkcolor=C["card"])
        style.configure("TButton", background=C["card"], foreground=C["fg"],
                        bordercolor=C["border_strong"], padding=(px(12), px(5)),
                        anchor="center", focuscolor=C["card"], **flat)
        style.map("TButton",
                  background=[("disabled", C["bg"]), ("pressed", C["pressed"]), ("active", C["hover"])],
                  lightcolor=[("disabled", C["bg"]), ("pressed", C["pressed"]), ("active", C["hover"])],
                  darkcolor=[("disabled", C["bg"]), ("pressed", C["pressed"]), ("active", C["hover"])],
                  foreground=[("disabled", C["faint"])],
                  bordercolor=[("disabled", C["border"]), ("focus", C["accent"])],
                  focuscolor=[("pressed", C["pressed"]), ("active", C["hover"])])
        style.configure("Accent.TButton", background=C["accent"], foreground=C["on_accent"],
                        bordercolor=C["accent"], lightcolor=C["accent"], darkcolor=C["accent"],
                        font=self.f_bold, padding=(px(16), px(6)), focuscolor=C["accent"])
        style.map("Accent.TButton",
                  background=[("disabled", C["accent_disabled"]), ("pressed", C["accent_pressed"]),
                              ("active", C["accent_hover"])],
                  lightcolor=[("disabled", C["accent_disabled"]), ("pressed", C["accent_pressed"]),
                              ("active", C["accent_hover"])],
                  darkcolor=[("disabled", C["accent_disabled"]), ("pressed", C["accent_pressed"]),
                             ("active", C["accent_hover"])],
                  bordercolor=[("disabled", C["accent_disabled"]), ("pressed", C["accent_pressed"]),
                               ("active", C["accent_hover"])],
                  foreground=[("disabled", C["on_accent"])],
                  focuscolor=[("pressed", C["accent_pressed"]), ("active", C["accent_hover"])])
        style.configure("Small.TButton", padding=(px(8), px(3)))
        style.configure("Icon.TButton", padding=(px(4), px(3)), width=3)
        style.configure("Ghost.TButton", bordercolor=C["card"], padding=(px(6), px(3)),
                        foreground=C["muted"])
        style.map("Ghost.TButton", bordercolor=[("active", C["border"])],
                  foreground=[("active", C["danger"])])

        # Inputs
        style.configure("TEntry", fieldbackground=C["field"], bordercolor=C["border_strong"],
                        lightcolor=C["field"], darkcolor=C["field"], padding=(px(6), px(4)))
        style.map("TEntry",
                  bordercolor=[("focus", C["accent"])], lightcolor=[("focus", C["accent"])],
                  fieldbackground=[("disabled", C["bg"])], foreground=[("disabled", C["faint"])])
        style.configure("TCombobox", fieldbackground=C["field"], background=C["card"],
                        bordercolor=C["border_strong"], arrowcolor=C["muted"],
                        lightcolor=C["field"], darkcolor=C["field"], padding=(px(6), px(3)),
                        arrowsize=px(12))
        style.map("TCombobox",
                  fieldbackground=[("readonly", C["field"]), ("disabled", C["bg"])],
                  background=[("active", C["hover"]), ("readonly", C["card"])],
                  selectbackground=[("readonly", C["field"])],
                  selectforeground=[("readonly", C["fg"])],
                  foreground=[("disabled", C["faint"])],
                  bordercolor=[("focus", C["accent"]), ("active", C["accent"])],
                  lightcolor=[("focus", C["accent"])],
                  arrowcolor=[("disabled", C["border_strong"])])
        style.configure("TSpinbox", fieldbackground=C["field"], background=C["card"],
                        bordercolor=C["border_strong"], arrowcolor=C["muted"],
                        lightcolor=C["field"], darkcolor=C["field"], padding=(px(6), px(3)),
                        arrowsize=px(11))
        style.map("TSpinbox", bordercolor=[("focus", C["accent"])],
                  lightcolor=[("focus", C["accent"])], fieldbackground=[("disabled", C["bg"])])
        self.option_add("*TCombobox*Listbox.background", C["field"])
        self.option_add("*TCombobox*Listbox.foreground", C["fg"])
        self.option_add("*TCombobox*Listbox.selectBackground", C["accent_soft"])
        self.option_add("*TCombobox*Listbox.selectForeground", C["fg"])
        self.option_add("*TCombobox*Listbox.font", self.f_base)
        self.option_add("*TCombobox*Listbox.borderWidth", 0)

        # Checkboxes: drawn check-mark images instead of clam's "X" box.
        self._make_check_images()
        img = self._check_img
        try:
            style.element_create(
                "Check.indicator", "image", img["off"],
                ("disabled", "selected", img["on_disabled"]),
                ("disabled", img["off_disabled"]),
                ("selected", "active", img["on_hover"]),
                ("selected", img["on"]),
                ("active", img["off_hover"]),
                sticky="w",
            )
        except tk.TclError:  # already created (second App in one process)
            pass
        style.layout("TCheckbutton", [
            ("Checkbutton.padding", {"sticky": "nswe", "children": [
                ("Check.indicator", {"side": "left", "sticky": ""}),
                ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                    ("Checkbutton.label", {"sticky": "nswe"})]})]})])
        radio = self._radio_img
        try:
            style.element_create(
                "Radio.indicator", "image", radio["off"],
                ("disabled", "selected", radio["on_disabled"]),
                ("disabled", radio["off_disabled"]),
                ("selected", "active", radio["on_hover"]),
                ("selected", radio["on"]),
                ("active", radio["off_hover"]),
                sticky="w",
            )
        except tk.TclError:
            pass
        style.layout("TRadiobutton", [
            ("Radiobutton.padding", {"sticky": "nswe", "children": [
                ("Radio.indicator", {"side": "left", "sticky": ""}),
                ("Radiobutton.focus", {"side": "left", "sticky": "w", "children": [
                    ("Radiobutton.label", {"sticky": "nswe"})]})]})])

        for name in ("TCheckbutton", "TRadiobutton"):
            style.configure(name, background=C["card"], foreground=C["fg"],
                            indicatorbackground=C["field"], indicatorforeground="white",
                            upperbordercolor=C["border_strong"], lowerbordercolor=C["border_strong"],
                            indicatormargin=(0, 0, px(6), 0), focuscolor=C["card"],
                            padding=(0, px(2)))
            style.map(name,
                      background=[("active", C["card"])],
                      indicatorbackground=[("disabled", C["bg"]), ("selected", C["accent"]),
                                           ("pressed", C["accent_soft"])],
                      upperbordercolor=[("selected", C["accent"]), ("active", C["accent"])],
                      lowerbordercolor=[("selected", C["accent"]), ("active", C["accent"])],
                      foreground=[("disabled", C["faint"])])

        # File list
        style.configure("Treeview", background=C["field"], fieldbackground=C["field"],
                        foreground=C["fg"], rowheight=px(28), bordercolor=C["border"],
                        lightcolor=C["field"], darkcolor=C["field"], borderwidth=0)
        style.map("Treeview",
                  background=[("selected", "focus", C["accent_soft"]), ("selected", C["row_selected"])],
                  foreground=[("selected", C["fg"])])
        style.configure("Treeview.Heading", background=C["bg"], foreground=C["muted"],
                        font=self.f_small, relief="flat", bordercolor=C["border"],
                        lightcolor=C["bg"], darkcolor=C["bg"], padding=(px(6), px(4)))
        style.map("Treeview.Heading", background=[("active", C["pressed"])])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

        # Scrollbars and progress
        for orient in ("Vertical", "Horizontal"):
            style.configure(f"{orient}.TScrollbar", background=C["scroll"], troughcolor=C["card"],
                            bordercolor=C["card"], lightcolor=C["scroll"], darkcolor=C["scroll"],
                            arrowcolor=C["muted"], gripcount=0, arrowsize=px(12))
            style.map(f"{orient}.TScrollbar",
                      background=[("disabled", C["card"]), ("active", C["scroll_active"])],
                      lightcolor=[("disabled", C["card"]), ("active", C["scroll_active"])],
                      darkcolor=[("disabled", C["card"]), ("active", C["scroll_active"])],
                      arrowcolor=[("disabled", C["border"])])
        style.configure("Horizontal.TProgressbar", background=C["accent"], troughcolor=C["pressed"],
                        bordercolor=C["pressed"], lightcolor=C["accent"], darkcolor=C["accent"],
                        thickness=px(6))

    def _blank_image(self, kind, key, width, height):
        """A PhotoImage to draw on (blank pixels are transparent). When the
        theme changes the same image is drawn again, so the ttk elements made
        from it follow."""
        store = self.__dict__.setdefault("_images", {})
        img = store.get((kind, key))
        if img is None:
            img = store[(kind, key)] = tk.PhotoImage(width=width, height=height)
        else:
            img.blank()
        return img

    def _make_check_images(self):
        """Rounded check boxes (plus a gap before the label) as images."""
        s, gap = self.px(16), self.px(7)
        line = max(1, self.px(1))
        tick = max(2, self.px(2))

        def make(key, fill, border, check=None):
            img = self._blank_image("check", key, s + gap, s)
            img.put(border, to=(0, 0, s, s))
            img.put(fill, to=(line, line, s - line, s - line))
            for x, y in ((0, 0), (s - 1, 0), (0, s - 1), (s - 1, s - 1)):
                img.transparency_set(x, y, True)
            if check:
                points = [(0.23, 0.52), (0.42, 0.71), (0.78, 0.31)]
                for (x0, y0), (x1, y1) in zip(points, points[1:]):
                    steps = s * 2
                    for i in range(steps + 1):
                        x = (x0 + (x1 - x0) * i / steps) * s - tick / 2
                        y = (y0 + (y1 - y0) * i / steps) * s - tick / 2
                        img.put(check, to=(int(round(x)), int(round(y)),
                                           int(round(x)) + tick, int(round(y)) + tick))
            return img

        self._check_img = {
            "off": make("off", C["field"], C["border_strong"]),
            "off_hover": make("off_hover", C["field"], C["accent"]),
            "off_disabled": make("off_disabled", C["bg"], C["border"]),
            "on": make("on", C["accent"], C["accent"], "white"),
            "on_hover": make("on_hover", C["accent_hover"], C["accent_hover"], "white"),
            "on_disabled": make("on_disabled", C["accent_disabled"], C["accent_disabled"],
                                "white"),
        }

        def rgb(color):
            return tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))

        def blend(a, b, t):
            return tuple(round(x + (y - x) * t) for x, y in zip(a, b))

        def radio(key, fill, border, dot=None):
            """Anti-aliased (4x4 supersampled) round indicator."""
            img = self._blank_image("radio", key, s + gap, s)
            r, ring, dot_r = s / 2, line + 0.4, s * 0.2
            for y in range(s):
                for x in range(s):
                    outer = inner = centre = 0
                    for sy in range(4):
                        for sx in range(4):
                            dx, dy = x + (sx + 0.5) / 4 - r, y + (sy + 0.5) / 4 - r
                            d = (dx * dx + dy * dy) ** 0.5
                            outer += d <= r
                            inner += d <= r - ring
                            centre += bool(dot) and d <= dot_r
                    if not outer:
                        continue  # stays transparent
                    color = blend(rgb(C["card"]), rgb(border), outer / 16)
                    color = blend(color, rgb(fill), inner / 16)
                    if dot:
                        color = blend(color, rgb(dot), centre / 16)
                    img.put("#%02x%02x%02x" % color, to=(x, y, x + 1, y + 1))
            return img

        self._radio_img = {
            "off": radio("off", C["field"], C["border_strong"]),
            "off_hover": radio("off_hover", C["field"], C["accent"]),
            "off_disabled": radio("off_disabled", C["bg"], C["border"]),
            "on": radio("on", C["accent"], C["accent"], "#ffffff"),
            "on_hover": radio("on_hover", C["accent_hover"], C["accent_hover"], "#ffffff"),
            "on_disabled": radio("on_disabled", C["accent_disabled"], C["accent_disabled"],
                                 "#ffffff"),
        }

    def _set_icon(self):
        """A small accent-blue square with a white down arrow."""
        size = 32
        img = self._blank_image("icon", "app", size, size)
        img.put(C["accent"], to=(0, 0, size, size))
        for x, y in ((0, 0), (size - 1, 0), (0, size - 1), (size - 1, size - 1)):
            img.put(C["bg"], to=(x, y, x + 1, y + 1))
        img.put("white", to=(14, 6, 18, 18))
        for i in range(8):
            img.put("white", to=(8 + i, 16 + i, 24 - i, 17 + i))
        self._icon = img
        try:
            self.iconphoto(True, img)
        except tk.TclError:
            pass

    # ================================================================ layout
    def _build(self):
        # One frame holds the whole UI: it is the drag & drop target (the Tk
        # root itself cannot be one).
        self.shell = shell = tk.Frame(self, bg=C["bg"])
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(2, weight=1)
        self._build_header()
        tk.Frame(shell, bg=C["border"], height=1).grid(row=1, column=0, sticky="ew")

        self.paned = tk.PanedWindow(shell, orient="horizontal", bg=C["bg"], bd=0,
                                    sashwidth=self.px(6), sashrelief="flat", showhandle=False)
        self.paned.grid(row=2, column=0, sticky="nsew")
        self.left = ScrollFrame(self.paned, bg=C["bg"], increment=self.px(20))
        self.right = tk.Frame(self.paned, bg=C["bg"])
        self._left_min = self.px(440)
        self.paned.add(self.left, minsize=self._left_min, width=self.px(LEFT_WIDTH),
                       stretch="never")
        self.paned.add(self.right, minsize=self.px(380), stretch="always")
        self.paned.bind("<Configure>", self._keep_preview_visible, add="+")

        body = self.left.inner
        tk.Frame(body, bg=C["bg"], height=self.px(14)).pack(fill="x")
        self._build_files_card(body)
        self._build_options_card(body)
        self._build_output_card(body)
        self._build_preview(self.right)

        tk.Frame(shell, bg=C["border"], height=1).grid(row=3, column=0, sticky="ew")
        self._build_footer()

    def _keep_preview_visible(self, event=None):
        """When the window shrinks, move the divider left so the preview
        pane keeps its minimum width (Tk takes it from the last pane)."""
        width = self.paned.winfo_width()
        try:
            x = self.paned.sash_coord(0)[0]
        except tk.TclError:
            return
        limit = width - self.px(380) - self.px(6)
        if x > limit:
            self.paned.sash_place(0, max(self._left_min, limit), 1)

    def _fit_left_pane(self):
        """Keep the left pane at least as wide as its content needs (it
        changes with the options shown), so no button is ever cut off."""
        self.update_idletasks()
        need = max(self.px(440), self.left.inner.winfo_reqwidth()
                   + self.left.vbar.winfo_reqwidth() + self.px(4))
        if need != self._left_min:
            self._left_min = need
            self.paned.paneconfigure(self.left, minsize=need)
            self.minsize(max(self.px(860), need + self.px(380 + 6)), self.px(600))
        try:
            x = self.paned.sash_coord(0)[0]
        except tk.TclError:
            return
        if x < need:
            self.paned.sash_place(0, need, 1)

    def _label(self, parent, text="", muted=False, small=False, bold=False, **kw):
        font = self.f_bold if bold else self.f_small if small else self.f_base
        return tk.Label(parent, text=text, bg=kw.pop("bg", C["card"]),
                        fg=C["muted"] if muted else C["fg"], font=font, **kw)

    def _card(self, parent, title, pady=(0, 14)):
        px = self.px
        outer = tk.Frame(parent, bg=C["card"], highlightbackground=C["border"],
                         highlightcolor=C["border"], highlightthickness=1, bd=0)
        outer.pack(fill="x", padx=(px(16), px(8)), pady=(0, px(pady[1])))
        head = tk.Frame(outer, bg=C["card"])
        head.pack(fill="x", padx=px(16), pady=(px(12), 0))
        title_label = tk.Label(head, text=title, bg=C["card"], fg=C["fg"], font=self.f_section)
        title_label.pack(side="left")
        body = tk.Frame(outer, bg=C["card"])
        body.pack(fill="x", padx=px(16), pady=(px(10), px(16)))
        return outer, head, body, title_label

    def _subhead(self, parent, text, hint=None, top=14):
        row = tk.Frame(parent, bg=C["card"])
        row.pack(fill="x", pady=(self.px(top), self.px(6)))
        self._label(row, text, bold=True).pack(side="left")
        if hint:
            self._label(row, hint, muted=True, small=True).pack(side="left", padx=(self.px(8), 0))
        return row

    def _build_header(self):
        px = self.px
        header = tk.Frame(self.shell, bg=C["card"])
        header.grid(row=0, column=0, sticky="ew")
        logo = tk.Label(header, text="M↓", bg=C["accent"], fg=C["on_accent"],
                        font=(self.f_title[0], self.f_title[1] - 2, "bold"),
                        width=3, pady=px(2))
        logo.pack(side="left", padx=(px(20), px(12)), pady=px(12))
        texts = tk.Frame(header, bg=C["card"])
        texts.pack(side="left", pady=px(8))
        tk.Label(texts, text=APP_NAME, bg=C["card"], fg=C["fg"], font=self.f_title).pack(anchor="w")
        tk.Label(texts, text="Excel, Word, PowerPoint, OpenDocument and PDF files to clean Markdown",
                 bg=C["card"], fg=C["muted"], font=self.f_small).pack(anchor="w")
        help_btn = ttk.Button(header, text="Shortcuts", style="Small.TButton",
                              command=self._show_shortcuts, cursor="hand2")
        help_btn.pack(side="right", padx=px(20))
        self._build_theme_switch(header).pack(side="right")

    # ---------------------------------------------------------------- theme
    def _build_theme_switch(self, parent):
        """Light | Dark | System, as one segmented control."""
        px = self.px
        box = tk.Frame(parent, bg=C["card"])
        self._label(box, "Theme", muted=True, small=True).pack(side="left", padx=(0, px(8)))
        outline = tk.Frame(box, bg=C["border_strong"], padx=1, pady=1)
        outline.pack(side="left")
        self._theme_segments = {}
        tips = {"light": "Light theme", "dark": "Dark theme",
                "system": "Follow Windows: light or dark, as set for apps in the Windows "
                          "colour settings"}
        for i, (value, text) in enumerate(ui_theme.CHOICES):
            seg = tk.Label(outline, text=text, font=self.f_small, padx=px(11), pady=px(3),
                           cursor="hand2", bd=0)
            seg.pack(side="left", padx=(1 if i else 0, 0))   # the gap shows the outline
            seg.bind("<Button-1>", lambda e, v=value: self._set_theme(v))
            seg.bind("<Enter>", lambda e, v=value: self._paint_theme_switch(hover=v), add="+")
            seg.bind("<Leave>", lambda e: self._paint_theme_switch(), add="+")
            Tooltip(seg, tips[value])
            self._theme_segments[value] = seg
        self._paint_theme_switch()
        return box

    def _paint_theme_switch(self, hover=None):
        for value, seg in self._theme_segments.items():
            if value == self.theme_choice:
                seg.configure(bg=C["accent"], fg=C["on_accent"])
            elif value == hover:
                seg.configure(bg=C["hover"], fg=C["fg"])
            else:
                seg.configure(bg=C["card"], fg=C["muted"])

    def _set_theme(self, choice):
        """Light, Dark or System (follows Windows), switched at once."""
        self.theme_choice = choice
        self._apply_theme(ui_theme.resolve(choice))
        self._paint_theme_switch()
        self.saved["theme"] = choice
        save_settings(self.saved)
        name = dict(ui_theme.CHOICES)[choice]
        if choice == "system":
            name += f" (now {self.theme})"
        self._set_status(f"Theme: {name}.", "info")

    def _apply_theme(self, name):
        """Recolour the whole window with palette `name` ('light' or 'dark')."""
        if name == self.theme:
            return
        self.theme = name
        mapping = ui_theme.apply_palette(C, name)
        self._setup_style()                 # ttk styles and the check/radio pictures
        ui_theme.recolor(self, mapping)     # tk widgets, text tags, list rows, menus
        ui_theme.style_combobox_lists(self, C)
        ui_theme.dialog_options(self, C)
        ui_theme.set_title_bar(self, name == "dark")
        self._set_icon()
        self._paint_theme_switch()

    def _follow_system_theme(self):
        """In System mode, follow Windows when its setting changes."""
        self._theme_job = None
        try:
            if self.theme_choice == "system":
                name = ui_theme.system_theme()
                if name != self.theme:
                    self._apply_theme(name)
            self._theme_job = self.after(3000, self._follow_system_theme)
        except tk.TclError:
            pass    # closing

    # ---------------------------------------------------------------- files
    def _build_files_card(self, parent):
        px = self.px
        self.files_card, head, body, _ = self._card(parent, "Files")
        self.files_count = self._label(head, "", muted=True, small=True)
        self.files_count.pack(side="right")

        frame = tk.Frame(body, bg=C["field"], highlightbackground=C["border"],
                         highlightthickness=1)
        frame.pack(fill="x")
        self.tree = ttk.Treeview(frame, columns=("type", "status"), show="tree headings",
                                 selectmode="extended", height=5)
        self.tree.heading("#0", text="File", anchor="w")
        self.tree.heading("type", text="Type", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.column("#0", width=px(220), minwidth=px(120), stretch=True)
        self.tree.column("type", width=px(90), minwidth=px(70), stretch=False)
        self.tree.column("status", width=px(110), minwidth=px(80), stretch=False)
        self.tree.tag_configure("done", foreground=C["success"])
        self.tree.tag_configure("error", foreground=C["danger"])
        self.tree.tag_configure("busy", foreground=C["accent"])
        vbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_select())
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        self.tree.bind("<Button-3>", self._on_tree_menu)
        if sys.platform == "darwin":
            self.tree.bind("<Button-2>", self._on_tree_menu)

        self.tree_empty = tk.Label(frame, text="", bg=C["field"], fg=C["muted"],
                                   font=self.f_base, justify="center", cursor="hand2")
        self.tree_empty.bind("<Button-1>", lambda e: self._browse_files())

        btns = tk.Frame(body, bg=C["card"])
        btns.pack(fill="x", pady=(px(10), 0))
        add = ttk.Button(btns, text="Add files…", command=self._browse_files, cursor="hand2")
        add.pack(side="left")
        Tooltip(add, "Add spreadsheets or documents (Ctrl+O)")
        folder = ttk.Button(btns, text="Add folder…", command=self._browse_folder, cursor="hand2")
        folder.pack(side="left", padx=(px(8), 0))
        Tooltip(folder, "Add every supported file in a folder (Ctrl+Shift+O)")
        self.clear_btn = ttk.Button(btns, text="Clear", command=self._clear_files, cursor="hand2")
        self.clear_btn.pack(side="right")
        self.remove_btn = ttk.Button(btns, text="Remove", command=self._remove_selected,
                                     cursor="hand2")
        self.remove_btn.pack(side="right", padx=(0, px(8)))
        Tooltip(self.remove_btn, "Remove the selected files from the list (Delete)")

        self.tree_menu = tk.Menu(self, tearoff=False, bg=C["card"], fg=C["fg"],
                                 activebackground=C["accent_soft"], activeforeground=C["fg"],
                                 disabledforeground=C["faint"], relief="flat", bd=1)
        self.tree_menu.add_command(label="Convert", command=lambda: self._convert("selected"))
        self.tree_menu.add_command(label="Reload from disk", command=self._reload_focused)
        self.tree_menu.add_command(label="Show in folder", command=self._reveal_focused)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Remove", command=self._remove_selected)

    # ---------------------------------------------------------------- options
    def _build_options_card(self, parent):
        self.options_card, head, body, self.options_title = self._card(parent, "Options")
        self.options_sub = self._label(head, "", muted=True, small=True)
        self.options_sub.pack(side="right")
        self.options_body = body
        self.options_empty = self._label(
            body, "Add a file and select it to see its options.", muted=True)
        self.sheet_status = tk.Frame(body, bg=C["card"])
        self.sheet_status_label = self._label(self.sheet_status, "", muted=True,
                                              wraplength=self.px(420), justify="left")
        self.sheet_status_label.pack(anchor="w")
        self.sheet_retry = ttk.Button(self.sheet_status, text="Try again", style="Small.TButton",
                                      command=self._reload_focused)
        self._build_sheet_panel(body)
        self._build_doc_panel(body)

    def _build_sheet_panel(self, parent):
        px = self.px
        self.sheet_panel = panel = tk.Frame(parent, bg=C["card"])

        # Sheet + header row
        row = self.sheet_row = tk.Frame(panel, bg=C["card"])
        row.pack(fill="x")
        self.sheet_error = tk.Label(panel, text="", bg=C["card"], fg=C["danger"],
                                    font=self.f_small, wraplength=px(430), justify="left",
                                    anchor="w")
        self._label(row, "Sheet").pack(side="left")
        self.sheet_combo = ttk.Combobox(row, state="readonly", width=20)
        self.sheet_combo.pack(side="left", padx=(px(8), px(14)))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_selected())
        self._label(row, "Header row").pack(side="left")
        self.header_var = tk.StringVar(value="1")
        self.header_spin = ttk.Spinbox(row, from_=1, to=1000, width=4, textvariable=self.header_var,
                                       command=self._on_header_row)
        self.header_spin.pack(side="left", padx=(px(8), 0))
        self.header_var.trace_add("write", lambda *a: self._on_header_row_typed())
        Tooltip(self.header_spin, "Row number that holds the column names")
        reload_btn = ttk.Button(row, text="↻", style="Icon.TButton", command=self._reload_focused,
                                cursor="hand2")
        reload_btn.pack(side="right")
        Tooltip(reload_btn, "Reload the file from disk (F5)")

        # Columns
        self._subhead(panel, "Columns", "double-click to add or remove")
        grid = tk.Frame(panel, bg=C["card"])
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1, uniform="cols")
        grid.columnconfigure(2, weight=1, uniform="cols")
        self._label(grid, "Available", muted=True, small=True).grid(row=0, column=0, sticky="w")
        self.out_count = self._label(grid, "In the table, in order", muted=True, small=True)
        self.out_count.grid(row=0, column=2, sticky="w")
        self.col_search = HintEntry(grid, hint="Search columns…")
        self.col_search.grid(row=1, column=0, sticky="ew", pady=(px(4), px(4)))
        self.col_search.var.trace_add("write", lambda *a: self._refresh_column_lists())

        self.avail_list = self._listbox(grid, row=2, column=0)
        self.avail_list.bind("<Double-Button-1>", lambda e: self._add_columns())
        self.avail_list.bind("<Return>", lambda e: self._add_columns())

        mid = tk.Frame(grid, bg=C["card"])
        mid.grid(row=2, column=1, padx=px(8))
        for text, cmd, tip in (("→", self._add_columns, "Add the selected columns"),
                               ("←", self._remove_columns, "Remove the selected columns"),
                               ("»", self._add_all_columns, "Add all (shown) columns"),
                               ("«", self._clear_columns, "Remove all columns")):
            b = ttk.Button(mid, text=text, style="Icon.TButton", command=cmd, cursor="hand2")
            b.pack(pady=px(2))
            Tooltip(b, tip)

        self.out_list = self._listbox(grid, row=1, column=2, rowspan=2)
        self.out_list.bind("<Double-Button-1>", lambda e: self._remove_columns())
        self.out_list.bind("<Delete>", lambda e: self._remove_columns())

        side = tk.Frame(grid, bg=C["card"])
        side.grid(row=1, column=3, rowspan=2, padx=(px(8), 0))
        for text, delta, tip in (("▲", -1, "Move up"), ("▼", 1, "Move down")):
            b = ttk.Button(side, text=text, style="Icon.TButton", cursor="hand2",
                           command=lambda d=delta: self._move_columns(d))
            b.pack(pady=px(2))
            Tooltip(b, tip)

        # Filters
        head = self._subhead(panel, "Filter rows")
        self.add_rule_btn = ttk.Button(head, text="+ Add rule", style="Small.TButton",
                                       command=self._add_rule, cursor="hand2")
        self.add_rule_btn.pack(side="right")
        self.match_row = tk.Frame(panel, bg=C["card"])
        self._label(self.match_row, "Keep rows that match").pack(side="left")
        self.match_combo = ttk.Combobox(self.match_row, state="readonly", width=5,
                                        values=("all", "any"))
        self.match_combo.pack(side="left", padx=px(6))
        self.match_combo.bind("<<ComboboxSelected>>", lambda e: self._on_match())
        self._label(self.match_row, "of these rules").pack(side="left")
        self.rules_frame = tk.Frame(panel, bg=C["card"])
        self.rules_frame.pack(fill="x")
        self.rules_empty = self._label(panel, "No filter: every row is kept.", muted=True)

        # Table
        self.table_head = self._subhead(panel, "Table")
        grid = tk.Frame(panel, bg=C["card"])
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)

        self._label(grid, "Sort by").grid(row=0, column=0, sticky="w", pady=px(3))
        sort_row = tk.Frame(grid, bg=C["card"])
        sort_row.grid(row=0, column=1, sticky="ew", padx=(px(10), 0))
        self.sort_combo = ttk.Combobox(sort_row, state="readonly", width=20)
        self.sort_combo.pack(side="left", fill="x", expand=True)
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self._on_sort())
        self.order_combo = ttk.Combobox(sort_row, state="readonly", width=7, values=SORT_ORDERS)
        self.order_combo.pack(side="left", padx=(px(8), 0))
        self.order_combo.bind("<<ComboboxSelected>>", lambda e: self._on_sort())

        self._label(grid, "Title").grid(row=1, column=0, sticky="w", pady=px(3))
        self.title_var = tk.StringVar()
        self.title_entry = HintEntry(grid, hint="no heading", textvariable=self.title_var)
        self.title_entry.grid(row=1, column=1, sticky="ew", padx=(px(10), 0), pady=px(3))
        self.title_var.trace_add("write", lambda *a: self._on_table_option())

        self.count_var = tk.BooleanVar()
        count_cb = ttk.Checkbutton(grid, text="Count line", variable=self.count_var,
                                   command=self._on_table_option)
        count_cb.grid(row=2, column=0, sticky="w", pady=px(3))
        self.count_label_var = tk.StringVar()
        self.count_entry = HintEntry(grid, hint="Total rows", textvariable=self.count_label_var)
        self.count_entry.grid(row=2, column=1, sticky="ew", padx=(px(10), 0), pady=px(3))
        self.count_label_var.trace_add("write", lambda *a: self._on_table_option())

        checks = tk.Frame(panel, bg=C["card"])
        checks.pack(fill="x", pady=(px(6), 0))
        self.rownum_var = tk.BooleanVar()
        self.skip_empty_var = tk.BooleanVar()
        self.breaks_var = tk.BooleanVar()
        for var, text, tip in (
            (self.rownum_var, "# column", "Add a leading column numbering the rows 1, 2, 3…"),
            (self.skip_empty_var, "Skip empty rows", "Leave out rows whose table cells are all empty"),
            (self.breaks_var, "Line breaks as <br>",
             "Keep line breaks inside cells as <br> instead of joining them with spaces"),
        ):
            cb = ttk.Checkbutton(checks, text=text, variable=var, command=self._on_table_option)
            cb.pack(side="left", padx=(0, px(14)))
            Tooltip(cb, tip)

        self._label(panel, "Each spreadsheet keeps its own settings; newly added ones start "
                           "from the last settings you used.", muted=True, small=True,
                    wraplength=px(430), justify="left").pack(anchor="w", pady=(px(12), 0))

    def _listbox(self, parent, row, column, rowspan=1):
        frame = tk.Frame(parent, bg=C["field"], highlightthickness=1,
                         highlightbackground=C["border_strong"], highlightcolor=C["accent"])
        frame.grid(row=row, column=column, rowspan=rowspan, sticky="nsew")
        box = tk.Listbox(frame, height=8, selectmode="extended", exportselection=False,
                         bg=C["field"], fg=C["fg"], relief="flat", bd=0, highlightthickness=0,
                         selectbackground=C["accent_soft"], selectforeground=C["fg"],
                         activestyle="none", font=self.f_base)
        bar = ttk.Scrollbar(frame, orient="vertical", command=box.yview)
        box.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        box.pack(side="left", fill="both", expand=True, padx=(self.px(4), 0), pady=self.px(2))
        box.bind("<FocusIn>", lambda e: frame.configure(highlightbackground=C["accent"]))
        box.bind("<FocusOut>", lambda e: frame.configure(highlightbackground=C["border_strong"]))
        return box

    def _build_doc_panel(self, parent):
        px = self.px
        self.doc_panel = panel = tk.Frame(parent, bg=C["card"])
        saved = self.saved.get("doc") or {}
        self.doc_vars = {}
        for key, text, default, hint in DOC_OPTIONS:
            var = tk.BooleanVar(value=bool(saved.get(key, default)))
            self.doc_vars[key] = var
            ttk.Checkbutton(panel, text=text, variable=var,
                            command=self._on_doc_option).pack(anchor="w", pady=(px(4), 0))
            self._label(panel, hint, muted=True, small=True, wraplength=px(420),
                        justify="left").pack(anchor="w", padx=(px(26), 0))
        self._label(panel, "These options apply to every document in the list.", muted=True,
                    small=True).pack(anchor="w", pady=(px(14), 0))

    # ---------------------------------------------------------------- output
    def _build_output_card(self, parent):
        px = self.px
        _, _, body, _ = self._card(parent, "Output")
        body.columnconfigure(1, weight=1)

        self.out_mode = tk.StringVar(value=self.saved.get("out_mode", "source"))
        self._label(body, "Save to").grid(row=0, column=0, sticky="nw", pady=px(3))
        ttk.Radiobutton(body, text="Next to each source file", value="source",
                        variable=self.out_mode, command=self._on_output_option
                        ).grid(row=0, column=1, columnspan=2, sticky="w", padx=(px(10), 0))
        folder_row = tk.Frame(body, bg=C["card"])
        folder_row.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(px(10), 0), pady=(px(2), 0))
        ttk.Radiobutton(folder_row, text="Folder", value="folder", variable=self.out_mode,
                        command=self._on_output_option).pack(side="left")
        self.out_folder = HintEntry(folder_row, hint="choose a folder")
        self.out_folder.set(self.saved.get("out_folder", ""))
        self.out_folder.pack(side="left", fill="x", expand=True, padx=(px(8), px(8)))
        self.out_folder.var.trace_add("write", lambda *a: self._on_output_folder_typed())
        ttk.Button(folder_row, text="Browse…", command=self._browse_output_folder,
                   cursor="hand2").pack(side="left")

        self._label(body, "File name").grid(row=2, column=0, sticky="w", pady=(px(12), 0))
        name_row = tk.Frame(body, bg=C["card"])
        name_row.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(px(10), 0), pady=(px(12), 0))
        self.out_name = HintEntry(name_row)
        self.out_name.pack(side="left", fill="x", expand=True, padx=(0, px(8)))
        save_as = ttk.Button(name_row, text="Save as…", command=self._browse_output_file,
                             cursor="hand2")
        save_as.pack(side="left")
        Tooltip(save_as, "Pick the exact output file")
        self.out_name_hint = self._label(body, "", muted=True, small=True, justify="left",
                                         wraplength=px(400))
        self.out_name_hint.grid(row=3, column=1, columnspan=2, sticky="w", padx=(px(10), 0))

        self.merge_var = tk.BooleanVar(value=bool(self.saved.get("merge", False)))
        merge_row = tk.Frame(body, bg=C["card"])
        merge_row.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(px(10), 0), pady=(px(12), 0))
        ttk.Checkbutton(merge_row, text="Merge into one file", variable=self.merge_var,
                        command=self._on_output_option).pack(side="left")
        self.merged_name = HintEntry(merge_row, hint="merged.md", width=18)
        self.merged_name.set(self.saved.get("merged_name", ""))
        self.merged_name.pack(side="left", fill="x", expand=True, padx=(px(8), 0))
        self._label(body, "When converting several files, writes one .md with a section per file.",
                    muted=True, small=True, wraplength=px(400), justify="left"
                    ).grid(row=5, column=1, columnspan=2, sticky="w", padx=(px(10), 0))

        self.open_after_var = tk.BooleanVar(value=bool(self.saved.get("open_after", False)))
        ttk.Checkbutton(body, text="Open the folder when done", variable=self.open_after_var
                        ).grid(row=6, column=1, columnspan=2, sticky="w", padx=(px(10), 0),
                               pady=(px(10), 0))

    # ---------------------------------------------------------------- preview
    def _build_preview(self, parent):
        px = self.px
        card = tk.Frame(parent, bg=C["card"], highlightbackground=C["border"],
                        highlightthickness=1)
        card.pack(fill="both", expand=True, padx=(px(8), px(16)), pady=(px(14), px(14)))

        bar = tk.Frame(card, bg=C["card"])
        bar.pack(fill="x", padx=px(8))
        self._tabs = {}
        for name in ("Preview", "Log"):
            tab = tk.Frame(bar, bg=C["card"], cursor="hand2")
            tab.pack(side="left", padx=(px(6), px(6)))
            label = tk.Label(tab, text=name, bg=C["card"], fg=C["muted"], font=self.f_bold,
                             padx=px(4), pady=px(9), cursor="hand2")
            label.pack()
            line = tk.Frame(tab, bg=C["card"], height=px(2))
            line.pack(fill="x")
            for w in (tab, label):
                w.bind("<Button-1>", lambda e, n=name: self._select_tab(n))
            self._tabs[name] = (label, line)

        self.wrap_var = tk.BooleanVar(value=bool(self.saved.get("wrap", False)))
        self.copy_btn = ttk.Button(bar, text="Copy", style="Small.TButton",
                                   command=self._copy_preview, cursor="hand2")
        self.copy_btn.pack(side="right", pady=px(6))
        Tooltip(self.copy_btn, "Copy the full Markdown of this file")
        self.wrap_cb = ttk.Checkbutton(bar, text="Wrap", variable=self.wrap_var,
                                       command=self._apply_wrap)
        self.wrap_cb.pack(side="right", padx=px(10))
        self.clear_log_btn = ttk.Button(bar, text="Clear log", style="Small.TButton",
                                        command=self._clear_log, cursor="hand2")
        tk.Frame(card, bg=C["border"], height=1).pack(fill="x")

        info = tk.Frame(card, bg=C["card"])
        info.pack(fill="x", padx=px(14), pady=(px(6), px(6)))
        self.preview_name = tk.Label(info, text="", bg=C["card"], fg=C["fg"], font=self.f_bold)
        self.preview_name.pack(side="left")
        self.preview_stats = tk.Label(info, text="", bg=C["card"], fg=C["muted"], font=self.f_small)
        self.preview_stats.pack(side="left", padx=(px(10), 0))
        self.info_bar = info

        self.pages = tk.Frame(card, bg=C["card"])
        self.pages.pack(fill="both", expand=True)
        self.preview_page = self._text_page(self.pages)
        self.preview = self.preview_page.text
        self.preview_note = tk.Label(self.preview_page, text="", bg=C["accent_soft"], fg=C["fg"],
                                     font=self.f_small, anchor="w", padx=px(12), pady=px(4))
        self.log_page = self._text_page(self.pages)
        self.log = self.log_page.text
        self.log.configure(wrap="word")

        p = self.preview
        p.tag_configure("h", foreground=C["heading"], font=self.f_mono_bold)
        p.tag_configure("tsep", foreground=C["faint"])
        p.tag_configure("table", foreground=C["fg"])
        p.tag_configure("quote", foreground=C["muted"])
        p.tag_configure("meta", font=self.f_mono_bold)
        p.tag_configure("marker", foreground=C["accent"])
        p.tag_configure("rule", foreground=C["faint"])
        p.tag_configure("msg", foreground=C["muted"], font=self.f_base, lmargin1=px(6),
                        spacing1=px(10))
        p.tag_configure("err", foreground=C["danger"], font=self.f_base, lmargin1=px(6),
                        lmargin2=px(6), spacing1=px(10))
        self.log.tag_configure("time", foreground=C["faint"])
        self.log.tag_configure("ok", foreground=C["success"])
        self.log.tag_configure("error", foreground=C["danger"])
        self.log.tag_configure("info", foreground=C["fg"])
        self._apply_wrap()
        self._select_tab("Preview")

    def _text_page(self, parent):
        page = tk.Frame(parent, bg=C["code_bg"])
        text = tk.Text(page, bg=C["code_bg"], fg=C["fg"], font=self.f_mono, relief="flat", bd=0,
                       highlightthickness=0, padx=self.px(14), pady=self.px(10), wrap="none",
                       insertwidth=0, selectbackground=C["select"], selectforeground=C["fg"],
                       inactiveselectbackground=C["select_inactive"],
                       undo=False, cursor="arrow", spacing1=1, spacing3=1)
        vbar = ttk.Scrollbar(page, orient="vertical", command=text.yview)
        hbar = ttk.Scrollbar(page, orient="horizontal", command=text.xview)
        text.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set, state="disabled")
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        text.pack(side="left", fill="both", expand=True)
        page.text, page.hbar = text, hbar
        return page

    # ---------------------------------------------------------------- footer
    def _build_footer(self):
        px = self.px
        bar = tk.Frame(self.shell, bg=C["card"])
        bar.grid(row=4, column=0, sticky="ew")
        # Buttons are packed first so a long status text can never push
        # them out of the window.
        self.convert_sel_btn = ttk.Button(bar, text="Convert selected", style="Accent.TButton",
                                          command=lambda: self._convert("selected"), cursor="hand2")
        self.convert_sel_btn.pack(side="right", padx=(px(8), px(20)), pady=px(10))
        Tooltip(self.convert_sel_btn, "Convert the selected files (Ctrl+Enter)")
        self.convert_all_btn = ttk.Button(bar, text="Convert all", command=lambda: self._convert("all"),
                                          cursor="hand2")
        self.convert_all_btn.pack(side="right")
        Tooltip(self.convert_all_btn, "Convert every file in the list (Ctrl+Shift+Enter)")
        self.open_btn = ttk.Button(bar, text="Show in folder", command=self._open_last_output,
                                   cursor="hand2")
        self.open_btn.pack(side="right", padx=(0, px(8)))
        Tooltip(self.open_btn, "Open the folder of the last written file")

        self.status_dot = tk.Label(bar, text="●", bg=C["card"], fg=C["faint"], font=self.f_small)
        self.status_dot.pack(side="left", padx=(px(20), px(6)), pady=px(12))
        self.progress = ttk.Progressbar(bar, mode="determinate", length=px(160))
        self.status_label = tk.Label(bar, text="", bg=C["card"], fg=C["fg"], font=self.f_base,
                                     anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True, padx=(0, px(12)))

    # ================================================================ behaviour
    def _setup_dnd(self):
        if TkinterDnD is None:
            return False
        try:
            TkinterDnD._require(self)
            self.shell.drop_target_register(DND_FILES)
            self.shell.dnd_bind("<<Drop>>", self._on_drop)
            self.shell.dnd_bind("<<DropEnter>>", lambda e: self._drop_highlight(True))
            self.shell.dnd_bind("<<DropLeave>>", lambda e: self._drop_highlight(False))
            return True
        except Exception:
            return False

    def _drop_highlight(self, on):
        self.files_card.configure(highlightbackground=C["accent"] if on else C["border"],
                                  highlightthickness=2 if on else 1)
        return "copy"

    def _on_drop(self, event):
        self._drop_highlight(False)
        self._add_paths(self.tk.splitlist(event.data))
        return "copy"

    def _bind_keys(self):
        self.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.bind_all("<Button-4>", self._on_wheel, add="+")
        self.bind_all("<Button-5>", self._on_wheel, add="+")
        # A wheel turn over a combobox / spinbox would silently change its
        # value while scrolling the panel; only scroll the panel instead.
        for cls in ("TCombobox", "TSpinbox"):
            self.unbind_class(cls, "<MouseWheel>")
            self.unbind_class(cls, "<Button-4>")
            self.unbind_class(cls, "<Button-5>")
        self.bind("<Control-o>", lambda e: self._browse_files())
        self.bind("<Control-O>", lambda e: self._browse_folder())
        self.bind("<Control-Return>", lambda e: self._convert("selected"))
        self.bind("<Control-Shift-Return>", lambda e: self._convert("all"))
        self.bind("<F5>", lambda e: self._reload_focused())

    def _on_wheel(self, event):
        try:
            w = self.winfo_containing(event.x_root, event.y_root)
        except (KeyError, tk.TclError):  # e.g. over a combobox drop-down
            return None
        while w is not None:
            if isinstance(w, (tk.Listbox, tk.Text, ttk.Treeview)):
                return None  # these scroll themselves
            if w is self.left:
                if event.num in (4, 5):  # X11
                    units = -3 if event.num == 4 else 3
                elif sys.platform == "darwin":
                    units = -event.delta
                else:  # Windows: 120 per notch, less on precision touchpads
                    units = round(-3 * event.delta / 120) or (-1 if event.delta > 0 else 1)
                self.left.scroll(units)
                return "break"
            w = w.master
        return None

    # ---------------------------------------------------------------- tasks
    def _bg(self, work, done=None, fail=None):
        """Run `work()` on a worker thread; `done(result)` / `fail(exc)` run
        back on the Tk thread."""
        def run():
            try:
                result = work()
            except Exception as exc:
                self._queue.put((fail or self._bg_failed, exc))
            else:
                if done:
                    self._queue.put((done, result))
        threading.Thread(target=run, daemon=True).start()

    def _post(self, fn, *args):
        """Schedule `fn(*args)` on the Tk thread (safe from any thread)."""
        self._queue.put((lambda _arg: fn(*args), None))

    def _pump(self):
        try:
            while True:
                fn, arg = self._queue.get_nowait()
                try:
                    fn(arg)
                except Exception as exc:
                    self._log(f"Internal error: {core.describe_error(exc)}", "error")
        except queue.Empty:
            pass
        self._pump_job = self.after(40, self._pump)

    def _bg_failed(self, exc):
        self._log(core.describe_error(exc), "error")

    # ---------------------------------------------------------------- files
    def _browse_files(self):
        paths = filedialog.askopenfilenames(parent=self, title="Add files",
                                            initialdir=self.last_dir, filetypes=FILE_DIALOG_TYPES)
        if paths:
            self.last_dir = str(Path(paths[0]).parent)
            self._add_paths(paths)

    def _browse_folder(self):
        folder = filedialog.askdirectory(parent=self, title="Add every supported file in…",
                                         initialdir=self.last_dir)
        if folder:
            self.last_dir = folder
            self._add_paths([folder])

    def _add_paths(self, paths):
        added, skipped = [], []
        known = {os.path.normcase(os.path.abspath(e.path)) for e in self.files.values()}
        candidates = []
        for raw in paths:
            p = Path(os.path.abspath(raw))  # relative command-line paths too
            if p.is_dir():
                found = sorted((c for c in p.iterdir()
                                if c.is_file() and core.file_kind(c) and not c.name.startswith("~$")),
                               key=lambda c: c.name.lower())
                if not found:
                    skipped.append(p.name + " (no supported files)")
                candidates.extend(found)
            elif p.is_file():
                candidates.append(p)
            else:
                skipped.append(p.name + " (not found)")
        for p in candidates:
            kind = core.file_kind(p)
            if not kind:
                skipped.append(p.name)
                continue
            key = os.path.normcase(os.path.abspath(p))
            if key in known:
                continue
            known.add(key)
            self._iid_counter += 1
            iid = f"f{self._iid_counter}"
            entry = FileEntry(path=p, kind=kind, iid=iid)
            self.files[iid] = entry
            self.tree.insert("", "end", iid=iid, text=p.name,
                             values=(core.FILE_TYPES.get(p.suffix.lower(), p.suffix), "Ready"))
            added.append(iid)
        if added:
            self.tree.selection_set(added[0])
            self.tree.focus(added[0])
            self.tree.see(added[0])
            self._set_status(f"Added {len(added)} file{'s' if len(added) != 1 else ''}.", "info")
        if skipped:
            shown = ", ".join(skipped[:4]) + (" …" if len(skipped) > 4 else "")
            self._set_status(f"Skipped {len(skipped)} unsupported item(s): {shown}", "warn")
            self._log(f"Skipped: {', '.join(skipped)}", "error")
        self._refresh_all()

    def _remove_selected(self):
        if self.busy:
            return
        selected = list(self.tree.selection())
        if not selected:
            return
        index = self.tree.index(selected[0])
        for iid in selected:
            self.tree.delete(iid)
            self.files.pop(iid, None)
        remaining = self.tree.get_children()
        if remaining:
            nxt = remaining[min(index, len(remaining) - 1)]
            self.tree.selection_set(nxt)
            self.tree.focus(nxt)
        else:
            self.focus_iid = None
        self._refresh_all()

    def _clear_files(self):
        if self.busy or not self.files:
            return
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self._doc_cache.clear()
        self.focus_iid = None
        self._set_status("List cleared.", "info")
        self._refresh_all()

    def _on_select(self):
        selection = self.tree.selection()
        focus = self.tree.focus()
        if focus not in selection:
            focus = selection[0] if selection else None
        if focus != self.focus_iid:
            self._flush_header_row()  # a header row typed for the old file
            old = self._focused()
            if old:
                old.out_name = self.out_name.get().strip()
            self.focus_iid = focus
            new = self._focused()
            self.out_name.set(new.out_name if new else "")
            self._show_options()
            self._schedule_preview(0)
        self._update_controls()

    def _on_tree_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        self.tree.focus(iid)
        try:
            self.tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_menu.grab_release()

    def _focused(self):
        return self.files.get(self.focus_iid)

    def _set_file_status(self, entry, status, tag=None):
        entry.status = status
        if self.tree.exists(entry.iid):
            self.tree.set(entry.iid, "status", status)
            self.tree.item(entry.iid, tags=(tag,) if tag else ())

    def _reload_focused(self):
        entry = self._focused()
        if not entry:
            return
        if entry.kind == "sheet":
            # Callbacks of reads still running for the old state are ignored.
            entry.generation += 1
            entry.frames.clear()
            entry.frame_errors.clear()
            entry.pending.clear()
            entry.loading = False
            entry.sheets = None
            entry.error = ""
            keep = entry.settings  # keep the user's choices across a reload
            entry.settings = None
            self._load_sheet(entry, keep)
        else:
            for key in [k for k in self._doc_cache if k[0] == str(entry.path)]:
                self._doc_cache.pop(key, None)
        self._show_options()
        self._schedule_preview(0)

    def _reveal_focused(self):
        entry = self._focused()
        if entry:
            self._open_location(entry.path)

    # ---------------------------------------------------------------- sheets
    def _load_sheet(self, entry, keep=None):
        """Read sheet names and the first sheet of a spreadsheet in the
        background; `keep` reuses existing settings instead of the template."""
        if entry.loading:
            return
        entry.loading = True
        entry.error = ""
        self._set_file_status(entry, "Loading…", "busy")
        base = deepcopy(keep) if keep else deepcopy(self.template)
        path = entry.path
        generation = entry.generation

        def work():
            sheets = core.sheet_names(path)
            if not sheets:
                raise ValueError("the workbook has no sheets")
            frames = {}

            def read(sheet, header_row):
                key = (sheet, header_row)
                if key not in frames:
                    frames[key] = core.read_excel(path, sheet, header_row)
                return frames[key]

            settings, _ = initial_settings(base, sheets, read)
            return sheets, settings, frames

        def done(result):
            if entry.generation != generation:
                return
            entry.loading = False
            entry.sheets, entry.settings, frames = result
            entry.frames.update(frames)
            if entry.intent is None:
                entry.intent = self._choices(entry.settings)
            self._set_file_status(entry, "Ready")
            if entry.iid == self.focus_iid:
                self._show_options()
                self._schedule_preview(0)

        def fail(exc):
            if entry.generation != generation:
                return
            entry.loading = False
            entry.error = core.describe_error(exc)
            self._set_file_status(entry, "Error", "error")
            self._log(f"{entry.path.name}: {entry.error}", "error")
            if entry.iid == self.focus_iid:
                self._show_options()
                self._schedule_preview(0)

        self._bg(work, done, fail)

    def _sheet_frame(self, entry):
        s = entry.settings
        return entry.frames.get((s.sheet, s.header_row)) if s else None

    def _with_frame(self, entry, then):
        """Call `then(df)` once the current sheet / header row is loaded. A
        failed read is recorded in `entry.frame_errors`; the options panel
        then shows it next to the sheet / header row controls."""
        s = entry.settings
        key = (s.sheet, s.header_row)
        if key in entry.frames:
            then(entry.frames[key])
            return
        if key in entry.frame_errors:
            self._refresh_sheet_view(entry)
            return
        if entry.iid == self.focus_iid:
            self._show_message("Loading sheet…")
        if key in entry.pending:  # already being read; its callback refreshes the UI
            return
        entry.pending.add(key)
        generation = entry.generation
        self._set_file_status(entry, "Loading…", "busy")

        def current():
            return (entry.generation == generation and entry.settings is s
                    and (s.sheet, s.header_row) == key)

        def done(df):
            if entry.generation != generation:
                return
            entry.pending.discard(key)
            entry.frames[key] = df
            self._set_file_status(entry, "Ready")
            if current():
                then(df)

        def fail(exc):
            if entry.generation != generation:
                return
            entry.pending.discard(key)
            msg = core.describe_error(exc)
            entry.frame_errors[key] = msg
            self._set_file_status(entry, "Error", "error")
            self._log(f"{entry.path.name} (sheet {key[0]}, header row {key[1]}): {msg}", "error")
            if current():
                self._refresh_sheet_view(entry)

        self._bg(lambda: core.read_excel(entry.path, *key), done, fail)

    def _refresh_sheet_view(self, entry):
        if entry.iid == self.focus_iid:
            self._show_options()
            self._schedule_preview(0)

    def _show_options(self):
        try:
            self._show_options_panel()
        finally:
            self._fit_left_pane()

    def _show_options_panel(self):
        entry = self._focused()
        for w in (self.options_empty, self.sheet_status, self.sheet_panel, self.doc_panel):
            w.pack_forget()
        self.sheet_retry.pack_forget()
        if entry is None:
            self.options_title.configure(text="Options")
            self.options_sub.configure(text="")
            self.options_empty.pack(anchor="w")
            return
        if entry.kind == "doc":
            self.options_title.configure(text="Document options")
            self.options_sub.configure(text="all documents")
            self.doc_panel.pack(fill="x")
            return
        self.options_title.configure(text="Spreadsheet options")
        self.options_sub.configure(text=entry.path.name)
        if entry.error:
            self.sheet_status_label.configure(text=f"Could not open this file.\n{entry.error}",
                                              fg=C["danger"])
            self.sheet_status.pack(fill="x")
            self.sheet_retry.pack(anchor="w", pady=(self.px(8), 0))
            return
        s = entry.settings
        if s is None or (self._sheet_frame(entry) is None
                         and (s.sheet, s.header_row) not in entry.frame_errors):
            self.sheet_status_label.configure(text="Loading spreadsheet…", fg=C["muted"])
            self.sheet_status.pack(fill="x")
            if s is None:
                self._load_sheet(entry)
            else:
                self._with_frame(entry, lambda df: self._refresh_sheet_view(entry))
            return
        # Loaded, or this sheet / header row failed: the panel stays usable
        # so the header row or sheet can be corrected.
        self.sheet_panel.pack(fill="x")
        self._populate_sheet_panel(entry)

    def _populate_sheet_panel(self, entry):
        s = entry.settings
        df = self._sheet_frame(entry)
        self._columns = list(df.columns) if df is not None else []
        error = entry.frame_errors.get((s.sheet, s.header_row))
        if error:
            self.sheet_error.configure(
                text=f"Could not read sheet “{s.sheet}” with header row {s.header_row}: {error}")
            self.sheet_error.pack(fill="x", pady=(self.px(8), 0), after=self.sheet_row)
        else:
            self.sheet_error.pack_forget()
        self._ui_lock += 1
        try:
            self.sheet_combo.configure(values=entry.sheets or [])
            self.sheet_combo.set(s.sheet)
            self.header_var.set(str(s.header_row))
            self.col_search.set("")
            self._refresh_column_lists()
            self._rebuild_rules()
            self.sort_combo.configure(values=[NO_SORT] + self._columns)
            self.sort_combo.set(s.sort_by or NO_SORT)
            self.order_combo.set(SORT_ORDERS[1 if s.descending else 0])
            self.order_combo.state(["!disabled" if s.sort_by else "disabled"])
            self.title_var.set(s.title)
            self.count_var.set(s.show_count)
            self.count_label_var.set(s.count_label)
            self.count_entry.state(["!disabled" if s.show_count else "disabled"])
            self.rownum_var.set(s.row_numbers)
            self.skip_empty_var.set(s.skip_empty_rows)
            self.breaks_var.set(s.keep_line_breaks)
        finally:
            self._ui_lock -= 1

    def _sheet_settings(self):
        entry = self._focused()
        return entry.settings if entry and entry.kind == "sheet" else None

    @staticmethod
    def _choices(settings):
        return deepcopy({"columns": settings.columns, "rules": settings.rules,
                         "sort_by": settings.sort_by})

    def _sheet_changed(self, delay=200):
        """The user changed a setting of the focused spreadsheet. It becomes
        that file's intent and the template for new files, but only while
        the sheet is actually loaded (never from an error state)."""
        entry = self._focused()
        if entry and entry.settings and self._sheet_frame(entry) is not None:
            entry.intent = self._choices(entry.settings)
            self.template = deepcopy(entry.settings)
        self._schedule_preview(delay)

    def _on_sheet_selected(self):
        entry = self._focused()
        if self._ui_lock or not entry or not entry.settings:
            return
        if self.sheet_combo.get() == entry.settings.sheet:
            return
        self._stash_choices(entry)
        entry.settings.sheet = self.sheet_combo.get()
        self._after_frame_change(entry)

    def _stash_choices(self, entry):
        """Remember columns / filters / sort of the current sheet and header
        row (if it loaded), so trying another one and coming back loses
        nothing."""
        if self._sheet_frame(entry) is not None:
            s = entry.settings
            entry.stash[(s.sheet, s.header_row)] = self._choices(s)

    def _on_header_row(self, entry=None):
        if self._header_job:
            self.after_cancel(self._header_job)
            self._header_job = None
        entry = entry or self._focused()
        if self._ui_lock or not entry or not entry.settings:
            return
        try:
            value = max(1, min(int(self.header_var.get()), 1000))
        except ValueError:
            return
        if value == entry.settings.header_row:
            return
        self._stash_choices(entry)
        entry.settings.header_row = value
        self._after_frame_change(entry)

    def _on_header_row_typed(self):
        if self._ui_lock:
            return
        if self._header_job:
            self.after_cancel(self._header_job)
        entry = self._focused()
        self._header_job = self.after(500, lambda: self._on_header_row(entry))
        self._header_entry = entry

    def _flush_header_row(self):
        """Apply a header row that is still waiting for its typing pause."""
        if self._header_job:
            self._on_header_row(self._header_entry)

    def _after_frame_change(self, entry):
        """The sheet or header row changed: restore the choices made earlier
        for this sheet / header row, or else re-apply the user's own last
        choices (not whatever survived the previous, perhaps wrong, header
        row), keeping what exists in the new columns."""
        def then(df):
            s = entry.settings
            saved = entry.stash.get((s.sheet, s.header_row)) or entry.intent
            if saved:
                s.columns, s.rules, s.sort_by = deepcopy(
                    (saved["columns"], saved["rules"], saved["sort_by"]))
            adapt_settings(s, list(df.columns))
            if s.columns:  # a sheet that fits becomes the template
                self.template = deepcopy(s)
            self._refresh_sheet_view(entry)

        self._with_frame(entry, then)

    # columns
    def _refresh_column_lists(self):
        s = self._sheet_settings()
        if s is None:
            return
        term = core._fold(self.col_search.get().strip())
        self._avail_items = [c for c in self._columns
                             if c not in s.columns and term in core._fold(c)]
        self.avail_list.delete(0, "end")
        for c in self._avail_items:
            self.avail_list.insert("end", c)
        self.out_list.delete(0, "end")
        for c in s.columns:
            self.out_list.insert("end", c)
        n = len(s.columns)
        self.out_count.configure(text=f"In the table, in order ({n})" if n else
                                 "In the table, in order")

    def _add_columns(self):
        s = self._sheet_settings()
        if s is None:
            return
        picked = [self._avail_items[i] for i in self.avail_list.curselection()]
        if not picked:
            return
        s.columns.extend(c for c in picked if c not in s.columns)
        self._refresh_column_lists()
        self.out_list.selection_set(len(s.columns) - len(picked), "end")
        self.out_list.see("end")
        self._sheet_changed()

    def _add_all_columns(self):
        s = self._sheet_settings()
        if s is None or not self._avail_items:
            return
        s.columns.extend(c for c in self._avail_items if c not in s.columns)
        self._refresh_column_lists()
        self._sheet_changed()

    def _remove_columns(self):
        s = self._sheet_settings()
        if s is None:
            return
        picked = set(self.out_list.curselection())
        if not picked:
            return
        s.columns[:] = [c for i, c in enumerate(s.columns) if i not in picked]
        self._refresh_column_lists()
        self._sheet_changed()

    def _clear_columns(self):
        s = self._sheet_settings()
        if s is None or not s.columns:
            return
        s.columns.clear()
        self._refresh_column_lists()
        self._sheet_changed()

    def _move_columns(self, delta):
        s = self._sheet_settings()
        if s is None:
            return
        sel = sorted(self.out_list.curselection())
        if not sel or (delta < 0 and sel[0] == 0) or (delta > 0 and sel[-1] == len(s.columns) - 1):
            return
        cols = s.columns
        # Move the item nearest the target edge first so shifts don't collide.
        for i in (reversed(sel) if delta > 0 else sel):
            cols[i], cols[i + delta] = cols[i + delta], cols[i]
        self._refresh_column_lists()
        for i in sel:
            self.out_list.selection_set(i + delta)
        self.out_list.see(sel[0] + delta)
        self._sheet_changed()

    # filter rules
    def _rebuild_rules(self):
        s = self._sheet_settings()
        for w in self.rules_frame.winfo_children():
            w.destroy()
        # A Tk frame keeps its old size when its last child goes away.
        self.rules_frame.configure(height=1)
        self._rule_rows = []
        if s is None:
            return
        px = self.px
        for i, rule in enumerate(s.rules):
            row = tk.Frame(self.rules_frame, bg=C["card"])
            row.pack(fill="x", pady=(0, px(6)))
            col = ttk.Combobox(row, state="readonly", width=13, values=self._columns)
            col.set(rule.get("column", ""))
            col.pack(side="left")
            op = ttk.Combobox(row, state="readonly", width=14, values=list(OP_LABELS.values()))
            op.set(OP_LABELS.get(rule.get("op", "word"), OP_LABELS["word"]))
            op.pack(side="left", padx=(px(6), 0))
            remove = ttk.Button(row, text="✕", style="Ghost.TButton", width=2, cursor="hand2",
                                command=lambda i=i: self._remove_rule(i))
            remove.pack(side="right", padx=(px(4), 0))
            Tooltip(remove, "Remove this rule")
            value = HintEntry(row)
            value.set(rule.get("value", ""))
            value.pack(side="left", fill="x", expand=True, padx=(px(6), 0))
            col.bind("<<ComboboxSelected>>", lambda e, i=i, w=col: self._on_rule(i, "column", w.get()))
            op.bind("<<ComboboxSelected>>",
                    lambda e, i=i, w=op: self._on_rule(i, "op", LABEL_TO_OP.get(w.get(), "word")))
            value.on_change(lambda i=i, w=value: self._on_rule(i, "value", w.get(), delay=350))
            self._rule_rows.append((col, op, value))
            self._style_rule(i)

        if s.rules:
            self.rules_empty.pack_forget()
        else:
            self.rules_empty.pack(anchor="w", before=self.table_head)
        if len(s.rules) > 1:
            self.match_row.pack(fill="x", pady=(0, px(8)), before=self.rules_frame)
            self.match_combo.set(s.match if s.match in ("all", "any") else "all")
        else:
            self.match_row.pack_forget()
        self.add_rule_btn.state(["!disabled" if self._columns else "disabled"])

    def _style_rule(self, i):
        s = self._sheet_settings()
        _, _, value = self._rule_rows[i]
        op = s.rules[i].get("op", "word")
        if op in core.VALUELESS_OPS:
            value.state(["disabled"])
            value.set_hint("")
        else:
            value.state(["!disabled"])
            value.set_hint({"word": "word, or * for all", "regex": "regular expression"}
                           .get(op, "value"))

    def _on_rule(self, i, key, value, delay=150):
        s = self._sheet_settings()
        if self._ui_lock or s is None or i >= len(s.rules):
            return
        s.rules[i][key] = value
        if key == "op":
            self._style_rule(i)
        self._sheet_changed(delay)

    def _add_rule(self):
        s = self._sheet_settings()
        if s is None or not self._columns:
            return
        column = core.find_column(self._columns, "Property") or self._columns[0]
        s.rules.append({"column": column, "op": "word", "value": ""})
        self._ui_lock += 1
        try:
            self._rebuild_rules()
        finally:
            self._ui_lock -= 1
        self._fit_left_pane()
        self._rule_rows[-1][2].focus_set()
        self.left.see(self.rules_frame)
        self._sheet_changed()

    def _remove_rule(self, i):
        s = self._sheet_settings()
        if s is None or i >= len(s.rules):
            return
        del s.rules[i]
        self._ui_lock += 1
        try:
            self._rebuild_rules()
        finally:
            self._ui_lock -= 1
        self._fit_left_pane()
        self._sheet_changed()

    def _on_match(self):
        s = self._sheet_settings()
        if s is not None and not self._ui_lock:
            s.match = self.match_combo.get()
            self._sheet_changed()

    # table options
    def _on_sort(self):
        s = self._sheet_settings()
        if s is None or self._ui_lock:
            return
        index = self.sort_combo.current()  # 0 is "no sorting"
        s.sort_by = self._columns[index - 1] if 0 < index <= len(self._columns) else ""
        s.descending = self.order_combo.get() == SORT_ORDERS[1]
        self.order_combo.state(["!disabled" if s.sort_by else "disabled"])
        self._sheet_changed()

    def _on_table_option(self):
        s = self._sheet_settings()
        if s is None or self._ui_lock:
            return
        s.title = self.title_var.get()
        s.show_count = self.count_var.get()
        s.count_label = self.count_label_var.get()
        s.row_numbers = self.rownum_var.get()
        s.skip_empty_rows = self.skip_empty_var.get()
        s.keep_line_breaks = self.breaks_var.get()
        self.count_entry.state(["!disabled" if s.show_count else "disabled"])
        self._sheet_changed(300)

    # ---------------------------------------------------------------- documents
    def _doc_options(self):
        return {key: var.get() for key, var in self.doc_vars.items()}

    def _on_doc_option(self):
        entry = self._focused()
        if entry and entry.kind == "doc":
            self._schedule_preview(0)

    @staticmethod
    def _convert_options(opts, path, merging):
        """Core options for one document from the UI options."""
        result = {k: v for k, v in opts.items() if k != "title_from_name"}
        if opts.get("title_from_name") and not merging:
            result["title"] = Path(path).stem
        return result

    def _doc_cache_key(self, path, opts):
        try:
            st = os.stat(path)
            stamp = (st.st_mtime_ns, st.st_size)
        except OSError:
            stamp = (0, 0)
        return (str(path), stamp, tuple(sorted(opts.items())))

    def _cache_doc(self, key, md):
        self._doc_cache[key] = md
        self._doc_cache.move_to_end(key)
        while len(self._doc_cache) > 12:
            self._doc_cache.popitem(last=False)

    # ---------------------------------------------------------------- preview
    def _schedule_preview(self, delay=200):
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(delay, self._refresh_preview)

    def _refresh_preview(self):
        self._preview_job = None
        self._preview_token += 1
        token = self._preview_token
        entry = self._focused()
        self._preview_md = ""
        self.copy_btn.state(["disabled"])
        if entry is None:
            self._set_preview_info("", "")
            if self.files:
                self._show_message("Select a file to preview its Markdown.")
            else:
                hint = "drag files here, " if self.dnd else ""
                self._show_message(f"Add files to get started: {hint}use “Add files…”, or "
                                   "“Add folder…”.\n\nThe Markdown preview of the selected file "
                                   "appears here.")
            return
        self._set_preview_info(entry.path.name, "")
        if entry.kind == "sheet":
            self._preview_sheet(entry)
        else:
            self._preview_doc(entry, token)

    def _preview_sheet(self, entry):
        if entry.error:
            self._show_message(f"Could not open this file.\n{entry.error}", error=True)
            return
        df = self._sheet_frame(entry)
        s = entry.settings
        error = entry.frame_errors.get((s.sheet, s.header_row)) if s else None
        if error:
            self._show_message(f"Could not read sheet “{s.sheet}” with header row "
                               f"{s.header_row}.\n{error}\n\nPick another header row or sheet "
                               "on the left.", error=True)
            return
        if s is None or df is None:
            self._show_message("Loading spreadsheet…")
            return
        if not s.columns:
            self._set_preview_info(entry.path.name, f"{len(df)} rows · {len(df.columns)} columns")
            self._show_message("Pick at least one column under “Columns” to build the table.")
            return
        try:
            md, total = core.table_to_markdown(df, s.columns, limit=PREVIEW_ROWS,
                                               **s.table_options())
        except Exception as exc:
            self._show_message(core.describe_error(exc), error=True)
            return
        stats = f"{total} of {len(df)} rows · {len(s.columns)} column{'s' if len(s.columns) != 1 else ''}"
        note = (f"Preview shows the first {PREVIEW_ROWS} rows; the file will contain all {total}."
                if total > PREVIEW_ROWS else "")
        self._show_markdown(md, stats, note)
        self.copy_btn.state(["!disabled"])

    def _preview_doc(self, entry, token):
        opts = self._convert_options(self._doc_options(), entry.path, merging=False)
        key = self._doc_cache_key(entry.path, opts)
        if key in self._doc_cache:
            self._show_doc(self._doc_cache[key])
            return
        self._show_message("Converting preview…")
        if self._doc_preview_busy:
            self._doc_preview_again = True
            return
        self._doc_preview_busy = True
        path = entry.path

        def finish():
            self._doc_preview_busy = False
            if self._doc_preview_again:
                self._doc_preview_again = False
                self._schedule_preview(0)

        def done(md):
            self._cache_doc(key, md)
            if token == self._preview_token:
                self._show_doc(md)
            finish()

        def fail(exc):
            if token == self._preview_token:
                self._show_message(f"Could not convert this file.\n{core.describe_error(exc)}",
                                   error=True)
            finish()

        self._bg(lambda: core.document_to_markdown(path, **opts), done, fail)

    def _show_doc(self, md):
        words = len(md.split())
        headings = sum(1 for line in md.split("\n") if re.match(r"#{1,6} ", line))
        tables = sum(1 for line in md.split("\n") if re.match(r"\|(---\|)+$", line))
        parts = [f"{words:,} words", f"{headings} heading{'s' if headings != 1 else ''}"]
        if tables:
            parts.append(f"{tables} table{'s' if tables != 1 else ''}")
        note = ""
        if not md.strip():
            self._show_message("No text found. A scanned PDF has no text layer; "
                               "it needs OCR before it can be converted.", error=True)
            return
        if len(md) > PREVIEW_CHARS:
            note = "Preview is cut short; the file will contain the whole document."
        self._show_markdown(md[:PREVIEW_CHARS], " · ".join(parts), note)
        self._preview_md = md
        self.copy_btn.state(["!disabled"])

    def _set_preview_info(self, name, stats):
        self.preview_name.configure(text=name)
        self.preview_stats.configure(text=stats)

    def _show_message(self, text, error=False):
        self.preview_note.pack_forget()
        p = self.preview
        p.configure(state="normal")
        p.delete("1.0", "end")
        p.insert("1.0", text, "err" if error else "msg")
        p.configure(state="disabled")

    def _show_markdown(self, md, stats="", note=""):
        self.preview_stats.configure(text=stats)
        if note:
            self.preview_note.configure(text=note)
            self.preview_note.pack(side="top", fill="x", before=self.preview_page.hbar)
        else:
            self.preview_note.pack_forget()
        p = self.preview
        p.configure(state="normal")
        p.delete("1.0", "end")
        p.insert("1.0", md)
        for n, line in enumerate(md.split("\n")[:HIGHLIGHT_LINES], 1):
            if not line:
                continue
            start, end = f"{n}.0", f"{n}.end"
            if line.startswith("#"):
                p.tag_add("h", start, end)
            elif line.startswith("|"):
                p.tag_add("tsep" if re.match(r"\|(---\|)+$", line) else "table", start, end)
            elif line.startswith(">"):
                p.tag_add("quote", start, end)
            elif line == "---":
                p.tag_add("rule", start, end)
            elif re.match(r"\*\*[^*]+:\*\* ", line):
                p.tag_add("meta", start, end)
            else:
                m = re.match(r"(\s*)(- |\d+\. )", line)
                if m:
                    p.tag_add("marker", f"{n}.{len(m.group(1))}", f"{n}.{m.end()}")
        p.configure(state="disabled")
        p.yview_moveto(0)
        p.xview_moveto(0)

    def _apply_wrap(self):
        self.preview.configure(wrap="word" if self.wrap_var.get() else "none")

    def _copy_preview(self):
        entry = self._focused()
        if entry is None:
            return
        md = self._preview_md
        if entry.kind == "sheet":
            df = self._sheet_frame(entry)
            if df is None or not entry.settings or not entry.settings.columns:
                return
            try:
                md, _ = core.table_to_markdown(df, entry.settings.columns,
                                               **entry.settings.table_options())
            except Exception as exc:
                self._warn(core.describe_error(exc))
                return
        if md:
            self.clipboard_clear()
            self.clipboard_append(md)
            self._set_status(f"Copied the Markdown of {entry.path.name} to the clipboard.", "success")

    def _select_tab(self, name):
        self._current_tab = name
        for tab, (label, line) in self._tabs.items():
            active = tab == name
            label.configure(fg=C["accent"] if active else C["muted"])
            line.configure(bg=C["accent"] if active else C["card"])
        px = self.px
        if name == "Preview":
            self.log_page.pack_forget()
            self.clear_log_btn.pack_forget()
            self.copy_btn.pack(side="right", pady=px(6))
            self.wrap_cb.pack(side="right", padx=px(10))
            self.info_bar.pack(fill="x", padx=px(14), pady=(px(6), px(6)), before=self.pages)
            self.preview_page.pack(fill="both", expand=True)
        else:
            self.preview_page.pack_forget()
            self.info_bar.pack_forget()
            self.copy_btn.pack_forget()
            self.wrap_cb.pack_forget()
            self.log_page.pack(fill="both", expand=True)
            self.clear_log_btn.pack(side="right", pady=px(6))
            self._log_unseen = 0
        self._tabs["Log"][0].configure(
            text="Log" if not self._log_unseen else f"Log ({self._log_unseen})")

    def _log(self, text, kind="info"):
        log = self.log
        log.configure(state="normal")
        log.insert("end", time.strftime("%H:%M:%S  "), "time")
        log.insert("end", text + "\n", kind)
        log.configure(state="disabled")
        log.see("end")
        if self._current_tab != "Log":
            self._log_unseen += 1
            label = self._tabs["Log"][0]
            label.configure(text=f"Log ({self._log_unseen})")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ---------------------------------------------------------------- output
    def _on_output_option(self):
        self._update_controls()

    def _on_output_folder_typed(self):
        if self.out_folder.get().strip() and self.out_mode.get() != "folder":
            self.out_mode.set("folder")
            self._update_controls()

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(parent=self, title="Save Markdown files in…",
                                         initialdir=self.out_folder.get() or self.last_dir)
        if folder:
            self.out_folder.set(folder)
            self.out_mode.set("folder")
            self._update_controls()

    def _browse_output_file(self):
        entry = self._focused()
        initial = f"{entry.path.stem}.md" if entry else "output.md"
        initial_dir = self.out_folder.get() if self.out_mode.get() == "folder" else (
            str(entry.path.parent) if entry else self.last_dir)
        path = filedialog.asksaveasfilename(parent=self, title="Save Markdown as",
                                            defaultextension=".md", initialfile=initial,
                                            initialdir=initial_dir or self.last_dir,
                                            filetypes=[("Markdown", "*.md"), ("All files", "*.*")])
        if path:
            self.out_name.set(path)

    def _plan_outputs(self, entries):
        """{"merge": Path | None, "targets": {iid: Path}} for a conversion."""
        folder = None
        if self.out_mode.get() == "folder":
            raw = self.out_folder.get().strip()
            if not raw:
                raise ValueError("Choose an output folder, or pick “Next to each source file”.")
            folder = Path(raw).expanduser()

        def place(name, entry, typed):
            # '~' only means "home" in a name the user typed, never in one
            # made from a source file called '~notes.docx'.
            p = Path(name).expanduser() if typed else Path(name)
            if not p.suffix:
                p = p.with_name(p.name + ".md")
            p = p if p.is_absolute() else (folder or entry.path.parent) / p
            return Path(os.path.abspath(p))

        def key(path):
            return os.path.normcase(os.path.abspath(path))

        if self.merge_var.get() and len(entries) > 1:
            name = self.merged_name.get().strip()
            return {"merge": place(name or "merged.md", entries[0], bool(name)), "targets": {}}
        if len(entries) == 1:
            entry = entries[0]
            name = (self.out_name.get().strip() if entry.iid == self.focus_iid
                    else entry.out_name)
            target = place(name or f"{entry.path.stem}.md", entry, bool(name))
            return {"merge": None, "targets": {entry.iid: target}}
        targets, used = {}, set()
        for entry in entries:
            target = place(f"{entry.path.stem}.md", entry, False)
            if key(target) in used:  # a.xlsx + a.docx -> a_docx.md
                target = target.with_name(f"{entry.path.stem}_{entry.path.suffix[1:].lower()}.md")
            n = 2
            while key(target) in used:
                target = target.with_name(f"{entry.path.stem}_{n}.md")
                n += 1
            used.add(key(target))
            targets[entry.iid] = target
        return {"merge": None, "targets": targets}

    # ---------------------------------------------------------------- convert
    def _convert(self, which):
        if self.busy:
            return
        self._flush_header_row()  # a header row typed just before converting
        if which == "selected":
            iids = [i for i in self.tree.get_children() if i in self.tree.selection()]
            if not iids and self.focus_iid:
                iids = [self.focus_iid]
        else:
            iids = list(self.tree.get_children())
        if not iids:
            self._warn("Add a file first." if not self.files else "Select a file to convert.")
            return
        entries = [self.files[i] for i in iids]
        for entry in entries:
            if entry.kind == "sheet" and entry.settings is not None and not entry.settings.columns:
                self._warn(f"{entry.path.name}: pick at least one column to extract.")
                return
        try:
            plan = self._plan_outputs(entries)
        except ValueError as exc:
            self._warn(str(exc))
            return
        targets = [plan["merge"]] if plan["merge"] else list(plan["targets"].values())
        existing = [t for t in targets if t.exists()]
        if existing:
            names = "\n".join(f"  {t.name}" for t in existing[:6])
            more = f"\n  … and {len(existing) - 6} more" if len(existing) > 6 else ""
            if not messagebox.askyesno(
                    "Replace existing files?",
                    f"{len(existing)} output file{'s' if len(existing) != 1 else ''} already "
                    f"exist{'s' if len(existing) == 1 else ''}:\n\n{names}{more}\n\nReplace them?",
                    parent=self, icon="warning"):
                return

        merging = plan["merge"] is not None
        doc_opts = self._doc_options()
        template = deepcopy(self.template)
        jobs = []
        for entry in entries:
            opts = self._convert_options(doc_opts, entry.path, merging)
            cached = (self._doc_cache.get(self._doc_cache_key(entry.path, opts))
                      if entry.kind == "doc" else None)
            jobs.append((entry.iid, entry.path, entry.kind, deepcopy(entry.settings), opts, cached,
                         plan["targets"].get(entry.iid)))

        self._set_busy(True, len(jobs))
        verb = "Merging" if merging else "Converting"
        self._log(f"{verb} {len(jobs)} file{'s' if len(jobs) != 1 else ''}…", "info")

        def work():
            results, parts = [], []
            for n, (iid, path, kind, settings, opts, cached, target) in enumerate(jobs, 1):
                self._post(self._on_job_start, iid, n, len(jobs))
                try:
                    md, count = self._render(path, kind, settings, opts, cached, template)
                    if merging:
                        parts.append((path.name, md))
                    else:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        target.write_text(md, encoding="utf-8")
                    results.append((iid, target, None, count, md))
                except Exception as exc:
                    results.append((iid, None, core.describe_error(exc), None, ""))
                self._post(self._on_job_end, results[-1], merging)
            merged = None
            if merging and parts:
                merged = plan["merge"]
                merged.parent.mkdir(parents=True, exist_ok=True)
                merged.write_text(core.merge_markdown(parts), encoding="utf-8")
            return results, merged

        def fail(exc):
            # Only the merged file write can get here; nothing was saved.
            self._set_busy(False)
            for iid, *_ in jobs:
                entry = self.files.get(iid)
                if entry and entry.status in ("Converting…", "Merging…"):
                    self._set_file_status(entry, "Error", "error")
            self._log(f"Conversion failed: {core.describe_error(exc)}", "error")
            self._set_status("Conversion failed. See the log.", "error")
            self._warn(f"Could not write the output.\n{core.describe_error(exc)}")

        self._bg(work, self._on_convert_done, fail)

    @staticmethod
    def _render(path, kind, settings, opts, cached, template):
        """Markdown for one file (runs on a worker thread)."""
        if kind == "doc":
            md = cached if cached is not None else core.document_to_markdown(path, **opts)
            body = md.strip()
            if opts.get("title") and body == f"# {opts['title']}":
                body = ""
            if not body:
                raise ValueError("no text found. A scanned PDF has no text layer; "
                                 "it needs OCR before it can be converted")
            return md, None
        if settings is None:  # never opened in the UI: start from the template
            sheets = core.sheet_names(path)
            if not sheets:
                raise ValueError("the workbook has no sheets")
            settings, _ = initial_settings(
                template, sheets, lambda sheet, row: core.read_excel(path, sheet, row))
            if not settings.columns:
                raise ValueError("no columns to extract: open the file and pick columns first")
        # Read again so the output reflects the file as it is now.
        df = core.read_excel(path, settings.sheet, settings.header_row)
        return core.table_to_markdown(df, settings.columns, **settings.table_options())

    def _on_job_start(self, iid, n, total):
        entry = self.files.get(iid)
        if entry:
            self._set_file_status(entry, "Converting…", "busy")
        self._set_status(f"Converting {n} of {total}: {entry.path.name if entry else ''}", "busy")

    def _on_job_end(self, result, merging=False):
        iid, target, error, count, _ = result
        self.progress.step(1)
        entry = self.files.get(iid)
        if entry is None:
            return
        if error:
            self._set_file_status(entry, "Error", "error")
            self._log(f"✗ {entry.path.name}: {error}", "error")
        elif merging:  # done only once the merged file is written
            self._set_file_status(entry, "Merging…", "busy")
        else:
            self._set_file_status(entry, "Done ✓", "done")

    def _on_convert_done(self, payload):
        results, merged = payload
        self._set_busy(False)
        ok = [r for r in results if not r[2]]
        failed = [r for r in results if r[2]]
        if merged:
            for iid, *_ in ok:
                if iid in self.files:
                    self._set_file_status(self.files[iid], "Done ✓", "done")
        for iid, target, _, count, _ in ok:
            entry = self.files.get(iid)
            name = entry.path.name if entry else iid
            rows = f" ({count} row{'s' if count != 1 else ''})" if count is not None else ""
            self._log(f"✓ {name} → {target.name if target else 'merged file'}{rows}", "ok")
        if merged:
            self._log(f"Saved {merged}", "info")
            self.last_output = merged
        elif ok:
            for folder in dict.fromkeys(r[1].parent for r in ok):
                self._log(f"Saved in {folder}", "info")
            self.last_output = ok[-1][1]
        self._update_controls()

        if ok and not failed:
            if merged:
                text = f"Merged {len(ok)} files into {merged.name}"
            elif len(ok) == 1:
                text = f"Saved {ok[0][1].name} in {ok[0][1].parent}"
            else:
                folders = {r[1].parent for r in ok}
                where = f" in {folders.pop()}" if len(folders) == 1 else " next to their sources"
                text = f"Converted {len(ok)} files{where}"
            self._set_status(text, "success")
        elif ok:
            self._set_status(f"Converted {len(ok)}, failed {len(failed)}. See the log.", "warn")
        else:
            self._set_status("Conversion failed. See the log.", "error")
        if failed:
            first = self.files.get(failed[0][0])
            detail = "\n".join(f"• {self.files[r[0]].path.name}: {r[2]}" for r in failed[:5]
                               if r[0] in self.files)
            self._warn(f"{len(failed)} of {len(results)} file{'s' if len(results) != 1 else ''} "
                       f"could not be converted:\n\n{detail}")
            if first:
                self._select_tab("Log")
        if ok and self.open_after_var.get() and self.last_output:
            self._open_location(self.last_output)

    def _set_busy(self, busy, total=0):
        self.busy = busy
        if busy:
            self.progress.configure(maximum=max(total, 1), value=0)
            self.progress.pack(side="left", padx=(0, self.px(12)), before=self.status_label)
        else:
            self.progress.pack_forget()
        self._update_controls()

    # ---------------------------------------------------------------- misc
    def _refresh_all(self):
        n = len(self.files)
        self.files_count.configure(text=f"{n} file{'s' if n != 1 else ''}" if n else "")
        if n:
            self.tree_empty.place_forget()
        else:
            drop = "Drop files or folders here\nor click " if self.dnd else "Click "
            self.tree_empty.configure(text=f"{drop}“Add files…” to get started")
            self.tree_empty.place(relx=0.5, rely=0.58, anchor="center")
        if not n and not self.status_label.cget("text"):
            self._set_status("Ready. Add files to begin.", "info")
        self._on_select()
        self._show_options()
        self._schedule_preview(0)
        self._update_controls()

    def _update_controls(self):
        selected = len(self.tree.selection())
        total = len(self.files)
        idle = not self.busy
        self.convert_sel_btn.configure(
            text=f"Convert selected ({selected})" if selected > 1 else "Convert selected")
        self.convert_sel_btn.state(["!disabled" if idle and selected else "disabled"])
        self.convert_all_btn.configure(text=f"Convert all ({total})" if total > 1 else "Convert all")
        self.convert_all_btn.state(["!disabled" if idle and total else "disabled"])
        self.remove_btn.state(["!disabled" if idle and selected else "disabled"])
        self.clear_btn.state(["!disabled" if idle and total else "disabled"])
        self.open_btn.state(["!disabled" if self.last_output else "disabled"])
        self.merged_name.state(["!disabled" if self.merge_var.get() else "disabled"])
        entry = self._focused()
        if selected > 1:
            hint = "Converting several files: each gets <file name>.md, so this is not used."
            if self.merge_var.get():
                hint = "Converting several files into one merged file, so this is not used."
            self.out_name.set_hint("")
        else:
            hint = "Leave empty to use the source file name. A full path is also accepted."
            self.out_name.set_hint(f"{entry.path.stem}.md" if entry else "<file name>.md")
        self.out_name_hint.configure(text=hint)

    def _set_status(self, text, kind="info"):
        colors = {"info": C["faint"], "success": C["success"], "error": C["danger"],
                  "warn": C["warn"], "busy": C["accent"]}
        self.status_dot.configure(fg=colors.get(kind, C["faint"]))
        self.status_label.configure(text=text, fg=C["danger"] if kind == "error" else C["fg"])

    def _open_last_output(self):
        if self.last_output:
            self._open_location(self.last_output)

    def _open_location(self, path):
        path = Path(path)
        try:
            if sys.platform == "win32":
                if path.is_file():
                    # One quoted argument, so paths with commas or spaces work.
                    subprocess.Popen(f'explorer /select,"{path}"')
                else:
                    os.startfile(str(path if path.is_dir() else path.parent))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(path)] if path.exists() else ["open", str(path.parent)])
            else:
                subprocess.Popen(["xdg-open", str(path.parent if path.is_file() else path)])
        except Exception as exc:
            self._warn(f"Could not open the folder.\n{exc}")

    def _show_shortcuts(self):
        dnd = ("Drag & drop: drop files or folders anywhere on the window."
               if self.dnd else
               "Drag & drop: install the optional package with  pip install tkinterdnd2")
        messagebox.showinfo(
            "Shortcuts",
            "Ctrl+O\tAdd files\n"
            "Ctrl+Shift+O\tAdd a folder\n"
            "Delete\tRemove the selected files\n"
            "Ctrl+Enter\tConvert the selected files\n"
            "Ctrl+Shift+Enter\tConvert all files\n"
            "F5\tReload the highlighted file from disk\n"
            "Right-click\tFile menu (convert, reload, show in folder)\n\n"
            "Filter values: 'contains word' matches whole words at their start, "
            "so 'gereksinim' also matches 'gereksinimler'. Matching ignores case "
            "(Turkish İ/i included). Use * to keep every row.\n\n" + dnd,
            parent=self,
        )

    def _warn(self, msg):
        messagebox.showerror("Something went wrong", msg, parent=self)

    # ---------------------------------------------------------------- window
    def _restore_window(self):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(self.px(1240), int(sw * 0.92))
        height = min(self.px(820), int(sh * 0.88))
        geometry = self.saved.get("geometry") or ""
        m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", geometry)
        x, y = (sw - width) // 2, max(0, (sh - height) // 3)
        if m:
            w, h, gx, gy = map(int, m.groups())
            width, height = min(w, sw), min(h, sh)
            if 0 <= gx < sw - 100 and 0 <= gy < sh - 100:
                x, y = gx, gy
        # The size and place used when the window is restored down.
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(self.px(860), self.px(600))
        self._maximize()
        sash = self.saved.get("sash")

        def place():
            if isinstance(sash, int) and sash > self.px(LEFT_WIDTH):
                self.paned.sash_place(0, sash, 1)  # the user's own wider layout
            self._fit_left_pane()
            self._keep_preview_visible()
        self.after(60, place)

    def _maximize(self):
        """Start with the window maximized."""
        try:
            if sys.platform == "win32":
                self.state("zoomed")
            elif sys.platform == "darwin":
                self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
            else:
                self.attributes("-zoomed", True)
        except tk.TclError:
            pass

    def _on_close(self):
        if self.busy and not messagebox.askyesno(
                "Quit?", "A conversion is still running. Quit anyway?", parent=self):
            return
        zoomed = self.state() == "zoomed"
        data = dict(self.saved)
        data.update({
            "geometry": self.saved.get("geometry") if zoomed else self.geometry(),
            "sash": self.paned.sash_coord(0)[0],
            "doc": self._doc_options(),
            "out_mode": self.out_mode.get(),
            "out_folder": self.out_folder.get(),
            "merge": self.merge_var.get(),
            "merged_name": self.merged_name.get(),
            "open_after": self.open_after_var.get(),
            "wrap": self.wrap_var.get(),
            "last_dir": self.last_dir,
            "sheet_template": asdict(self.template) if self.template else None,
            "theme": self.theme_choice,
        })
        save_settings(data)
        for job in (getattr(self, "_pump_job", None), self._preview_job, self._header_job,
                    getattr(self, "_theme_job", None)):
            if job:
                try:
                    self.after_cancel(job)
                except tk.TclError:
                    pass
        self.destroy()


def main():
    _enable_dpi_awareness()
    app = App()
    if len(sys.argv) > 1:  # files passed on the command line / dropped on the .exe
        app._add_paths(sys.argv[1:])
    app.mainloop()


if __name__ == "__main__":
    main()
