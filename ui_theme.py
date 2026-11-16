"""Light and dark colours for a Tkinter UI, switched while the program runs.

The UI keeps its colours in one dict (`C`), which every widget reads when it
is made. To switch, `apply_palette` replaces the dict's values and `recolor`
walks the widgets and swaps each old colour for the colour of the same name
in the new palette (tk options, Text tags, Canvas items, Treeview tags).
ttk styles and pictures drawn from the colours are made again by the app.

Shared by the DOCX to LaTeX Converter, MD Converter and Code Reviewer: copy this
file as it is.
"""

import subprocess
import sys

LIGHT = {
    "bg": "#f3f5f9",
    "card": "#ffffff",
    "border": "#e1e5ec",
    "border_strong": "#c9d1de",
    "fg": "#1f2430",
    "muted": "#6b7280",
    "faint": "#9aa3b2",
    "accent": "#2f6fed",
    "accent_hover": "#2458c7",
    "accent_pressed": "#1f4fb5",
    "accent_soft": "#e8efff",
    "accent_disabled": "#a9c0f5",
    "field": "#ffffff",
    "hover": "#f1f4f9",
    "pressed": "#e5e9f0",
    "code_bg": "#fbfcfe",
    "gutter_fg": "#aab2c0",
    "current_line": "#f2f5fb",
    "found": "#fff1a8",
    "found_current": "#ffc94d",
    "error_line": "#fde2e2",
    "note_warn": "#fff4e0",
    "pdf_bg": "#e9ecf2",
    "pdf_shadow": "#c9cfd9",
    "cmd": "#1d4ed8",              # LaTeX commands
    "env": "#9333ea",
    "comment": "#94a3b8",
    "heading": "#1d4ed8",          # Markdown headings
    "success": "#15803d",
    "danger": "#dc2626",
    "warn": "#b45309",
    "tooltip": "#262b36",
    "on_accent": "white",          # text on accent-coloured buttons
    "select": "#cfdcfb",           # selected text
    "select_inactive": "#dde6fb",
    "row_selected": "#eef1f6",     # selected list row without focus
    "scroll": "#d5dbe5",
    "scroll_active": "#b8c1cf",
    # Code Reviewer: notices, charts, source viewer, findings
    "warn_bg": "#fff7e6",
    "warn_border": "#f3d19c",
    "danger_bg": "#fdecec",
    "danger_border": "#f5b5b5",
    "hint_bg": "#fff4cc",
    "grid": "#eef1f6",             # chart grid lines (= row_selected)
    "bar": "#2f6fed",              # chart bars (= accent)
    "bar_hover": "#1f4fb5",        # (= accent_pressed)
    "sel": "#cfdcfb",              # (= select)
    "nav_sel": "#e8efff",          # (= accent_soft)
    "code_inline": "#eef2f8",
    "syn_keyword": "#7c3aed",
    "syn_comment": "#6b7f6b",
    "syn_preproc": "#2563eb",
    "syn_number": "#0f766e",
    "focus_line": "#dbe6ff",
    "found_strong": "#ffd54d",
    "sev_high": "#b91c1c",
    "sev_warning": "#92400e",
    "sev_info": "#374151",
}

DARK = {
    "bg": "#15171c",
    "card": "#1d2027",
    "border": "#2b303a",
    "border_strong": "#3c4250",
    "fg": "#e3e6ec",
    "muted": "#a0a8b6",
    "faint": "#6f7787",
    "accent": "#3b7cf6",
    "accent_hover": "#5a91f7",
    "accent_pressed": "#2c66d6",
    "accent_soft": "#1f2d48",
    "accent_disabled": "#2a3a5c",
    "field": "#1d2027",
    "hover": "#252932",
    "pressed": "#2e333d",
    "code_bg": "#181a20",
    "gutter_fg": "#525a68",
    "current_line": "#20232b",
    "found": "#4d4415",
    "found_current": "#7a5f0e",
    "error_line": "#45222a",
    "note_warn": "#3a2f19",
    "pdf_bg": "#101216",
    "pdf_shadow": "#08090b",
    "cmd": "#79a6ff",
    "env": "#cf9cff",
    "comment": "#6b7689",
    "heading": "#79a6ff",
    "success": "#45b97c",
    "danger": "#f26d6d",
    "warn": "#e3a446",
    "tooltip": "#343a46",
    "on_accent": "white",
    "select": "#284a7a",
    "select_inactive": "#2a3446",
    "row_selected": "#272b34",
    "scroll": "#3a404c",
    "scroll_active": "#4c5361",
    "warn_bg": "#352a17",
    "warn_border": "#6b5221",
    "danger_bg": "#3d2125",
    "danger_border": "#7a3a40",
    "hint_bg": "#3b3416",
    "grid": "#272b34",
    "bar": "#3b7cf6",
    "bar_hover": "#2c66d6",
    "sel": "#284a7a",
    "nav_sel": "#1f2d48",
    "code_inline": "#2a2f3a",
    "syn_keyword": "#b48cff",
    "syn_comment": "#7f9a7f",
    "syn_preproc": "#6ea8ff",
    "syn_number": "#4fc1b0",
    "focus_line": "#24324d",
    "found_strong": "#8a6d12",
    "sev_high": "#ff8a8a",
    "sev_warning": "#f2b866",
    "sev_info": "#c5cbd6",
}

PALETTES = {"light": LIGHT, "dark": DARK}
CHOICES = (("light", "Light"), ("dark", "Dark"), ("system", "System"))

COLOR_OPTIONS = ("background", "foreground", "highlightbackground", "highlightcolor",
                 "insertbackground", "selectbackground", "selectforeground",
                 "inactiveselectbackground", "activebackground", "activeforeground",
                 "disabledforeground", "troughcolor", "readonlybackground",
                 "selectcolor")


def check_palettes(*palettes):
    """Colours that are equal in one palette must be equal in every other,
    or `recolor` could not tell which one a widget has. Raises ValueError."""
    palettes = palettes or tuple(PALETTES.values())
    for a in palettes:
        seen = {}
        for key, value in a.items():
            first = seen.setdefault(value.lower(), key)
            for b in palettes:
                if b[first].lower() != b[key].lower():
                    raise ValueError(f"'{key}' and '{first}' have the same colour in one "
                                     "palette but not in another")
        for b in palettes:
            if set(a) != set(b):
                raise ValueError("the palettes have different names: %s"
                                 % sorted(set(a) ^ set(b)))


def system_theme():
    """'dark' or 'light', as the operating system is set."""
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
                                ) as key:
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value else "dark"
        except OSError:
            return "light"
    if sys.platform == "darwin":
        try:
            out = subprocess.run(["defaults", "read", "-g", "AppleInterfaceStyle"],
                                 capture_output=True, text=True, timeout=2)
            return "dark" if "dark" in out.stdout.lower() else "light"
        except Exception:
            return "light"
    return "light"


def resolve(choice):
    """Palette name for a choice ('light', 'dark' or 'system')."""
    if choice == "system":
        return system_theme()
    return choice if choice in PALETTES else "light"


def apply_palette(colors, name):
    """Put the colours of palette `name` into the dict `colors` (in place, so
    every module holding it sees them); returns {old colour: new colour}."""
    new = PALETTES[name]
    mapping = {value.lower(): new[key] for key, value in colors.items() if key in new}
    colors.clear()
    colors.update(new)
    return mapping


def recolor(widget, mapping):
    """Swap the colours of `widget` and everything inside it."""
    import tkinter as tk
    from tkinter import ttk

    def swap(value):
        return mapping.get(str(value).lower())

    try:
        keys = widget.keys()
    except tk.TclError:
        return
    for option in COLOR_OPTIONS:
        if option in keys:
            try:
                new = swap(widget.cget(option))
                if new is not None:
                    widget.configure(**{option: new})
            except tk.TclError:
                pass
    if isinstance(widget, tk.Text):
        for tag in widget.tag_names():
            for option in ("foreground", "background"):
                new = swap(widget.tag_cget(tag, option))
                if new is not None:
                    widget.tag_configure(tag, **{option: new})
    elif isinstance(widget, tk.Canvas):
        for item in widget.find_all():
            for option in ("fill", "outline"):
                try:
                    new = swap(widget.itemcget(item, option))
                except tk.TclError:
                    continue
                if new is not None:
                    widget.itemconfigure(item, **{option: new})
    elif isinstance(widget, ttk.Treeview):
        for tag in widget.tk.splitlist(widget.tk.call(widget, "tag", "names")):
            for option in ("foreground", "background"):
                new = swap(widget.tag_configure(tag, option))
                if new is not None:
                    widget.tag_configure(tag, **{option: new})
    elif isinstance(widget, tk.Menu):
        last = widget.index("end")
        for i in range(0 if last is None else last + 1):
            for option in ("background", "foreground", "activebackground",
                           "activeforeground"):
                try:
                    new = swap(widget.entrycget(i, option))
                except tk.TclError:
                    continue
                if new is not None:
                    widget.entryconfigure(i, **{option: new})
    for child in widget.winfo_children():
        recolor(child, mapping)


def style_combobox_lists(root, colors):
    """The drop-down lists of ttk comboboxes are made once; colour them too."""
    from tkinter import ttk

    def walk(widget):
        if isinstance(widget, ttk.Combobox):
            try:
                popdown = widget.tk.call("ttk::combobox::PopdownWindow", widget)
                widget.tk.call(f"{popdown}.f.l", "configure",
                               "-background", colors["field"], "-foreground", colors["fg"],
                               "-selectbackground", colors["accent_soft"],
                               "-selectforeground", colors["fg"])
            except Exception:
                pass
        for child in widget.winfo_children():
            walk(child)
    walk(root)


def dialog_options(root, colors):
    """Colours for plain Tk dialogs made later (e.g. simpledialog)."""
    for pattern, value in (("*Toplevel.background", "card"),
                           ("*Toplevel*Frame.background", "card"),
                           ("*Toplevel*Label.background", "card"),
                           ("*Toplevel*Label.foreground", "fg"),
                           ("*Toplevel*Entry.background", "field"),
                           ("*Toplevel*Entry.foreground", "fg"),
                           ("*Toplevel*Entry.insertBackground", "fg"),
                           ("*Toplevel*Entry.selectBackground", "select"),
                           ("*Toplevel*Entry.selectForeground", "fg"),
                           ("*Toplevel*Button.background", "pressed"),
                           ("*Toplevel*Button.foreground", "fg"),
                           ("*Toplevel*Button.activeBackground", "hover"),
                           ("*Toplevel*Button.activeForeground", "fg")):
        root.option_add(pattern, colors[value])


def set_title_bar(window, dark):
    """Dark or light window title bar (Windows 10 20H1 and later)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        value = ctypes.c_int(1 if dark else 0)
        for attribute in (20, 19):      # DWMWA_USE_IMMERSIVE_DARK_MODE (new, old number)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
        # repaint the frame: NOSIZE | NOMOVE | NOZORDER | NOACTIVATE | FRAMECHANGED
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 |
                                          0x0010 | 0x0020)
    except Exception:
        pass
