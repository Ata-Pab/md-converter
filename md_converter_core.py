#!/usr/bin/env python3
"""Shared conversion core for MD-Converter.

Spreadsheets (.xlsx/.xlsm/.xls/.ods) are read into a pandas DataFrame and
turned into a Markdown table (pick columns, filter rows, sort). Documents
(Word .docx, PowerPoint .pptx, OpenDocument .odt, PDF .pdf) are converted
whole into Markdown, preserving document order. Used by both the CLI
(md-converter.py) and the GUI (md-converter-ui.py) so the two cannot drift
apart.

Requires: pandas, openpyxl (xlsx), xlrd (old .xls), python-docx, python-pptx,
odfpy (odt/ods), pdfplumber.
"""

import math
import re
from collections import Counter
from pathlib import Path

import pandas as pd

DEFAULT_FILTER_WORD = "gereksinim"

SPREADSHEET_SUFFIXES = (".xlsx", ".xlsm", ".xls", ".ods")
EXCEL_SUFFIXES = SPREADSHEET_SUFFIXES  # older name, kept for compatibility

# Filter operators: key -> label shown in the GUI.
FILTER_OPS = {
    "word": "contains word",
    "contains": "contains",
    "not_contains": "does not contain",
    "equals": "equals",
    "not_equals": "does not equal",
    "starts": "starts with",
    "regex": "matches regex",
    "empty": "is empty",
    "not_empty": "is not empty",
}
VALUELESS_OPS = ("empty", "not_empty")

# Python module -> pip package, for friendly "missing package" messages.
_PIP_NAMES = {
    "docx": "python-docx",
    "pptx": "python-pptx",
    "odf": "odfpy",
    "odfpy": "odfpy",
    "pdfplumber": "pdfplumber",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
}


def describe_error(exc):
    """Readable one-line message for an exception. A missing optional
    package becomes an install hint instead of a bare ImportError."""
    if isinstance(exc, ImportError):
        name = getattr(exc, "name", None)
        if not name:
            match = re.search(r"'([\w.]+)'", str(exc))
            name = match.group(1) if match else ""
        root = (name or "").split(".")[0]
        if root:
            pkg = _PIP_NAMES.get(root, root)
            return f"Missing Python package '{pkg}'. Install it with: pip install {pkg}"
    return str(exc) or type(exc).__name__


# --------------------------------------------------------------------------- #
#  Spreadsheets: rows -> Markdown table
# --------------------------------------------------------------------------- #
def _is_missing(value):
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


# Characters that would change how plain text renders: a backslash before
# punctuation, '*', '<' opening something tag-like (the text would vanish as
# raw HTML), and '_' at a word edge (intraword '_' like snake_case is safe).
_MD_INLINE = re.compile(r"\\(?=[!-/:-@\[-`{-~])|\*|<(?=[A-Za-z/!?])|(?<![^\W_])_|_(?![^\W_])")
# A paragraph starting like this would turn into a heading or a quote.
_MD_BLOCK_START = re.compile(r"(?:#{1,6}(?=\s|$)|>)")


def escape_inline(text):
    """Escape the Markdown characters in plain text that would otherwise
    become emphasis or vanish as HTML."""
    return _MD_INLINE.sub(lambda m: "\\" + m.group(0), text)


def escape_block(text):
    """escape_inline, plus a leading '#' or '>' that would make the text a
    heading or a quote."""
    text = escape_inline(text)
    return "\\" + text if _MD_BLOCK_START.match(text) else text


def _cell_text(value, keep_line_breaks=False, escape=True):
    """Table-cell text with Markdown characters and pipes escaped. Newlines
    become spaces, or <br> when `keep_line_breaks` is set (a Markdown cell
    cannot hold a newline). `escape=False` is for text that is already
    Markdown."""
    if _is_missing(value):
        return ""
    text = str(value)
    if escape:
        text = escape_inline(text)
    text = text.replace("|", "\\|")
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")
    if keep_line_breaks:
        return "<br>".join(line.strip() for line in text.split("\n")).strip()
    return text.replace("\n", " ").strip()


def escape_md(text):
    """Escape pipe chars and collapse newlines so cell content
    doesn't break the Markdown table layout."""
    return _cell_text(text)


def find_column(columns, target):
    """Return the actual column name matching `target` (an exact match
    first, then case-insensitive and whitespace-tolerant), or None."""
    if target is None:
        return None
    columns = list(columns)
    if target in columns:
        return target
    target_norm = str(target).strip().lower()
    for col in columns:
        if str(col).strip().lower() == target_norm:
            return col
    return None


def find_columns(df, names):
    """Return {name: actual_column} for every required name, or raise
    ValueError listing what is missing and which columns do exist."""
    found = {name: find_column(df.columns, name) for name in names}
    missing = [name for name, col in found.items() if col is None]
    if missing:
        raise ValueError(
            f"could not find required column(s) {missing}. "
            f"Available columns: {list(df.columns)}"
        )
    return found


def _fold(text):
    """Lowercase for comparison. Turkish İ/ı are folded to 'i' so
    'GEREKSİNİM' matches 'gereksinim' and 'KIRMIZI' matches 'kırmızı'
    (plain casefold keeps the dot and maps I to i, not ı)."""
    return str(text).replace("İ", "i").casefold().replace("ı", "i")


def word_pattern(value):
    """Regex matching `value` folded, at a word start, any suffix allowed
    (so 'gereksinim' also matches 'gereksinimler'). The leading boundary is
    dropped when the value does not start with a word character ('.NET')."""
    left = r"\b" if re.match(r"\w", value) else ""
    return rf"{left}{re.escape(value)}"


def _column_text(df, column):
    """A column as plain strings, empty cells as '' (never 'nan')."""
    return df[column].map(lambda v: "" if _is_missing(v) else str(v))


def rule_mask(df, rule):
    """Boolean mask for one filter rule {column, op, value}, or None when
    the rule keeps every row (no column, or an empty / '*' value)."""
    column = rule.get("column")
    op = rule.get("op") or "word"
    value = str(rule.get("value") or "")
    if not column:
        return None
    if column not in df.columns:
        raise ValueError(f"filter column '{column}' not found in the sheet")
    if op not in FILTER_OPS:
        raise ValueError(f"unknown filter operator '{op}'")

    # Plain Python string tests (not pandas .str): the same Unicode-aware
    # behaviour whatever string backend pandas uses (e.g. pyarrow's RE2,
    # whose \b only knows ASCII letters).
    text = _column_text(df, column)

    def mask(test, fold=True):
        return text.map(lambda s: bool(test(_fold(s) if fold else s))).astype(bool)

    if op == "empty":
        return mask(lambda s: not s.strip(), fold=False)
    if op == "not_empty":
        return mask(lambda s: bool(s.strip()), fold=False)

    needle = _fold(value).strip()
    if not needle or (needle == "*" and op in ("word", "contains")):
        return None
    if op == "word":
        pattern = re.compile(word_pattern(needle))
        return mask(pattern.search)
    if op == "contains":
        return mask(lambda s: needle in s)
    if op == "not_contains":
        return mask(lambda s: needle not in s)
    if op == "equals":
        return mask(lambda s: s.strip() == needle)
    if op == "not_equals":
        return mask(lambda s: s.strip() != needle)
    if op == "starts":
        return mask(lambda s: s.strip().startswith(needle))
    # regex
    try:
        pattern = re.compile(value.strip(), re.IGNORECASE)
    except re.error as e:
        raise ValueError(f"invalid regular expression '{value}': {e}") from None
    return mask(pattern.search, fold=False)


def apply_filters(df, rules, match="all"):
    """Rows matching all (match='all') or any (match='any') of `rules`.
    Rules that keep every row are ignored; no active rule keeps all rows."""
    masks = [m for m in (rule_mask(df, r) for r in rules or []) if m is not None]
    if not masks:
        return df.copy()
    combined = masks[0]
    for mask in masks[1:]:
        combined = (combined | mask) if match == "any" else (combined & mask)
    return df[combined].copy()


def filter_rows(df, filter_col, filter_value):
    """Rows whose `filter_col` contains `filter_value` at a word start
    (case-insensitive). Empty value or "*" keeps every row; a falsy
    `filter_col` keeps every row."""
    if not filter_col:
        return df.copy()
    return apply_filters(df, [{"column": filter_col, "op": "word", "value": filter_value}])


def _natural_key(value):
    """Sort key that orders 'REQ-2' before 'REQ-10'."""
    parts = re.split(r"([0-9]+)", _fold(value).strip())
    return tuple((0, int(p), "") if p.isdigit() and p.isascii() else (1, 0, p)
                 for p in parts if p)


def sort_rows(df, column, descending=False):
    """Rows sorted by `column` in natural order; empty cells always last."""
    if not column:
        return df
    if column not in df.columns:
        raise ValueError(f"sort column '{column}' not found in the sheet")
    texts = _column_text(df, column).tolist()
    filled = [i for i, t in enumerate(texts) if t.strip()]
    empty = [i for i, t in enumerate(texts) if not t.strip()]
    filled.sort(key=lambda i: _natural_key(texts[i]), reverse=descending)
    return df.iloc[filled + empty]


def select_rows(df, columns, rules=None, match="all", sort_by=None,
                descending=False, skip_empty_rows=True):
    """Filter, drop rows whose output cells are all empty, then sort."""
    rows = apply_filters(df, rules, match)
    present = [c for c in columns if c in rows.columns]
    if skip_empty_rows and present and len(rows):
        frame = rows[present]
        cellwise = frame.map if hasattr(frame, "map") else frame.applymap  # pandas < 2.1
        filled = cellwise(lambda v: not _is_missing(v) and str(v).strip() != "")
        rows = rows[filled.any(axis=1)]
    return sort_rows(rows, sort_by, descending)


def table_to_markdown(df, columns, *, rules=None, match="all", sort_by=None,
                      descending=False, title="Requirements", show_count=True,
                      count_label="Total requirements", keep_line_breaks=False,
                      skip_empty_rows=True, row_numbers=False, limit=None):
    """Markdown table of the selected rows. Returns (markdown, row_count).
    `limit` renders only the first N rows (for previews); the count and the
    heading still describe every matching row."""
    columns = list(columns)
    if not columns:
        raise ValueError("select at least one column to extract")
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"column(s) {missing} not found in the sheet. "
                         f"Available columns: {list(df.columns)}")

    rows = select_rows(df, columns, rules, match, sort_by, descending, skip_empty_rows)
    headers = (["#"] if row_numbers else []) + [_cell_text(c) for c in columns]
    lines = []
    if title:
        lines += [f"# {title}", ""]
    if show_count:
        lines += [f"**{count_label or 'Total rows'}:** {len(rows)}", ""]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("|" + "---|" * len(headers))
    data = rows[columns] if limit is None else rows[columns].head(limit)
    for n, values in enumerate(data.itertuples(index=False, name=None), 1):
        cells = [_cell_text(v, keep_line_breaks) for v in values]
        if row_numbers:
            cells.insert(0, str(n))
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines) + "\n", len(rows)


def build_markdown(df, extract_cols, filter_col=None, filter_value="",
                   title="Requirements", **options):
    """Build a Markdown requirements table from a DataFrame. `filter_col` /
    `filter_value` add a 'contains word' rule in front of any `rules`
    given in `options` (see table_to_markdown for the other options)."""
    rules = list(options.pop("rules", None) or [])
    if filter_col:
        rules.insert(0, {"column": filter_col, "op": "word", "value": filter_value})
    md, _ = table_to_markdown(df, extract_cols, rules=rules, title=title, **options)
    return md


def sheet_names(path):
    """Sheet names of a workbook. The file is closed afterwards so it is not
    left locked on Windows."""
    with pd.ExcelFile(path) as xl:
        return [str(n) for n in xl.sheet_names]


def read_excel(path, sheet=0, header_row=1):
    """Read one sheet as a DataFrame of strings, columns as str. `sheet` is
    a sheet name or a 0-based index (also as a string, e.g. from the CLI);
    a sheet literally named '2024' wins over index 2024. `header_row` is the
    1-based row holding the column names."""
    with pd.ExcelFile(path) as xl:
        names = [str(n) for n in xl.sheet_names]
        if isinstance(sheet, str) and sheet in names:
            index = names.index(sheet)
        else:
            try:
                index = int(sheet)
            except (TypeError, ValueError):
                raise ValueError(f"sheet '{sheet}' not found. "
                                 f"Available sheets: {names}") from None
            if not 0 <= index < len(names):
                raise ValueError(f"sheet index {index} is out of range. "
                                 f"Available sheets: {names}")
        header = max(int(header_row or 1), 1) - 1
        try:
            df = xl.parse(sheet_name=xl.sheet_names[index], header=header, dtype=str)
        except ValueError as e:
            if header and "header" in str(e).lower():  # header row past the data
                raise ValueError(f"the sheet has fewer than {header + 1} rows") from None
            raise
    df.columns = _unique_names(str(c) for c in df.columns)
    return df


def _unique_names(names):
    """Column names made unique after str(): pandas only de-duplicates the
    raw header values, so a numeric 1 and a text '1' would both be '1'."""
    used, out = set(), []
    for name in names:
        candidate, n = name, 1
        while candidate in used:
            candidate = f"{name}.{n}"
            n += 1
        used.add(candidate)
        out.append(candidate)
    return out


# --------------------------------------------------------------------------- #
#  Shared helpers for documents
# --------------------------------------------------------------------------- #
def _cell_md(text, escape=True):
    """Table-cell text: escape pipes and keep line breaks as <br>, since a
    Markdown table cell cannot contain a real newline."""
    if text is None:
        return ""
    return _cell_text(text, keep_line_breaks=True, escape=escape)


def _rows_to_md(rows, escape=True):
    """Render a list of rows (each a list of strings) as a Markdown table.
    Fully empty rows and trailing empty columns are dropped. `escape=False`
    when the cells already hold Markdown (bold, links)."""
    rows = [[_cell_md(c, escape) for c in row] for row in rows]
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [(r + [""] * width)[:width] for r in rows]
    while width > 1 and all(not r[width - 1].strip() for r in rows):
        rows = [r[:width - 1] for r in rows]
        width -= 1
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _wrap_marker(text, marker):
    """`**text**` with surrounding spaces kept outside the markers: Markdown
    ignores '** bold **', so a run like 'bold ' must become '**bold** '."""
    core = text.strip()
    if not core or not marker:
        return text
    lead = text[:len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    return f"{lead}{marker}{core}{marker}{trail}"


def _link(text, url):
    core = text.strip()
    if not core:
        return text
    lead = text[:len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    core = core.replace("[", "\\[").replace("]", "\\]")
    target = f"<{url}>" if re.search(r"[\s()<>]", url) else url
    return f"{lead}[{core}]({target}){trail}"


def _segments_to_md(segments, keep_formatting=True, line_break=" "):
    """Inline Markdown from (text, bold, italic, url) segments. Adjacent
    segments with the same formatting are merged so bold text does not come
    out as '**a****b**'. Manual line breaks become `line_break`."""
    def clean(text):
        text = escape_inline(text)
        text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")
        return text.replace("\n", line_break).replace("\t", " ")

    segments = [(clean(t), b, i, u) for t, b, i, u in segments if t]
    if not keep_formatting:
        return "".join(t for t, *_ in segments).strip()

    out = []
    pos = 0
    while pos < len(segments):
        url = segments[pos][3]
        group = []
        while pos < len(segments) and segments[pos][3] == url:
            group.append(segments[pos])
            pos += 1
        parts, buffer, style = [], [], None

        def flush():
            if buffer:
                bold, italic = style
                marker = "***" if bold and italic else "**" if bold else "*" if italic else ""
                parts.append(_wrap_marker("".join(buffer), marker))

        for text, bold, italic, _ in group:
            if (bold, italic) != style:
                flush()
                buffer, style = [], (bold, italic)
            buffer.append(text)
        flush()
        text = "".join(parts)
        out.append(_link(text, url) if url else text)
    return "".join(out).strip()


def _escape_start(md):
    """Escape a leading '#' or '>' that would turn a paragraph into a
    heading or a quote."""
    return "\\" + md if _MD_BLOCK_START.match(md) else md


def _list_levels(levels):
    """Valid nesting for a run of list items: the first item is top level
    and no item is more than one level deeper than the one before it.
    Otherwise Markdown reads the item as code or as text of its parent."""
    out, prev = [], -1
    for level in levels:
        prev = max(0, min(level, prev + 1))
        out.append(prev)
    return out


def _join_blocks(blocks):
    """Join (kind, markdown) blocks: consecutive list items on adjacent lines
    (a tight list), everything else separated by a blank line. A list item
    is one line indented 4 spaces per level; levels are normalised so the
    Markdown nests correctly. A top-level switch between bullets and numbers
    starts a new list after a blank line."""
    out = []
    prev, list_type, level = None, None, -1
    for kind, md in blocks:
        if not md.strip():
            continue
        tight = False
        if kind == "list":
            item = md.lstrip(" ")
            wanted = (len(md) - len(item)) // 4
            level = max(0, min(wanted, level + 1)) if prev == "list" else 0
            marker = "ordered" if item[:1].isdigit() else "bullet"
            tight = prev == "list" and not (level == 0 and marker != list_type)
            if level == 0:
                list_type = marker
            md = "    " * level + item
        if out:
            out.append("\n" if tight else "\n\n")
        out.append(md)
        prev = kind
    return "".join(out).strip()


def _finish(body, title=None, header="", footer=""):
    parts = []
    if title:
        parts.append(f"# {title}")
    if header:
        parts += [header, "---"]
    parts.append(body)
    if footer:
        parts += ["---", footer]
    return "\n\n".join(p for p in parts if p.strip()).strip() + "\n"


def shift_headings(md, by=1):
    """Push every ATX heading down `by` levels (max level 6), skipping
    fenced code blocks. Used when several files are merged under one
    top-level heading each."""
    out, fenced = [], False
    for line in md.split("\n"):
        if line.lstrip().startswith("```"):
            fenced = not fenced
        elif not fenced:
            match = re.match(r"^(#{1,6})(?=\s)", line)
            if match:
                level = min(6, len(match.group(1)) + by)
                line = "#" * level + line[len(match.group(1)):]
        out.append(line)
    return "\n".join(out)


def merge_markdown(parts):
    """One Markdown document from [(heading, markdown), ...]: each part under
    its own '# heading', with its own headings pushed one level down."""
    sections = [f"# {name}\n\n{shift_headings(md).strip()}" for name, md in parts]
    return "\n\n".join(sections).strip() + "\n"


# --------------------------------------------------------------------------- #
#  Word: whole document -> Markdown
# --------------------------------------------------------------------------- #
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _w(tag):
    return f"{{{_W_NS}}}{tag}"


# Inline containers whose runs are part of the visible text (tracked
# insertions, content controls, simple fields, ...). Deleted text (w:del,
# w:moveFrom) is skipped.
_DOCX_INLINE_CONTAINERS = {
    _w(t) for t in ("ins", "moveTo", "smartTag", "sdt", "sdtContent",
                    "fldSimple", "customXml", "dir", "bdo")
}


class _PartHolder:
    """Minimal parent for python-docx proxies: they only need `.part`."""

    def __init__(self, part):
        self.part = part


class _DocxNumbering:
    """Looks up whether a list level is bulleted or numbered."""

    def __init__(self, document_part):
        try:
            self._el = document_part.numbering_part.element
        except Exception:  # no numbering part in this document
            self._el = None
        self._cache = {}

    def num_format(self, num_id, ilvl):
        key = (num_id, ilvl)
        if key not in self._cache:
            self._cache[key] = self._lookup(num_id, ilvl)
        return self._cache[key]

    def _lookup(self, num_id, ilvl):
        if self._el is None:
            return None
        nums = self._el.xpath(f'w:num[@w:numId="{num_id}"]')
        if not nums:
            return None
        override = nums[0].xpath(f'w:lvlOverride[@w:ilvl="{ilvl}"]/w:lvl/w:numFmt/@w:val')
        if override:
            return override[0]
        abstract = nums[0].xpath("w:abstractNumId/@w:val")
        if not abstract:
            return None
        fmt = self._el.xpath(
            f'w:abstractNum[@w:abstractNumId="{abstract[0]}"]'
            f'/w:lvl[@w:ilvl="{ilvl}"]/w:numFmt/@w:val'
        )
        return fmt[0] if fmt else None


class _DocxContext:
    def __init__(self, doc, keep_formatting=True):
        self.numbering = _DocxNumbering(doc.part)
        self.keep_formatting = keep_formatting


def _docx_runs(container, part, url=None):
    """Yield (w:r element, link url) for the visible runs of a paragraph,
    including runs inside hyperlinks, tracked insertions and content
    controls (python-docx's `paragraph.runs` skips all of those)."""
    for child in container.iterchildren():
        tag = child.tag
        if tag == _w("r"):
            yield child, url
        elif tag == _w("hyperlink"):
            link = url
            rid = child.get(_R_ID)
            if rid and rid in part.rels and part.rels[rid].is_external:
                link = part.rels[rid].target_ref
            yield from _docx_runs(child, part, link)
        elif tag in _DOCX_INLINE_CONTAINERS:
            yield from _docx_runs(child, part, url)


def _docx_inline(paragraph, ctx, plain=False, line_break=" "):
    from docx.text.run import Run

    segments = []
    for r, url in _docx_runs(paragraph._p, paragraph.part):
        run = Run(r, paragraph)
        text = run.text
        if not text:
            continue
        bold, italic = run.bold, run.italic
        if (bold is None or italic is None) and run.style is not None:
            font = run.style.font
            bold = font.bold if bold is None else bold
            italic = font.italic if italic is None else italic
        segments.append((text, bool(bold), bool(italic), url))
    return _segments_to_md(segments, ctx.keep_formatting and not plain, line_break)


def _style_chain(paragraph):
    style = paragraph.style
    for _ in range(12):  # guard against style loops
        if style is None:
            return
        yield style
        style = style.base_style


def _heading_level(paragraph):
    """Heading level 1-6 for a paragraph, or 0 if it is not a heading.
    Uses the style name ('Heading 2', Turkish 'Başlık 2', 'Title') and falls
    back to an explicit outline level."""
    name = (paragraph.style.name or "").strip() if paragraph.style is not None else ""
    match = re.match(r"(?:Heading|Başlık)\s*(\d+)", name, re.IGNORECASE)
    if match:
        return max(1, min(int(match.group(1)), 6))
    if name.lower() == "title":
        return 1
    pPr = paragraph._p.pPr
    candidates = [pPr] + [s.element.pPr for s in _style_chain(paragraph)]
    for ppr in candidates:
        if ppr is None:
            continue
        lvl = ppr.find(_w("outlineLvl"))
        if lvl is not None:
            value = int(lvl.get(_w("val"), "9"))
            return min(value + 1, 6) if value < 9 else 0
    return 0


def _list_info(paragraph, ctx):
    """(indent level, ordered) for a list paragraph, or None. Numbering may
    sit on the paragraph or on its style; the numbering definitions decide
    between bullets and numbers."""
    name = (paragraph.style.name or "").lower() if paragraph.style is not None else ""
    num_id = ilvl = None
    from_style = False
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.numPr is not None:
        num_pr = pPr.numPr
        num_id = num_pr.numId.val if num_pr.numId is not None else None
        ilvl = num_pr.ilvl.val if num_pr.ilvl is not None else None
    if num_id is None:
        for style in _style_chain(paragraph):
            spPr = style.element.pPr
            if spPr is not None and spPr.numPr is not None:
                num_pr = spPr.numPr
                if num_pr.numId is not None:
                    num_id = num_pr.numId.val
                    from_style = True
                    if ilvl is None and num_pr.ilvl is not None:
                        ilvl = num_pr.ilvl.val
                    break

    named_list = "list" in name and ("bullet" in name or "number" in name)
    if num_id is None:
        if not named_list:
            return None
        ordered, level = "number" in name, 0
    elif num_id == 0:  # numbering explicitly removed
        return None
    else:
        level = ilvl or 0
        fmt = ctx.numbering.num_format(num_id, level)
        ordered = ("number" in name) if fmt is None else fmt not in ("bullet", "none")
    if (from_style or num_id is None) and named_list:
        # 'List Bullet 2' is a second-level list even when its style
        # numbering says level 0.
        match = re.search(r"(\d+)\s*$", name)
        if match:
            level = max(level, int(match.group(1)) - 1)
    return min(level, 8), ordered


_DOCX_WRAPPERS = (_w("sdt"), _w("sdtContent"), _w("customXml"))


def _docx_children(parent, tag):
    """Children of `parent` with `tag`, looking through content controls
    and custom-XML wrappers."""
    for child in parent.iterchildren():
        if child.tag == tag:
            yield child
        elif child.tag in _DOCX_WRAPPERS:
            yield from _docx_children(child, tag)


def _docx_int(element, path, default):
    """Integer w:val of the descendant at `path` ('trPr/gridBefore')."""
    found = element.find("/".join(_w(p) for p in path.split("/")))
    try:
        return int(found.get(_w("val"))) if found is not None else default
    except (TypeError, ValueError):
        return default


def _docx_cell_lines(container, holder, ctx):
    """Text lines of a table cell; a nested table is flattened to one line
    per row."""
    from docx.text.paragraph import Paragraph

    lines = []
    for child in container.iterchildren():
        if child.tag == _w("p"):
            lines.append(_docx_inline(Paragraph(child, holder), ctx, line_break="\n"))
        elif child.tag == _w("tbl"):
            for tr in _docx_children(child, _w("tr")):
                texts = [" ".join(_docx_cell_lines(tc, holder, ctx))
                         for tc in _docx_children(tr, _w("tc"))]
                lines.append(" / ".join(t for t in texts if t))
        elif child.tag in _DOCX_WRAPPERS:
            lines.extend(_docx_cell_lines(child, holder, ctx))
    return [line for line in lines if line]


def _docx_table_to_md(tbl, part, ctx):
    """Markdown table from a w:tbl element, read straight from the XML so
    rows that start late (gridBefore), merged cells and content controls all
    keep their columns. A cell spanning several columns is written once with
    the spanned columns left empty; the continuation of a vertical merge is
    empty in the file and stays empty."""
    holder = _PartHolder(part)
    rows = []
    for tr in _docx_children(tbl, _w("tr")):
        cells = [""] * _docx_int(tr, "trPr/gridBefore", 0)
        for tc in _docx_children(tr, _w("tc")):
            cells.append("\n".join(_docx_cell_lines(tc, holder, ctx)))
            cells.extend([""] * (_docx_int(tc, "tcPr/gridSpan", 1) - 1))
        rows.append(cells)
    return _rows_to_md(rows, escape=False)  # cell text is Markdown already


def _docx_blocks(container, part, ctx):
    """(kind, markdown) blocks for the paragraphs and tables of a body,
    header or footer element, in order."""
    from docx.text.paragraph import Paragraph

    holder = _PartHolder(part)
    blocks = []
    for child in container.iterchildren():
        tag = child.tag
        if tag == _w("p"):
            paragraph = Paragraph(child, holder)
            level = _heading_level(paragraph)
            if level:
                text = _docx_inline(paragraph, ctx, plain=True)
                if text:
                    blocks.append(("heading", "#" * level + " " + text))
                continue
            text = _docx_inline(paragraph, ctx)
            if not text:
                continue
            info = _list_info(paragraph, ctx)
            if info:
                indent, ordered = info
                blocks.append(("list", "    " * indent + ("1. " if ordered else "- ") + text))
            else:
                blocks.append(("text", _escape_start(text)))
        elif tag == _w("tbl"):
            try:
                blocks.append(("table", _docx_table_to_md(child, part, ctx)))
            except Exception:  # one odd table must not sink the whole document
                blocks.append(("text", "*[A table here could not be converted.]*"))
        elif tag in _DOCX_WRAPPERS:
            blocks.extend(_docx_blocks(child, part, ctx))
    return blocks


def _docx_header_footer(doc, kind, ctx):
    """Markdown of every distinct header (kind='header') or footer
    (kind='footer') used in the document: first-page, default and even-page
    variants of every section, duplicates removed."""
    seen, out = set(), []
    even = doc.settings.odd_and_even_pages_header_footer
    for section in doc.sections:
        variants = []
        if section.different_first_page_header_footer:
            variants.append(getattr(section, f"first_page_{kind}"))
        variants.append(getattr(section, kind))
        if even:
            variants.append(getattr(section, f"even_page_{kind}"))
        for hf in variants:
            if hf.is_linked_to_previous:  # no own definition
                continue
            md = _join_blocks(_docx_blocks(hf._element, hf.part, ctx))
            if md and md not in seen:
                seen.add(md)
                out.append(md)
    return "\n\n".join(out)


def docx_to_markdown(path, title=None, include_header_footer=False,
                     keep_formatting=True, **_):
    """Convert a whole Word document to Markdown (headings, paragraphs,
    lists, tables, links) in document order. Headers and footers are added
    above / below the body when `include_header_footer` is set. Images,
    text boxes, footnotes and comments are not emitted."""
    from docx import Document

    doc = Document(str(path))
    ctx = _DocxContext(doc, keep_formatting)
    body = _join_blocks(_docx_blocks(doc.element.body, doc.part, ctx))
    header = footer = ""
    if include_header_footer:
        header = _docx_header_footer(doc, "header", ctx)
        footer = _docx_header_footer(doc, "footer", ctx)
    return _finish(body, title, header, footer)


# --------------------------------------------------------------------------- #
#  PowerPoint: whole presentation -> Markdown
# --------------------------------------------------------------------------- #
def _pptx_paragraph_md(paragraph, keep_formatting):
    """Inline Markdown for one PowerPoint paragraph, keeping line breaks
    (as spaces), fields, bold/italic and hyperlinks."""
    from pptx.text.text import _Run

    segments = []
    for child in paragraph._p.iterchildren():
        local = child.tag.rsplit("}", 1)[-1]
        if local == "r":
            run = _Run(child, paragraph)
            try:
                url = run.hyperlink.address
            except Exception:
                url = None
            segments.append((run.text, bool(run.font.bold), bool(run.font.italic), url))
        elif local == "br":
            segments.append((" ", False, False, None))
        elif local == "fld":
            text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
            segments.append((text, False, False, None))
    return _segments_to_md(segments, keep_formatting)


def _text_frame_to_md(text_frame, keep_formatting=True):
    """Bullet list from a PowerPoint text frame, keeping indentation level
    (normalised so the Markdown nests correctly)."""
    items = []
    for paragraph in text_frame.paragraphs:
        text = _pptx_paragraph_md(paragraph, keep_formatting)
        if text:
            items.append((paragraph.level or 0, text))
    levels = _list_levels([level for level, _ in items])
    return "\n".join("    " * level + "- " + text
                     for level, (_, text) in zip(levels, items))


def _table_to_md(table):
    """Markdown table from a PowerPoint table; cells covered by a merge are
    left empty."""
    rows = []
    for row in table.rows:
        rows.append(["" if getattr(cell, "is_spanned", False) else cell.text
                     for cell in row.cells])
    return _rows_to_md(rows)


def _iter_pptx_shapes(shapes):
    """Iterate a shape collection, skipping elements python-pptx cannot
    turn into a shape (e.g. ink content parts) instead of failing."""
    try:
        elements = list(shapes._iter_member_elms())
        factory = shapes._shape_factory
    except Exception:  # other python-pptx internals: plain iteration
        yield from shapes
        return
    for element in elements:
        try:
            yield factory(element)
        except Exception:
            continue


def _pptx_position(shape):
    try:
        return (shape.top or 0, shape.left or 0)
    except Exception:
        return (0, 0)


def _pptx_shapes(shapes):
    """Shapes in reading order (top to bottom, then left to right), with
    group shapes flattened."""
    from pptx.shapes.group import GroupShape

    flat = []
    for shape in _iter_pptx_shapes(shapes):
        if isinstance(shape, GroupShape):
            flat.extend(_pptx_shapes(shape.shapes))
        else:
            flat.append(shape)
    return sorted(flat, key=_pptx_position)


def _pptx_slide_title(slide):
    try:
        return slide.shapes.title
    except Exception:
        return None


def _is_pptx_footer(shape):
    from pptx.enum.shapes import PP_PLACEHOLDER

    if not shape.is_placeholder:
        return False
    try:
        kind = shape.placeholder_format.type
    except ValueError:
        return False
    return kind in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE)


def pptx_to_markdown(path, title=None, include_notes=True,
                     include_header_footer=False, keep_formatting=True, **_):
    """Convert a whole PowerPoint presentation to Markdown: one section per
    slide, the title as a heading, body text as bullets, tables as tables and
    speaker notes as blockquotes. Footer / date / slide-number boxes are
    only kept with `include_header_footer`. Images are not emitted."""
    from pptx import Presentation

    prs = Presentation(str(path))
    blocks = []
    for i, slide in enumerate(prs.slides, 1):
        title_shape = _pptx_slide_title(slide)
        heading = " ".join(title_shape.text.split()) if title_shape is not None else ""
        blocks.append(f"## Slide {i}" + (f": {escape_inline(heading)}" if heading else ""))

        title_id = title_shape.shape_id if title_shape is not None else None
        skipped = False
        for shape in _pptx_shapes(slide.shapes):
            try:
                if shape.shape_id == title_id:
                    continue
                if not include_header_footer and _is_pptx_footer(shape):
                    continue
                if getattr(shape, "has_table", False) and shape.has_table:
                    md = _table_to_md(shape.table)
                elif getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                    md = _text_frame_to_md(shape.text_frame, keep_formatting)
                else:
                    continue
            except Exception:  # one odd shape must not sink the presentation
                skipped = True
                continue
            if md:
                blocks.append(md)
        if skipped:
            blocks.append("*[Part of this slide could not be converted.]*")

        if include_notes and slide.has_notes_slide:
            notes_frame = slide.notes_slide.notes_text_frame
            notes = notes_frame.text.strip() if notes_frame is not None else ""
            if notes:
                blocks.append("\n".join(("> " + escape_inline(line)).rstrip()
                                        for line in notes.splitlines()))

    return _finish("\n\n".join(b for b in blocks if b.strip()), title)


# --------------------------------------------------------------------------- #
#  OpenDocument (.odt): whole document -> Markdown
# --------------------------------------------------------------------------- #
def _node_name(node):
    """Normalized local element name in an odfpy tree: lowercase, dashes
    removed (e.g. 'list-item' -> 'listitem', 'table-cell' -> 'tablecell')."""
    return (getattr(node, "qname", (None, None))[1] or "").replace("-", "").lower()


# Inline elements whose text is not part of the visible paragraph: reviewer
# comments, footnote / endnote bodies, tracked-change records.
_ODT_SKIP = {"annotation", "annotationend", "note", "trackedchanges", "changedregion"}


def _odt_collect(node, out):
    for child in node.childNodes:
        if getattr(child, "nodeType", None) == 3:  # text node
            out.append(child.data)
            continue
        name = _node_name(child)
        if name in _ODT_SKIP:
            continue
        if name == "s":
            out.append(" " * int(child.getAttribute("c") or 1))
        elif name in ("tab", "linebreak"):
            out.append(" ")
        else:
            _odt_collect(child, out)


def _odt_text(node):
    """Visible text of a paragraph / heading, whitespace collapsed (not
    Markdown-escaped)."""
    out = []
    _odt_collect(node, out)
    return " ".join("".join(out).split())


# Containers whose children are ordinary body content.
_ODT_CONTAINERS = {"section", "tableofcontent", "indexbody", "alphabeticalindex",
                   "illustrationindex", "tableindex", "objectindex", "userindex",
                   "bibliography", "indextitle"}


class _OdtLists:
    """Knows which list styles number their items (1. 2. 3.) rather than
    bullet them."""

    def __init__(self, doc):
        self.styles = {}
        for container in (doc.automaticstyles, doc.styles):
            for node in getattr(container, "childNodes", []):
                if _node_name(node) != "liststyle":
                    continue
                levels = {}
                for lvl in node.childNodes:
                    kind = _node_name(lvl)
                    if kind.startswith("listlevelstyle"):
                        levels[int(lvl.getAttribute("level") or 1)] = kind == "listlevelstylenumber"
                self.styles.setdefault(node.getAttribute("name"), levels)

    def ordered(self, style, level):
        return self.styles.get(style, {}).get(level + 1, False)


def _odt_heading(node):
    text = _odt_text(node)
    if not text:
        return None
    level = int(node.getAttribute("outlinelevel") or 1)
    return ("heading", "#" * max(1, min(level, 6)) + " " + escape_inline(text))


def _odt_list_blocks(list_node, lists, depth=0, level=0, style=None):
    """Blocks for a list. `depth` is the nesting in the file (it picks the
    bullet or number style); `level` is the Markdown nesting, which only
    grows below an item that actually printed a line. Headings inside list
    items (numbered headings) come out as headings."""
    style = list_node.getAttribute("stylename") or style
    marker = "1. " if lists.ordered(style, depth) else "- "
    blocks = []
    for item in list_node.childNodes:
        if _node_name(item) not in ("listitem", "listheader"):
            continue
        texts = []
        for child in item.childNodes:
            name = _node_name(child)
            if name == "h":
                heading = _odt_heading(child)
                if heading:
                    blocks.append(heading)
            elif name == "p":
                texts.append(_odt_text(child))
        text = " ".join(t for t in texts if t)
        if text:
            blocks.append(("list", "    " * level + marker + escape_inline(text)))
        for child in item.childNodes:
            if _node_name(child) == "list":
                blocks.extend(_odt_list_blocks(child, lists, depth + 1,
                                               level + 1 if text else level, style))
    return blocks


def _odt_cell_lines(node):
    """Plain-text lines of a table cell: one per paragraph, list items as
    '- item' (escaping happens when the table is rendered)."""
    lines = []
    for child in node.childNodes:
        name = _node_name(child)
        if name in ("p", "h"):
            lines.append(("- " if _node_name(node) in ("listitem", "listheader") else "")
                         + _odt_text(child))
        elif name in ("list", "listitem", "listheader"):
            lines.extend(_odt_cell_lines(child))
    return [line for line in lines if line.strip("- ")]


def _odt_cell_text(cell):
    return "\n".join(_odt_cell_lines(cell))


def _odt_table_rows(node, rows):
    for child in node.childNodes:
        name = _node_name(child)
        if name in ("tableheaderrows", "tablerows", "tablerowgroup"):
            _odt_table_rows(child, rows)
        elif name == "tablerow":
            cells = []
            for tc in child.childNodes:
                kind = _node_name(tc)
                if kind not in ("tablecell", "coveredtablecell"):
                    continue
                text = "" if kind == "coveredtablecell" else _odt_cell_text(tc)
                # Trailing filler cells can repeat ~1000 times; real repeated
                # content is rarely that wide, so cap the expansion.
                repeat = int(tc.getAttribute("numbercolumnsrepeated") or 1)
                cells.extend([text] * min(repeat, 64))
            rows.append(cells)


def _odt_table_to_md(table):
    rows = []
    _odt_table_rows(table, rows)
    return _rows_to_md(rows)


def _odt_blocks(node, lists):
    blocks = []
    for child in node.childNodes:
        name = _node_name(child)
        if name == "h":
            heading = _odt_heading(child)
            if heading:
                blocks.append(heading)
        elif name == "p":
            text = _odt_text(child)
            if text:
                blocks.append(("text", _escape_start(escape_inline(text))))
        elif name == "list":
            blocks.extend(_odt_list_blocks(child, lists))
        elif name == "table":
            blocks.append(("table", _odt_table_to_md(child)))
        elif name in _ODT_CONTAINERS:
            blocks.extend(_odt_blocks(child, lists))
    return blocks


def _odt_header_footer(doc, lists):
    headers, footers, seen = [], [], set()
    for master in getattr(doc.masterstyles, "childNodes", []):
        if _node_name(master) != "masterpage":
            continue
        for part in master.childNodes:
            name = _node_name(part)
            target = headers if name.startswith("header") else footers if name.startswith("footer") else None
            if target is None:
                continue
            md = _join_blocks(_odt_blocks(part, lists))
            if md and (name[:6], md) not in seen:
                seen.add((name[:6], md))
                target.append(md)
    return "\n\n".join(headers), "\n\n".join(footers)


def odt_to_markdown(path, title=None, include_header_footer=False, **_):
    """Convert a whole OpenDocument Text (.odt) file to Markdown: headings,
    paragraphs, lists and tables in document order (sections included)."""
    from odf.opendocument import load

    doc = load(str(path))
    lists = _OdtLists(doc)
    body = _join_blocks(_odt_blocks(doc.text, lists))
    header = footer = ""
    if include_header_footer:
        header, footer = _odt_header_footer(doc, lists)
    return _finish(body, title, header, footer)


# --------------------------------------------------------------------------- #
#  PDF: whole document -> Markdown (best effort)
# --------------------------------------------------------------------------- #
# Bullet glyphs at the start of a PDF line; '(cid:N)' is a glyph without a
# Unicode mapping, which in practice is almost always a bullet symbol.
_PDF_BULLET = re.compile(
    r"^(?:[\u2022\u25cf\u25aa\u25a0\u2023\u2043\u25e6\u00b7\uf0b7\u27a2]|\(cid:\d+\))\s*"
)


def _pdf_heading_level(size, body_size):
    if not body_size:
        return 0
    if size >= body_size * 1.3:
        return 1
    if size >= body_size * 1.12:
        return 2
    return 0


def _pdf_line_key(text):
    """Line text with digits folded, so 'Page 3 of 9' matches 'Page 4 of 9'."""
    return re.sub(r"\d+", "#", " ".join(text.lower().split()))


def _pdf_margin(line, height):
    return line["top"] < height * 0.1 or line["bottom"] > height * 0.9


def _pdf_running_lines(pages):
    """Keys of lines repeated in the top or bottom margin of at least half
    the pages: running headers, footers and page numbers."""
    if len(pages) < 2:
        return set()
    counts = Counter()
    for page, _, lines in pages:
        counts.update({_pdf_line_key(l["text"]) for l in lines if _pdf_margin(l, page.height)})
    threshold = max(2, math.ceil(len(pages) * 0.5))
    return {key for key, n in counts.items() if n >= threshold}


def pdf_to_markdown(path, title=None, include_header_footer=False, **_):
    """Convert a text-based PDF to Markdown (best effort). Text lines whose
    font is clearly larger than the body become headings, bullet characters
    become list items, ruled tables become Markdown tables, and paragraphs
    are split on vertical gaps. Running headers / footers (lines repeated in
    the page margins) are dropped unless `include_header_footer` is set.
    A scanned (image-only) PDF has no text layer and yields nothing."""
    import pdfplumber

    with pdfplumber.open(str(path)) as pdf:
        sizes = Counter(round(ch["size"], 1) for page in pdf.pages for ch in page.chars)
        body_size = sizes.most_common(1)[0][0] if sizes else 0

        pages = []
        for page in pdf.pages:
            try:
                tables = page.find_tables()
            except Exception:
                tables = []
            lines = [l for l in page.extract_text_lines() if l["text"].strip()]
            pages.append((page, tables, lines))
        running = set() if include_header_footer else _pdf_running_lines(pages)

        blocks = []
        for page, tables, lines in pages:
            boxes = [t.bbox for t in tables]
            items = []
            for line in lines:
                if any(b[0] <= line["x0"] and line["x1"] <= b[2]
                       and b[1] <= line["top"] and line["bottom"] <= b[3] for b in boxes):
                    continue
                if running and _pdf_margin(line, page.height) and _pdf_line_key(line["text"]) in running:
                    continue
                items.append((line["top"], 0, line))
            for n, table in enumerate(tables):
                items.append((table.bbox[1], 1, n))
            items.sort(key=lambda item: (item[0], item[1]))

            paragraph, kind, prev_bottom = [], "text", None

            def flush():
                # Wrapped lines are joined with spaces (Markdown would show
                # them that way anyway) so a line that happens to start with
                # '- ' or '#' cannot turn into a list item or heading.
                nonlocal paragraph, kind
                if paragraph:
                    text = " ".join(paragraph)
                    blocks.append((kind, text if kind == "list" else _escape_start(text)))
                paragraph, kind = [], "text"

            for top, is_table, obj in items:
                if is_table:
                    flush()
                    blocks.append(("table", _rows_to_md(tables[obj].extract() or [])))
                    prev_bottom = None
                    continue
                text = obj["text"].strip()
                size = max((c["size"] for c in obj.get("chars") or []), default=body_size)
                level = _pdf_heading_level(size, body_size) if len(text) <= 120 else 0
                if level:
                    flush()
                    blocks.append(("heading", "#" * level + " " + escape_inline(text)))
                    prev_bottom = None
                    continue
                gap = top - prev_bottom if prev_bottom is not None else 0
                if paragraph and gap > body_size * 0.6:
                    flush()
                bullet = _PDF_BULLET.match(text)
                if bullet:
                    flush()
                    paragraph, kind = ["- " + escape_inline(text[bullet.end():])], "list"
                else:
                    paragraph.append(escape_inline(text))
                prev_bottom = obj["bottom"]
            flush()

    return _finish(_join_blocks(blocks), title)


# --------------------------------------------------------------------------- #
#  Dispatch
# --------------------------------------------------------------------------- #
WHOLE_DOC_CONVERTERS = {
    ".docx": docx_to_markdown,
    ".pptx": pptx_to_markdown,
    ".odt": odt_to_markdown,
    ".pdf": pdf_to_markdown,
}
DOCUMENT_SUFFIXES = tuple(WHOLE_DOC_CONVERTERS)
SUPPORTED_SUFFIXES = SPREADSHEET_SUFFIXES + DOCUMENT_SUFFIXES

FILE_TYPES = {
    ".xlsx": "Excel", ".xlsm": "Excel", ".xls": "Excel 97", ".ods": "ODS sheet",
    ".docx": "Word", ".pptx": "PowerPoint", ".odt": "ODT text", ".pdf": "PDF",
}


def file_kind(path):
    """'sheet' for spreadsheets, 'doc' for documents, None otherwise."""
    suffix = Path(path).suffix.lower()
    if suffix in SPREADSHEET_SUFFIXES:
        return "sheet"
    if suffix in WHOLE_DOC_CONVERTERS:
        return "doc"
    return None


def document_to_markdown(path, **options):
    """Convert a whole document (Word / PowerPoint / OpenDocument / PDF) to
    Markdown, dispatching on the file extension. Options: title,
    include_header_footer, include_notes (PowerPoint), keep_formatting
    (Word / PowerPoint bold, italic and links). Options a format does not
    use are ignored."""
    converter = WHOLE_DOC_CONVERTERS.get(Path(path).suffix.lower())
    if converter is None:
        raise ValueError(f"unsupported document type: {Path(path).suffix}")
    return converter(path, **options)
