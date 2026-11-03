# MD-Converter

Convert spreadsheets and documents into clean Markdown.

- **Spreadsheets** (`.xlsx`, `.xlsm`, `.xls`, `.ods`): pick the columns to
  keep, filter and sort the rows, and get a Markdown table with a total count.
  Works with Turkish sheets too: by default rows are matched on the word
  `gereksinim` in the `Property` column.
- **Documents** (`.docx`, `.pptx`, `.odt`, `.pdf`): the whole file is
  converted in reading order: headings, paragraphs, bullet and numbered
  lists, tables, links and, for slides, speaker notes.

## What it does

- Loads one or more files; spreadsheets let you choose the sheet and the row
  that holds the column names.
- Selects which columns to extract and in what order.
- Filters rows with any number of rules (*contains word*, *contains*,
  *equals*, *starts with*, *matches regex*, *is empty*, and their opposites),
  matching **all** or **any** of them. Matching ignores case, Turkish `İ/i`
  included; `*` keeps every row.
- Sorts rows by a column in natural order (`REQ-2` before `REQ-10`).
- Table options: title, count line and its label, a `#` column, line breaks
  kept as `<br>`, skipping empty rows.
- Document options: include page headers & footers, include speaker notes,
  keep bold / italic / links, add the file name as a top heading.
- Escapes pipes and newlines so tables stay valid Markdown.
- Converts the selected files or all of them, one `.md` per file or merged
  into a single `.md` with a section per file.

## Files

| File | Description |
|---|---|
| `md-converter-ui.py` | Desktop GUI (Tkinter) with live preview. |
| `md-converter.py` | Command-line version. |
| `md_converter_core.py` | Conversion logic shared by the GUI and the CLI. |
| `requirements.txt` | Python packages to install. |

## Requirements

- Python 3.9+ (Tkinter is included with the python.org installers)
- `pandas`, `openpyxl`, `xlrd`, `python-docx`, `python-pptx`, `odfpy`, `pdfplumber`
- Optional: `tkinterdnd2` for drag & drop in the GUI

```bash
pip install -r requirements.txt
pip install tkinterdnd2   # optional: drag & drop
```

You only need the packages for the formats you use; a missing one is
reported with the exact `pip install` command.

## Usage

### GUI

```bash
python md-converter-ui.py [files...]
```

1. **Files**: *Add files…*, *Add folder…*, or drag & drop files onto the
   window (with `tkinterdnd2`). Click a file to see its options and preview.
2. **Options**
   - Spreadsheet: choose the sheet and header row, move columns from
     *Available* to *In the table* (double-click, or the arrow buttons; ▲ ▼ to
     reorder), add filter rules, and set sorting and table options. Each
     spreadsheet keeps its own settings; a newly added one starts from the
     last settings you used.
   - Document: the options apply to every document in the list.
3. **Output**: save next to each source file or into a folder, optionally
   with a custom file name (single file) or merged into one file.
4. Check the **Preview** on the right, then click **Convert selected** or
   **Convert all**. Results and errors are listed in the **Log** tab.

Shortcuts: `Ctrl+O` add files, `Ctrl+Shift+O` add folder, `Delete` remove,
`Ctrl+Enter` convert selected, `Ctrl+Shift+Enter` convert all, `F5` reload the
file from disk. Right-click a file for more actions.

Your choices (output folder, document options, last spreadsheet settings,
window size) are remembered in `%APPDATA%\MD-Converter\settings.json`
(`~/.config/md-converter/settings.json` on macOS/Linux).

### Command line

```bash
python md-converter.py input [output.md] [options]
```

Examples:

```bash
python md-converter.py specs.xlsx
python md-converter.py specs.xlsx requirements.md --sheet "Sheet1"
python md-converter.py specs.xlsx --columns "ID,Title,Priority" --sort-by ID
python md-converter.py specs.xlsx --filter-word "*" --header-row 3 --row-numbers
python md-converter.py spec.docx --include-header-footer
python md-converter.py deck.pptx --no-notes
python md-converter.py manual.pdf
```

For spreadsheets the CLI looks for `ID`, `Description` and `Property` columns
(by name, any case) and keeps rows whose `Property` contains `gereksinim`;
change this with `--columns`, `--filter-column` and `--filter-word`. Default
output is `requirements.md` for spreadsheets and `<input>.md` for documents.
Run `python md-converter.py --help` for every option.

### Windows executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name md-converter-ui md-converter-ui.py
# with drag & drop:
pyinstaller --onefile --windowed --name md-converter-ui --collect-all tkinterdnd2 md-converter-ui.py
```

## Notes and limits

- PDF conversion is best effort and needs a text layer: scanned PDFs need OCR
  first. Larger text becomes headings, bullet characters become lists, ruled
  tables become tables. Lines repeated at the top/bottom of every page
  (running headers, page numbers) are dropped unless *Include headers &
  footers* is on.
- Images, text boxes, footnotes and comments are not converted.
- Merged table cells: a cell spanning several columns is written once, with
  the spanned columns left empty.

## Output example

```markdown
# Requirements

**Total requirements:** 2

| ID | Description |
|---|---|
| REQ-1 | The system shall ... |
| REQ-2 | The system shall ... |
```
