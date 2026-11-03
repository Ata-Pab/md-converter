# Example files

One sample file per supported type, all about the same made-up project (a
library management system). Open them in the GUI (*Add folder…* → this
folder) or run the commands below from the project folder.

| File | Type | What it shows |
|---|---|---|
| `requirements.xlsx` | Excel | The default requirements filter: rows whose `Property` contains *gereksinim* (also `GEREKSİNİM` and `Güvenlik gereksinimi`); *Bilgi*, *Not* and *Başlık* rows are left out. Has an empty row, pipes, `<admin>` and `*` in the text. Second sheet **Revisions** has its column names on **row 4**. |
| `test-cases.xlsm` | Excel (macro-enabled) | Multi-line *Steps* cells (try *Line breaks as `<br>`*) and a *Result* column to filter on (`equals` → `Fail`). |
| `legacy-requirements.xls` | Excel 97-2003 | The old format with `ID`, `Description`, `Property` columns. |
| `risk-register.ods` | LibreOffice Calc | A numeric *Score* column: sort it Z → A, or filter *Status* with a regex such as `open\|mitigated`. |
| `design-spec.docx` | Word | Title, headings, bold/italic, a link, nested bullet and numbered lists, a table with a merged cell, page header and footer. |
| `project-kickoff.pptx` | PowerPoint | Five slides: nested bullets, a table, grouped shapes, a link, speaker notes and a footer on every slide. |
| `meeting-notes.odt` | LibreOffice Writer | Headings, numbered and nested bullet lists, a table with a header row, a section, page header and footer. |
| `user-manual.pdf` | PDF | Three pages: headings (by font size), bullets, a ruled table, and a running header and page number that are dropped unless *Include headers & footers* is on. |

## Try it from the command line

```bash
python md-converter.py example/requirements.xlsx
python md-converter.py example/requirements.xlsx revisions.md --sheet Revisions --header-row 4 --columns "Version,Date,Change" --filter-word "*"
python md-converter.py example/test-cases.xlsm failed.md --columns "Test ID,Title,Steps,Result" --filter-column Result --filter-word fail --keep-line-breaks
python md-converter.py example/risk-register.ods risks.md --columns "ID,Risk,Score,Status" --filter-word "*" --sort-by Score --descending
python md-converter.py example/design-spec.docx --include-header-footer
python md-converter.py example/project-kickoff.pptx --no-notes
python md-converter.py example/user-manual.pdf
```

Documents are written next to the source (`example/design-spec.md`, …);
spreadsheets default to `requirements.md` in the current folder.
