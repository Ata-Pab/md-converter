#!/usr/bin/env python3
"""
Convert a spreadsheet (.xlsx/.xlsm/.xls/.ods) or a document
(.docx/.pptx/.odt/.pdf) file to Markdown.

Documents: Word, PowerPoint, OpenDocument and PDF files are converted whole
to Markdown (headings, paragraphs, lists, tables, links and, for slides,
notes) in document order. PDF is best effort: only text-based PDFs work (a
scanned PDF needs OCR and is not supported).

Spreadsheets: looks for "ID", "Description" and "Property" columns (by
name, case-insensitive, regardless of column order or how many other columns
the sheet has) and writes an ID | Description Markdown table of the rows
whose "Property" contains the filter word, plus the total count. Use
--columns / --filter-column to pick other columns.

Usage:
    python md-converter.py input.docx [output.md] [--include-header-footer]
    python md-converter.py input.pptx [output.md] [--no-notes]
    python md-converter.py input.odt  [output.md]
    python md-converter.py input.pdf  [output.md]
    python md-converter.py input.xlsx [output.md] [--sheet SHEET] [--filter-word WORD]

Examples:
    python md-converter.py spec.docx spec.md --include-header-footer
    python md-converter.py deck.pptx --no-notes
    python md-converter.py manual.pdf
    python md-converter.py specs.xlsx --sheet "Sheet1" --filter-word Requirement
    python md-converter.py specs.xlsx --columns "ID,Title,Priority" --sort-by ID
    python md-converter.py specs.xlsx --filter-word "*" --header-row 3

Requires: pandas, openpyxl, xlrd, python-docx, python-pptx, odfpy, pdfplumber
    pip install pandas openpyxl xlrd python-docx python-pptx odfpy pdfplumber
"""

import argparse
import sys
from pathlib import Path

from md_converter_core import (
    DEFAULT_FILTER_WORD,
    SUPPORTED_SUFFIXES,
    WHOLE_DOC_CONVERTERS,
    describe_error,
    document_to_markdown,
    find_columns,
    read_excel,
    table_to_markdown,
)

DEFAULT_COLUMNS = "ID,Description"
DEFAULT_FILTER_COLUMN = "Property"


def convert_document(input_path, output_path, args):
    md = document_to_markdown(
        input_path,
        title=args.title,
        include_header_footer=args.include_header_footer,
        include_notes=not args.no_notes,
        keep_formatting=not args.plain,
    )
    body = md.strip()
    if not body or (args.title and body == f"# {args.title}"):
        sys.exit(f"Error: no text found in {input_path.name}. "
                 "A scanned PDF has no text layer; it needs OCR first.")
    output_path.write_text(md, encoding="utf-8")
    print(f"Done. Document written to {output_path}")


def convert_excel(input_path, output_path, args):
    df = read_excel(input_path, args.sheet, args.header_row)

    wanted = [c.strip() for c in args.columns.split(",") if c.strip()]
    if not wanted:
        sys.exit("Error: --columns needs at least one column name")
    filter_word = (args.filter_word or "").strip()
    use_filter = bool(filter_word) and filter_word != "*"
    required = list(dict.fromkeys(
        wanted
        + ([args.filter_column] if use_filter else [])
        + ([args.sort_by] if args.sort_by else [])
    ))
    try:
        cols = find_columns(df, required)
    except ValueError as e:
        sys.exit(f"Error: {e}")

    # Output headers are the names as typed (e.g. "ID"), whatever the case
    # of the source columns.
    df = df.copy()
    for name in wanted:
        df[name] = df[cols[name]]
    rules = []
    if use_filter:
        rules.append({"column": cols[args.filter_column], "op": "word", "value": filter_word})

    title = "Requirements" if args.title is None else args.title
    md, total = table_to_markdown(
        df, wanted,
        rules=rules,
        sort_by=cols[args.sort_by] if args.sort_by else None,
        descending=args.descending,
        title=title,
        show_count=not args.no_count,
        count_label=args.count_label,
        keep_line_breaks=args.keep_line_breaks,
        row_numbers=args.row_numbers,
    )
    output_path.write_text(md, encoding="utf-8")
    print(f"Done. {total} row(s) written to {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("input", help="Path to the input spreadsheet or document")
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="Path to the output .md file "
        "(default: requirements.md for spreadsheets, <input>.md for documents)",
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Spreadsheets: table heading (default: Requirements; '' = none). "
        "Documents: adds this as a top-level heading.",
    )

    sheet = parser.add_argument_group("spreadsheet options")
    sheet.add_argument(
        "--sheet",
        default="0",
        help="Sheet name or 0-based index to read (default: first sheet)",
    )
    sheet.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="Row number (1-based) that holds the column names (default: 1)",
    )
    sheet.add_argument(
        "--columns",
        default=DEFAULT_COLUMNS,
        help=f"Comma-separated columns to extract, in order (default: {DEFAULT_COLUMNS})",
    )
    sheet.add_argument(
        "--filter-column",
        default=DEFAULT_FILTER_COLUMN,
        help=f"Column the filter word is matched in (default: {DEFAULT_FILTER_COLUMN})",
    )
    sheet.add_argument(
        "--filter-word",
        default=DEFAULT_FILTER_WORD,
        help="Keep rows whose filter column contains this word "
        f"(default: {DEFAULT_FILTER_WORD}; '*' = all rows)",
    )
    sheet.add_argument("--sort-by", default=None, help="Sort rows by this column (natural order)")
    sheet.add_argument("--descending", action="store_true", help="Sort Z to A")
    sheet.add_argument("--no-count", action="store_true", help="Leave out the total count line")
    sheet.add_argument(
        "--count-label",
        default="Total requirements",
        help="Label of the count line (default: 'Total requirements')",
    )
    sheet.add_argument(
        "--keep-line-breaks",
        action="store_true",
        help="Keep line breaks inside cells as <br> (default: join with spaces)",
    )
    sheet.add_argument("--row-numbers", action="store_true", help="Add a leading # column")

    doc = parser.add_argument_group("document options")
    doc.add_argument(
        "--include-header-footer",
        action="store_true",
        help="Include page headers and footers (Word, ODT, PowerPoint footers, "
        "repeated PDF page headers)",
    )
    doc.add_argument("--no-notes", action="store_true", help="Leave out PowerPoint speaker notes")
    doc.add_argument(
        "--plain",
        action="store_true",
        help="Plain text: no bold, italic or links (Word, PowerPoint)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"Error: input file not found: {input_path}")
    suffix = input_path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        sys.exit(f"Error: unsupported file type '{suffix}'. "
                 f"Supported: {', '.join(SUPPORTED_SUFFIXES)}")
    is_doc = suffix in WHOLE_DOC_CONVERTERS

    if args.output:
        output_path = Path(args.output)
        if not output_path.suffix:
            output_path = output_path.with_suffix(".md")
    elif is_doc:
        output_path = input_path.with_suffix(".md")
    else:
        output_path = Path("requirements.md")

    try:
        if is_doc:
            convert_document(input_path, output_path, args)
        else:
            convert_excel(input_path, output_path, args)
    except Exception as e:  # report every failure as a one-line message
        sys.exit(f"Error converting {input_path.name}: {describe_error(e)}")


if __name__ == "__main__":
    main()
