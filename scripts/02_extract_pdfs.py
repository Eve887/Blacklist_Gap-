"""
02_extract_pdfs.py - STEP 2b: copy the tables and text out of every PDF.

For every PDF inside data/raw/ it:
  1. Saves the full text of the PDF as a .txt file in data/text/
     (needed for the Assam notices, which are letters, not tables).
  2. Finds every table on every page and writes each table row into
     output/extracted_tables.xlsx - one Excel row per table row, with
     the file name and page number, so you can always go back and check.
  3. For the Andhra Pradesh product list, it also notes which
     "for the Year 20xx-xx" heading each table sits under.

If a page has no text at all, the PDF is probably a scanned image.
The script tells you; those pages need OCR (reading text from images).
"""

import re
import pdfplumber                     # reads text and tables from PDFs
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

from common import RAW, TEXT, OUT

OUT_FILE = OUT / "extracted_tables.xlsx"
YEAR_HEADING = re.compile(r"for\s+the\s+Year\s+(\d{4}\s*-\s*\d{2,4})", re.I)
MAX_CELLS = 12                        # widest table we expect (AP product list has 9-10 columns)


def year_headings(page):
    """Find every 'for the Year 2019-20' heading on a page, with how far down the page it is."""
    found = []
    for m in page.search(YEAR_HEADING.pattern, regex=True, case=False):
        y = YEAR_HEADING.search(m["text"])
        if y:
            found.append((m["top"], re.sub(r"\s", "", y.group(1))))
    return sorted(found)


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "tables"
    header = ["buyer_folder", "file", "page", "table_no", "row_no", "year_section", "cells_filled",
               "table_columns"] + [f"cell_{i}" for i in range(1, MAX_CELLS + 1)]
    ws.append(header)

    pdfs = sorted(RAW.rglob("*.pdf"))          # rglob = look in every sub-folder too
    if not pdfs:
        print("No PDFs found in", RAW, "- run 01_download.py first, or save PDFs there by hand.")
        return

    scanned = []
    for pdf_path in pdfs:
        folder = pdf_path.parent.name
        print(f"Reading {folder}/{pdf_path.name}")
        all_text = []
        current_year = ""                      # carried from page to page
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for pno, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text() or ""
                    all_text.append(f"===== PAGE {pno} =====\n{text}")
                    if not text.strip():
                        scanned.append(f"{pdf_path.name} page {pno}")

                    headings = year_headings(page)
                    tables = page.find_tables()
                    for tno, table in enumerate(tables, start=1):
                        # which "Year" heading sits above this table? (bbox[1] = top edge)
                        for top, yr in headings:
                            if top < table.bbox[1]:
                                current_year = yr
                        for rno, row in enumerate(table.extract(), start=1):
                            table_columns = len(row)         # real width pdfplumber found for THIS row
                            cells = [(c or "").strip() for c in row]
                            if not any(cells):
                                continue                     # skip fully empty rows
                            # Some PDFs stack two "Year" tables so closely that pdfplumber
                            # reads them as ONE table - the second year's heading then shows
                            # up as an ordinary ROW inside that table, not as free text above
                            # it, so the "above the table" check above never sees it. Catch
                            # that here: if THIS row's own text is a "for the Year ..."
                            # heading, use it to update current_year and do not save the
                            # row as data (it is a heading, not a firm/product entry).
                            mid_table_heading = YEAR_HEADING.search(" ".join(cells))
                            if mid_table_heading:
                                current_year = re.sub(r"\s", "", mid_table_heading.group(1))
                                continue
                            filled = sum(1 for c in cells if c)
                            ws.append([folder, pdf_path.name, pno, tno, rno, current_year, filled,
                                       table_columns] + cells[:MAX_CELLS])
                    # a heading BELOW the last table belongs to the table on the next page
                    if headings and (not tables or headings[-1][0] > tables[-1].bbox[1]):
                        current_year = headings[-1][1]
        except Exception as e:                 # a broken PDF should not stop the whole run
            print(f"   could not read this PDF: {e}")
            continue

        (TEXT / (pdf_path.stem + ".txt")).write_text("\n".join(all_text), encoding="utf-8")

    # make the sheet readable: Arial font, wrapped text, sensible widths
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Arial", size=10, bold=(c.row == 1))
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, width in zip("ABCDEFGH", [14, 34, 6, 8, 7, 12, 8, 8]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "I2"
    wb.save(OUT_FILE)

    print(f"\nSaved: {OUT_FILE}  ({ws.max_row - 1} table rows)")
    print(f"Text files saved in: {TEXT}")
    if scanned:
        print("\nWARNING - these pages have no text (probably scanned images):")
        for s in scanned:
            print("   ", s)


if __name__ == "__main__":
    main()
