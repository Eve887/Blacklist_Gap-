"""
03_assam_master.py - STEP 2c (Assam): turn the Assam notices into the Assam
buyer file.

Reads:
  output/assam_index.csv   (made by 01_download.py)
  data/text/*.txt          (made by 02_extract_pdfs.py)
Writes:
  output/buyers/AS.xlsx  with one sheet:
    - Master list   : Assam rows, in the 19 master-list columns

This is the Assam half of what used to be 03_build_master.py, split out so
each buyer has its own script and its own output/buyers/<BUYER>.xlsx file -
10_merge_master.py combines them all into master_list.xlsx.

The notices are letters. The script reads each letter's text and looks for:
  - the "Sub:" (subject) line  -> used as "Reason (as written)"
  - key words                  -> first guess at the reason category
  - "debarred for 3 years"     -> period
"""

import csv
import re
from datetime import date

from openpyxl import Workbook

from common import OUT, TEXT, BUYERS, COLUMNS, standard_name, clean_date, one_line, \
    dedup_entry_ids, write_sheet, CATEGORY_PENDING

TODAY = date.today().strftime("%d-%m-%Y")


def assam_rows():
    index = OUT / "assam_index.csv"
    if not index.exists():
        print("assam_index.csv not found - run 01_download.py first.")
        return []
    rows = []
    with open(index, encoding="utf-8-sig") as f:
        for item in csv.DictReader(f):
            txt_path = TEXT / (item["file"].rsplit(".", 1)[0] + ".txt")
            text = txt_path.read_text(encoding="utf-8") if txt_path.exists() else ""
            flat = one_line(text)

            subject = ""
            m = re.search(r"Sub(?:ject)?\s*[:\-]\s*(.+?)(?:Ref\s*[:\-.]|Sir|Madam|$)", flat, re.I)
            if m:
                subject = m.group(1).strip()[:400]
            period = ""
            p = re.search(r"debarred\s+for\s+(\d+|one|two|three|five)\s*\(?\w*\)?\s*years?", flat, re.I)
            if p:
                period = p.group(0)
            level = "Firm" if re.search(r"firm\s+debarment", flat, re.I) else \
                    ("Product" if re.search(r"product\s+debarment", flat, re.I) else "To check")

            name = item["firm_as_written"]
            status = "Auto - check" if text else "No text - read the PDF yourself"
            # AS-<Sl no> - Assam's own list number from the source page, not a count.
            sl_no = item["sl_no"]
            entry_id = f"AS-{int(sl_no):03d}" if sl_no.isdigit() else f"AS-{sl_no}"
            rows.append([entry_id, name, standard_name(name), "See notice", "AMSCL", "Assam", level,
                         "See notice", clean_date(item["date"]), "See notice", period or "See notice",
                         subject or "See notice", CATEGORY_PENDING, "See notice", "See notice",
                         item["pdf_url"], f"Sl no {item['sl_no']} ({item['file']})", status, TODAY])
    return rows


def main():
    rows = assam_rows()
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [10, 30, 30, 34, 10, 14, 12, 26, 13, 13, 14, 34, 22, 18, 26, 40, 30, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "AS.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")


if __name__ == "__main__":
    main()
