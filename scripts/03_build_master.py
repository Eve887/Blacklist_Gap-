"""
03_build_master.py - STEP 2c (Andhra Pradesh): turn the extracted AP rows
into the AP buyer file.

Reads:
  output/extracted_tables.xlsx   (made by 02_extract_pdfs.py)
Writes:
  output/buyers/AP.xlsx  with two sheets:
    - Master list   : AP rows only, in the 19 master-list columns
    - Skipped rows  : table rows the script could not understand - check these by hand

This script used to ALSO build Assam's rows and the final master_list.xlsx.
It now builds ONLY the AP buyer file - Assam has its own script
(03_assam_master.py), and 10_merge_master.py combines every file in
output/buyers/ (plus Matches and Unmatched revocations) into master_list.xlsx.

Every value is plain typed text. No formulas, no merged cells.
The "Status" column says "Auto - check" on every row: the script makes a
first draft, and YOU confirm each row against the PDF.
"""

import re
from datetime import date

from openpyxl import Workbook, load_workbook

from common import OUT, BUYERS, COLUMNS, standard_name, split_name_address, clean_date, \
    one_line, looks_like_date, dedup_entry_ids, write_sheet, CATEGORY_PENDING

AP_FIRM_URL = "https://apmsidc.ap.nic.in/ftp/drugs/Firm_Blacklisted_APMSIDC_UpdatedDt 13-12-2022.pdf"
AP_ITEM_URL = "https://apmsidc.ap.nic.in/ftp/drugs/Items_Blacklisted_APMSIDC_UpdatedDt 13-09-2024 (1).pdf"
TODAY = date.today().strftime("%d-%m-%Y")


def is_data_row(cells):
    """A real data row starts with a serial number such as 1, 2, 15."""
    return bool(cells) and re.fullmatch(r"\d{1,4}\.?", cells[0].strip() or "x") is not None


# ------------------------------------------------------------------
# ANDHRA PRADESH
# Firm list    = 7 columns : S.No | Firm+address | From | Up to | Period | Reason | Remarks
# Product list = 9 columns : S No | Item | Firm+address | NSQ declared by | From | Up to | Period | Reason | Status
# The script decides which list a row belongs to by the FILE NAME,
# then checks that the number of columns is what it expects.
# ------------------------------------------------------------------
def andhra_rows(skipped):
    path = OUT / "extracted_tables.xlsx"
    if not path.exists():
        print("extracted_tables.xlsx not found - run 02_extract_pdfs.py first.")
        return []
    ws = load_workbook(path, read_only=True).active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        folder, fname, page, tno, rno, year, table_columns = r[0], r[1], r[2], r[3], r[4], r[5], r[7]
        if folder != "AP_APMSIDC":
            continue
        all_cells = [str(c) if c is not None else "" for c in r[8:]]
        # table_columns = pdfplumber's REAL column count for this row (from 02_extract_pdfs.py).
        # Use it instead of guessing from the cells - this keeps a genuinely blank last
        # column (Remarks / Status) instead of losing it, which used to cause a wrong
        # cell count and send the whole row to Skipped rows for no real reason.
        if isinstance(table_columns, (int, float)) and table_columns > 0:
            cells = all_cells[:int(table_columns)]
        else:
            cells = all_cells
        if not is_data_row(cells):
            continue                                     # header rows like "S.No | Name of the firm"
        where = f"{fname} page {page}, table {tno}, row {rno}"
        firm_list = "firm" in fname.lower()

        if firm_list and len(cells) == 7:
            sno, firm_cell, frm, to, period, reason, remarks = cells
            level, product, lab, link = "Firm", "-", "Not stated", AP_FIRM_URL
            ref = f"S.No {sno.strip('.')}"
            # AP-F-<S.No> - the firm list has one simple serial number, so that IS
            # a stable ID: it comes straight from the source PDF, not from counting.
            entry_id = f"AP-F-{int(sno.strip('.')):03d}"
        elif not firm_list and len(cells) == 9:
            sno, product, firm_cell, lab, frm, to, period, reason, remarks = cells
            level, link = "Product", AP_ITEM_URL
            ref = f"Year {year or '?'}, S No {sno.strip('.')}"
            # AP-P-<year>-<S No> - the product list RESTARTS its S.No at 1 for every
            # "Year ...." section, so S.No alone is not unique - the year has to be
            # part of the ID too, or Year 2020-21 S No 1 would collide with Year
            # 2021-22 S No 1.
            year_clean = re.sub(r"\s", "", str(year)).strip() if year else "UNKNOWN"
            entry_id = f"AP-P-{year_clean}-{int(sno.strip('.')):03d}"
        else:
            skipped.append([where, len(cells)] + cells)
            continue

        if not (looks_like_date(frm) and looks_like_date(to)):
            # column count looked right, but From/Up to isn't a real date - don't trust it
            skipped.append([where, len(cells)] + cells)
            continue

        name, address = split_name_address(firm_cell)
        court = one_line(remarks) or "Not stated"
        rows.append([entry_id, one_line(name), standard_name(name), one_line(address), "APMSIDC",
                     "Andhra Pradesh", level, one_line(product), clean_date(frm), clean_date(to),
                     one_line(period), one_line(reason), CATEGORY_PENDING, one_line(lab),
                     court, link, f"{ref} ({where})", "Auto - check", TODAY])
    return rows


def main():
    skipped = []
    rows = andhra_rows(skipped)
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [10, 30, 30, 34, 10, 14, 12, 26, 13, 13, 14, 34, 22, 18, 26, 40, 30, 16, 12])
    wb.active.title = "Master list"
    width = max([len(s) for s in skipped], default=3)
    write_sheet(wb.create_sheet("Skipped rows"),
                ["Where", "Number of cells"] + [f"cell_{i}" for i in range(1, width - 1)],
                skipped, [40, 8] + [20] * (width - 2))
    out = BUYERS / "AP.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")
    print(f"  Skipped rows     : {len(skipped)}  <- open the 'Skipped rows' sheet and add these by hand")


if __name__ == "__main__":
    main()
