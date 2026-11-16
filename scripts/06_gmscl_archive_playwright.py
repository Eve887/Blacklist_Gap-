"""
06_gmscl_archive_playwright.py - GMSCL's "Archive Blacklisted" view.

UPDATE: both automated attempts at this view are now permanently blocked
(see output/source_issues.csv) - the second attempt returned an explicit
CAPTCHA challenge, a deliberate access control this project will not try
to bypass. So this script now PREFERS a hand-saved copy of the page:

  If data/raw/GMSCL/archive_saved_by_hand/ contains one or more .html
  files, it parses those directly (same parse_rows()/build_row() code as
  05_gmscl.py - "same code as Current", literally, via import) and skips
  Playwright entirely - no point re-attempting a capture that is known to
  be blocked on every run.

  If that folder is empty, it prints exactly what to save and where, and
  stops - it no longer attempts the Playwright capture automatically.

HOW TO SAVE THE PAGE BY HAND (once per update you want captured):
  1. Open this page in your own browser: https://gmscl.gujarat.gov.in/black-list-firm.htm
  2. Find the "Current Blacklisted / Archive Blacklisted" dropdown near the
     top of the table and select "Archive Blacklisted".
  3. Wait for the table to refresh to the archive rows.
  4. Save the page (Ctrl+S or your browser's "Save Page As") as HTML into:
       data\\raw\\GMSCL\\archive_saved_by_hand\\
     Any filename is fine (e.g. archive_page1.html). If the archive list
     has its OWN pagination (page 2, 3, ...), repeat steps 2-4 for each
     page and save each one as a separate file in the same folder - this
     script reads every .html file there and combines them.

Original scope note (kept for history - the automated path below is no
longer used by default, only if you ever want to retry it by hand):
  - ONE browser, ONE page, headless.
  - Certificate verification is NOT disabled (no ignore_https_errors).
  - 3-second pause between page interactions (dropdown select, each
    pagination click), same politeness as the rest of this project.
  - No other GMSCL pages are touched.

Writes:
  data/raw/GMSCL/archive_blacklisted_pN.html   - the raw captured HTML,
                                                   one file per archive page
                                                   (kept as evidence)
  output/buyers/GMSCL_archive.xlsx             - a SEPARATE buyer file from
                                                   GMSCL.xlsx on purpose: if
                                                   this merged into
                                                   GMSCL.xlsx directly, a
                                                   later solo re-run of
                                                   05_gmscl.py would silently
                                                   wipe these rows out
                                                   (05_gmscl.py rebuilds
                                                   that file from scratch
                                                   every time, which is the
                                                   whole point of the buyer-
                                                   script architecture).
                                                   10_merge_master.py reads
                                                   every file in
                                                   output/buyers/ anyway, so
                                                   this is picked up
                                                   automatically. Rows
                                                   inside still say
                                                   Buyer = "GMSCL".

Parsing reuses the ACTUAL functions from 05_gmscl.py (imported, not
copy-pasted) so "same code as Current" is literally true.
"""

import importlib.util
import time
from pathlib import Path
from datetime import date

from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import RAW, BUYERS, COLUMNS, dedup_entry_ids, write_sheet

WAIT_SECONDS = 3
TODAY = date.today().strftime("%d-%m-%Y")
MAX_ARCHIVE_PAGES = 40

# import 05_gmscl.py by file path - its filename starts with a digit, so it
# cannot be imported with a normal "import 05_gmscl" statement.
_spec = importlib.util.spec_from_file_location("gmscl_current", Path(__file__).parent / "05_gmscl.py")
gmscl_current = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gmscl_current)

EVIDENCE_DIR = RAW / "GMSCL"
EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
HAND_SAVED_DIR = EVIDENCE_DIR / "archive_saved_by_hand"
HAND_SAVED_DIR.mkdir(parents=True, exist_ok=True)


def read_hand_saved_pages():
    """Returns the HTML text of every .html file in archive_saved_by_hand/,
    or an empty list if none exist yet."""
    files = sorted(HAND_SAVED_DIR.glob("*.html"))
    return [f.read_text(encoding="utf-8", errors="replace") for f in files], files


def capture_archive_pages():
    """Last-resort automated capture - NOT called by default any more (see
    module docstring); kept available in case you ever want to retry it by
    hand. Returns a list of raw HTML strings, one per archive page, saved
    to disk as we go."""
    from playwright.sync_api import sync_playwright         # imported here, not at module level,
                                                              # since this path is no longer used by default
    html_pages = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(gmscl_current.URL, wait_until="networkidle", timeout=60000)
        time.sleep(WAIT_SECONDS)

        page.select_option("#ContentPlaceHolder1_FirmBlackList1_ddlBlackList", "Archive_Blacklisted")
        page.wait_for_load_state("networkidle", timeout=60000)
        time.sleep(WAIT_SECONDS)

        page_no = 1
        while True:
            html = page.content()
            path = EVIDENCE_DIR / f"archive_blacklisted_p{page_no}.html"
            path.write_text(html, encoding="utf-8")
            html_pages.append(html)
            print(f"  captured page {page_no} -> {path}")

            next_link = page.locator(
                f"a[href*=\"gvDocument','Page${page_no + 1}'\"]"
            )
            if page_no >= MAX_ARCHIVE_PAGES or next_link.count() == 0:
                break
            next_link.first.click()
            page.wait_for_load_state("networkidle", timeout=60000)
            time.sleep(WAIT_SECONDS)
            page_no += 1

        browser.close()
    return html_pages


def main():
    html_pages, files = read_hand_saved_pages()
    if html_pages:
        print(f"Found {len(files)} hand-saved page(s) in {HAND_SAVED_DIR}:")
        for f in files:
            print(f"  {f.name}")
    else:
        print(f"No hand-saved pages found in {HAND_SAVED_DIR}\n")
        print("To capture the Archive Blacklisted view, save it there by hand:")
        print("  1. Open https://gmscl.gujarat.gov.in/black-list-firm.htm in your own browser.")
        print("  2. Select 'Archive Blacklisted' from the dropdown near the top of the table.")
        print("  3. Wait for the table to refresh to the archive rows.")
        print(f"  4. Save the page as HTML into: {HAND_SAVED_DIR}")
        print("     (one file per archive page, if the archive has its own pagination)")
        print("\nNot attempting the automated Playwright capture - it is permanently blocked "
              "(CAPTCHA), see output/source_issues.csv.")
        return

    all_entries = []
    for html in html_pages:
        soup = BeautifulSoup(html, "html.parser")
        all_entries.extend(gmscl_current.parse_rows(soup))

    rows = [gmscl_current.build_row(e, page_label="Archive tab", list_type="ARCH") for e in all_entries]
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [26, 34, 30, 14, 10, 10, 18, 30, 13, 13, 12, 34, 22, 14, 20, 40, 50, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "GMSCL_archive.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")

    nestor = [r for r in rows if "nestor" in r[1].lower()]
    if nestor:
        print("\nFound Nestor Pharmaceuticals in the archive:")
        for r in nestor:
            print(f"  {r[0]} | from {r[8]} to {r[9]} | product: {r[7]} | reason: {r[11]}")
    else:
        print("\nNestor Pharmaceuticals NOT found in the captured archive pages.")


if __name__ == "__main__":
    main()
