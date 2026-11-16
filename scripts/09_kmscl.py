"""
09_kmscl.py - Kerala Medical Services Corporation Limited (KMSCL) buyer file.

Downloads and extracts KMSCL's two blacklist PDFs - there is no HTML table
or listing page, just two standalone PDFs linked from the homepage:
  - QCblacklisting.pdf     - "Quality Control Division" - items blacklisted
                              for Not of Standard Quality (NSQ). Every row on
                              this document IS an NSQ case by definition (it
                              is the document's own title), so Reason
                              category is set to "Quality" directly rather
                              than guessed from keywords.
  - blacklistednov22.pdf   - "Essential Drugs Procurement Division" - firms
                              blacklisted for contractual/administrative
                              reasons (e.g. "Non-submission of agreement").
                              This one DOES give a real "w.e.f" (with effect
                              from) date - the QC file does not.

Both PDFs use plain http:// links on the KMSCL homepage, which time out -
https:// on the same path works fine (confirmed by testing).

Both are small (3 and 2 rows respectively, confirmed by reading the actual
PDF content, not estimated from file size) - a full pipeline stage would be
overkill; this script just downloads, extracts, and builds rows directly.

IDs: KMSCL-<NSQ/ED>-<sanitized order number>, e.g. KMSCL-NSQ-KMSCL-QC-21492-2024.
Both documents give a real order number for every row.

Writes:
  output/buyers/KMSCL.xlsx  with one sheet:
    - Master list   : KMSCL rows, in the 19 master-list columns
"""

import re
import time
import hashlib
from datetime import date
from urllib.parse import urlparse
from urllib import robotparser

import requests
import pdfplumber
from openpyxl import Workbook

from common import RAW, BUYERS, COLUMNS, standard_name, split_name_address, clean_date, \
    one_line, dedup_entry_ids, write_sheet, CATEGORY_PENDING

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TRIES = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")

FILES = [
    ("QCblacklisting.pdf", "https://kmscl.kerala.gov.in/KMSCL/uploads/blacklisting/QCblacklisting.pdf", "NSQ"),
    ("blacklistednov22.pdf", "https://kmscl.kerala.gov.in/KMSCL/uploads/blacklisting/blacklistednov22.pdf", "ED"),
]

session = requests.Session()
session.headers.update(HEADERS)

_robots_cache = {}


def robots_allows(url):
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if root not in _robots_cache:
        rp = robotparser.RobotFileParser()
        try:
            r = session.get(root + "/robots.txt", timeout=20)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            else:
                rp.parse([])
        except requests.RequestException:
            rp.parse([])
        _robots_cache[root] = rp
    return _robots_cache[root].can_fetch(HEADERS["User-Agent"], url)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def download(url, filename):
    target = RAW / "KMSCL" / filename
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and target.stat().st_size > 0:
        return target, "already had"
    if not robots_allows(url):
        print(f"  BLOCKED by robots.txt -> download it in your browser and save it as:\n"
              f"      {target}\n      link: {url}")
        return target, "blocked"
    for attempt in range(1, TRIES + 1):
        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            if not r.content.startswith(b"%PDF"):
                return target, "not a pdf"
            target.write_bytes(r.content)
            time.sleep(WAIT_SECONDS)
            return target, "downloaded"
        except requests.RequestException as e:
            print(f"  attempt {attempt}/{TRIES} failed: {e}")
            time.sleep(WAIT_SECONDS * attempt)
    return target, "failed"


# Order text often has a SECOND, later "dated DD.MM.YYYY" confirmation/review
# date tacked on after the real order reference (e.g. "KMSCL/QC/21492/2024
# dated 31.07.2025") - use the FIRST number/year found (part of the order's
# own reference), not whatever appears last.
NUM_RE = re.compile(r"\d{3,8}")
YEAR_RE = re.compile(r"\b(20\d{2})\b")
FULL_YEAR_RE = re.compile(r"\A20\d{2}\Z")


def order_id_part(text):
    text = text or ""
    num = None
    for n in NUM_RE.findall(text):
        if FULL_YEAR_RE.match(n):
            continue
        num = n
        break
    years = YEAR_RE.findall(text)
    year = years[0] if years else None
    if num and year:
        return f"{num}-{year}"
    if num:
        return num
    if year:
        return year
    cleaned = re.sub(r"[^0-9A-Za-z]+", "-", text).strip("-")
    return cleaned[:50] or "UNKNOWNORDER"


def is_data_row(cells):
    return bool(cells) and re.fullmatch(r"\d{1,3}\.?", (cells[0] or "").strip() or "x") is not None


def nsq_rows(path):
    """QCblacklisting.pdf - SL No | Product | Company+Address | Tender Year | Valid upto | Order No"""
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.find_tables():
                for cells in table.extract():
                    cells = [(c or "").strip() for c in cells]
                    if not is_data_row(cells) or len(cells) < 6:
                        continue
                    sno, product, firm_cell, year, upto, order = cells[0], cells[1], cells[2], cells[3], cells[4], cells[5]
                    name, address = split_name_address(firm_cell)
                    entry_id = f"KMSCL-NSQ-{order_id_part(order)}"
                    where = f"SL {sno}, Order {one_line(order)} (QCblacklisting.pdf)"
                    rows.append([entry_id, one_line(name), standard_name(name), one_line(address),
                                 "KMSCL", "Kerala", "Product (inferred)", one_line(product),
                                 "Not stated", clean_date(upto), f"Tender year {year}" if year else "Not stated",
                                 "Not of Standard Quality (NSQ) - Quality Control Division", "Quality",
                                 "Not stated", "Not stated",
                                 "https://kmscl.kerala.gov.in/KMSCL/uploads/blacklisting/QCblacklisting.pdf",
                                 where, "Auto - check", TODAY])
    return rows


def ed_rows(path):
    """blacklistednov22.pdf - Sl No | Name & Address | w.e.f | (merged) | Order No & Date | (merged) |
    valid upto | (merged) | Reason"""
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.find_tables():
                for cells in table.extract():
                    cells = [(c or "").strip() if c else "" for c in cells]
                    if not is_data_row(cells) or len(cells) < 10:
                        continue
                    sno, firm_cell, wef = cells[0], cells[1], cells[2]
                    order, upto, reason = cells[5], cells[6], cells[9]
                    name, address = split_name_address(firm_cell)
                    entry_id = f"KMSCL-ED-{order_id_part(order)}"
                    where = f"SL {sno}, Order {one_line(order)} (blacklistednov22.pdf)"
                    rows.append([entry_id, one_line(name), standard_name(name), one_line(address),
                                 "KMSCL", "Kerala", "Firm (inferred)", "-", clean_date(wef), clean_date(upto),
                                 "Not stated", one_line(reason), CATEGORY_PENDING, "Not stated",
                                 "Not stated",
                                 "https://kmscl.kerala.gov.in/KMSCL/uploads/blacklisting/blacklistednov22.pdf",
                                 where, "Auto - check", TODAY])
    return rows


def main():
    rows = []
    for filename, url, list_type in FILES:
        path, status = download(url, filename)
        print(f"  {filename}: {status}")
        if status not in ("downloaded", "already had"):
            continue
        if list_type == "NSQ":
            rows.extend(nsq_rows(path))
        else:
            rows.extend(ed_rows(path))

    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [26, 34, 30, 30, 10, 10, 16, 26, 13, 13, 16, 34, 22, 14, 16, 50, 40, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "KMSCL.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")


if __name__ == "__main__":
    main()
