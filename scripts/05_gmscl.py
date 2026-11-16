"""
05_gmscl.py - Gujarat Medical Services Corporation Limited (GMSCL) buyer file.

Fetches the "Current Blacklisted" table from:
  https://gmscl.gujarat.gov.in/black-list-firm.htm

This is an ASP.NET page with postback pagination (__doPostBack), NOT a
simple ?page=N URL. Confirmed by direct testing (not guessed) that the
pagination postback ('...gvDocument','Page$N') works correctly over plain
HTTP - no browser needed for that part.

The "Archive Blacklisted" dropdown on the same page does NOT work this way:
the postback registers the dropdown's new value (it comes back selected in
the response), but the data table itself never changes from "Current" no
matter what was tried (direct dropdown postback, combined with a simulated
search-button click, several guessed query-string variants). This looks
like it needs real browser JS to work. Logged to source_issues.csv. The
archive view will be captured separately with a one-off Playwright script
(06_gmscl_archive_playwright.py) and merged in afterwards - NOT built here.

IDs: GMSCL-DEBAR-<order number's serial>-<year>, e.g. GMSCL-DEBAR-18394-421-2025-26.
GMSCL gives a real order number/reference for every row (the "Debarred
Order" column), which is the most stable, traceable identifier available -
more useful than falling back to date+code. The full order text is kept in
"Page / row in source" regardless of how the short ID turned out.
Two rows CAN legitimately share one order number (one order covering two
different items, seen for real on this site) - the "-b"/"-c" rule handles
that; it is not a bug.

Writes:
  output/buyers/GMSCL.xlsx  with one sheet:
    - Master list   : GMSCL rows, in the 19 master-list columns
"""

import re
import time
import hashlib
from datetime import date
from urllib.parse import urljoin, urlparse
from urllib import robotparser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import OUT, BUYERS, COLUMNS, standard_name, clean_date, one_line, \
    dedup_entry_ids, write_sheet, looks_like_date, strip_leaked_address, CATEGORY_PENDING

BASE = "https://gmscl.gujarat.gov.in/"
PAGE = "black-list-firm.htm"
URL = urljoin(BASE, PAGE)

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")
MAX_PAGES = 40                           # safety cap - stop if something loops forever

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


def form_fields(soup):
    """Collect every input/select field EXCEPT submit/image/button controls -
    including one of those would tell the server a button was 'clicked',
    which is exactly what broke this the first time (triggered an unrelated
    search-validation error)."""
    data = {}
    for inp in soup.find_all("input"):
        name = inp.get("name")
        itype = (inp.get("type") or "text").lower()
        if not name or itype in ("submit", "image", "button", "checkbox", "radio"):
            continue
        data[name] = inp.get("value", "") or ""
    for sel in soup.find_all("select"):
        name = sel.get("name")
        if not name:
            continue
        opt = sel.find("option", selected=True) or sel.find("option")
        data[name] = opt.get("value", "") if opt else ""
    return data


def fetch_first_page():
    r = session.get(URL, timeout=TIMEOUT)
    r.raise_for_status()
    time.sleep(WAIT_SECONDS)
    return BeautifulSoup(r.text, "html.parser")


def fetch_page_n(soup, page_no):
    """Follow the confirmed-working GridView pager postback for page `page_no`."""
    data = form_fields(soup)
    data["__EVENTTARGET"] = "ctl00$ContentPlaceHolder1$FirmBlackList1$gvDocument"
    data["__EVENTARGUMENT"] = f"Page${page_no}"
    r = session.post(URL, data=data, timeout=TIMEOUT)
    r.raise_for_status()
    time.sleep(WAIT_SECONDS)
    return BeautifulSoup(r.text, "html.parser")


PAGE_LINK_RE = re.compile(r"gvDocument','Page\$(\d+)'")


def discover_page_numbers(soup):
    return {int(n) for n in PAGE_LINK_RE.findall(str(soup))}


def parse_rows(soup):
    gv = soup.find(id=lambda x: x and "gvDocument" in x)
    if gv is None:
        return []
    out = []
    for tr in gv.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        links = [urljoin(BASE, a["href"]) for a in tr.find_all("a", href=True)
                  if not a["href"].lower().startswith("javascript:")]
        if len(cells) < 9 or not cells[0].strip().isdigit():
            continue                                        # header rows
        sno, item, firm, frm, to, order, reason, remark, _attach = cells[:9]
        if not (looks_like_date(frm) and looks_like_date(to)):
            # the GridView's own page-number row (e.g. ['1','2','3',...,'10']) also starts
            # with a digit, so the isdigit() check alone lets it through - a real row
            # always has real dates in these two positions, the pager row never does.
            continue
        out.append({"sno": sno, "item": item, "firm": firm, "frm": frm, "to": to,
                     "order": order, "reason": reason, "remark": remark, "links": links})
    return out


# ------------------------------------------------------------------
# ENTRY ID - serial + year out of the order number, e.g.
# "GMSCL/Drug/Debar/18394-421/2025-26 Dt. 06/10/2026" -> "18394-421-2025-26"
# ------------------------------------------------------------------
SERIAL_RE = re.compile(r"\b(\d{3,6}-\d{2,6})\b")
YEAR_RANGE_RE = re.compile(r"\b(\d{4}-\d{2,4})\b")


def order_id_part(order_text):
    year_m = YEAR_RANGE_RE.search(order_text)
    serial = None
    for m in SERIAL_RE.finditer(order_text):
        if not year_m or m.group(0) != year_m.group(0):
            serial = m.group(1)
            break
    year = year_m.group(1) if year_m else None
    if serial and year:
        return f"{serial}-{year}"
    if serial:
        return serial
    if year:
        return year
    cleaned = re.sub(r"[^0-9A-Za-z-]", "", order_text)[:30]
    return cleaned or "UNKNOWNORDER"


STAY_RE = re.compile(r"\b(stay|revok)", re.I)


# GMSCL's live table never carries a lab name (no such column exists on
# the page) - "Lab that declared NSQ" is honestly "Not stated" for every
# row by default. Where the FULL order PDF was actually read and names a
# specific lab, that real finding is recorded here instead of staying
# "Not stated" - this directly feeds quality_level()'s Q1/Q2 split
# (common.py), since a named statutory lab is exactly what Q1 means.
LAB_OVERRIDES = {
    # Read 2026-10-31 (Meril case, PROGRESS_LOG.md) - the PDF itself
    # names the lab, even though the live page's own text never does.
    "GMSCL-DEBAR-18621-652-2025-26": "Food and Drug Laboratory, Vadodara (read from the full order "
                                     "PDF - not on the live page itself)",
}


def build_row(e, page_label="Current tab", list_type="DEBAR"):
    item = e["item"].strip()
    is_firm_wide = item.lower().startswith("all ")      # e.g. "All Drugs/Items", "All Capsule"
    level = ("Firm" if is_firm_wide else "Product") + " (inferred)"
    court = "Not stated"
    for text in (e["remark"], e["reason"]):
        if STAY_RE.search(text or ""):
            court = one_line(text)
            break
    source_link = e["links"][0] if e["links"] else URL
    entry_id = f"GMSCL-{list_type}-{order_id_part(e['order'])}"
    where = f"Order {e['order']}, S.No {e['sno']} ({page_label})"
    # GMSCL's own "Name of Firm" field occasionally has an address appended
    # directly (no separate address column exists to split from) - e.g.
    # "M/s Swiss Pharma Pvt Ltd,Ahmedabad". Trim that off before standardizing.
    firm = strip_leaked_address(e["firm"])
    lab = LAB_OVERRIDES.get(entry_id, "Not stated")
    return [entry_id, one_line(firm), standard_name(firm), "Not stated", "GMSCL",
            "Gujarat", level, one_line(item), clean_date(e["frm"]), clean_date(e["to"]),
            "Not stated", one_line(e["reason"]), CATEGORY_PENDING, lab,
            court, source_link, where, "Auto - check", TODAY]


def collect_current():
    if not robots_allows(URL):
        print(f"  BLOCKED by robots.txt - open in a browser: {URL}")
        return []
    soup = fetch_first_page()
    visited = {1}
    all_entries = list(parse_rows(soup))
    to_visit = sorted(discover_page_numbers(soup) - visited)
    while to_visit and len(visited) < MAX_PAGES:
        page_no = to_visit.pop(0)
        if page_no in visited:
            continue
        print(f"  fetching page {page_no} ...")
        soup = fetch_page_n(soup, page_no)
        visited.add(page_no)
        all_entries.extend(parse_rows(soup))
        new_pages = discover_page_numbers(soup) - visited
        to_visit.extend(sorted(p for p in new_pages if p not in to_visit))
    rows = [build_row(e) for e in all_entries]
    print(f"  pages visited: {sorted(visited)}  ({len(rows)} rows)")
    return rows


def main():
    rows = collect_current()
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [26, 34, 30, 14, 10, 10, 18, 30, 13, 13, 12, 34, 22, 14, 20, 40, 50, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "GMSCL.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")


if __name__ == "__main__":
    main()
