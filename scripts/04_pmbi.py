"""
04_pmbi.py - Jan Aushadhi / PMBI buyer file.

Fetches PMBI's own:
  - Debarred List            (debar.aspx)        - currently active
  - Blacklist                (Blacklist.aspx)     - currently active
  - Archive - Debarred List  (ArchDebarred.aspx)  - older, no longer active
  - Archive - Blacklist      (ArchBlack.aspx)     - older, no longer active
  - Revocation Order         (RevocationOrder.aspx)

The archive pages are included on purpose, even though their bans have
ended: Part C ("was a ban ACTIVE on a given date") needs the full history,
not just what is banned today.

PMBI does not number its lists with a real serial - every row's "Sr. No."
is a static "." placeholder (confirmed by reading the raw HTML; it is a
CSS/JS counter that never runs on a plain fetch, and there is no number
behind it even on the server side). So Entry IDs are built from the From
Date and the Drug Code instead (or "FIRM" when the action is against the
whole firm, not one product) - see pmbi_id() below.

"(inferred)" is added to the Level column everywhere, because PMBI's pages
never say what "Debarred" vs "Blacklisted" means - that split is this
script's own reading of the data pattern, not a stated PMBI definition.

Writes:
  output/buyers/PMBI.xlsx  with two sheets:
    - Master list           : PMBI rows, in the 19 master-list columns
    - Unmatched revocations : revocations that could not be matched to a
                               specific PMBI debarment/blacklist record
  output/pmbi_download_log.csv : SHA-256 fingerprint of every order-letter
                               document downloaded from the Blacklist page

Safe to run again: documents already downloaded are skipped.
"""

import csv
import re
import time
import hashlib
from datetime import date
from urllib.parse import urljoin, urlparse, unquote
from urllib import robotparser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import OUT, RAW, BUYERS, COLUMNS, DATE_RE, standard_name, clean_date, \
    one_line, dedup_entry_ids, write_sheet, CATEGORY_PENDING

BASE = "https://www.pmbi.co.in/"
DEBAR_PAGE = "debar.aspx"
BLACK_PAGE = "Blacklist.aspx"
ARCH_DEBAR_PAGE = "ArchDebarred.aspx"
ARCH_BLACK_PAGE = "ArchBlack.aspx"
REVOKE_PAGE = "RevocationOrder.aspx"

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TRIES = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")

session = requests.Session()
session.headers.update(HEADERS)
download_log = []

_robots_cache = {}


def robots_allows(url):
    """Ask the website's robots.txt whether automatic downloading is allowed."""
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if root not in _robots_cache:
        rp = robotparser.RobotFileParser()
        try:
            r = session.get(root + "/robots.txt", timeout=20)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            else:
                rp.parse([])          # no robots.txt = no rules = allowed
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


def safe_filename(text):
    text = unquote(text)
    keep = "".join(c if c.isalnum() or c in " -_." else "_" for c in text)
    return "_".join(keep.split())[:120]


def download_evidence(url, filename):
    """Download one order-letter document (PDF or image). Returns a status word."""
    target = RAW / "PMBI" / filename
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and target.stat().st_size > 0:
        status = "already had"
    elif not robots_allows(url):
        print(f"  BLOCKED by robots.txt -> download it in your browser and save it as:\n"
              f"      {target}\n      link: {url}")
        status = "blocked by robots.txt - download manually"
    else:
        status = "failed"
        for attempt in range(1, TRIES + 1):
            try:
                r = session.get(url, timeout=TIMEOUT)
                r.raise_for_status()
                target.write_bytes(r.content)
                time.sleep(WAIT_SECONDS)
                status = "downloaded"
                break
            except requests.RequestException as e:
                print(f"  attempt {attempt}/{TRIES} failed: {e}")
                time.sleep(WAIT_SECONDS * attempt)
    download_log.append({
        "date_time": TODAY,
        "status": status,
        "file": str(target) if target.exists() else str(target) + " (not saved)",
        "sha256": sha256(target) if target.exists() else "",
        "url": url,
    })
    return status


def fetch_page(page):
    url = urljoin(BASE, page)
    if not robots_allows(url):
        print(f"  BLOCKED by robots.txt - open in a browser: {url}")
        return None
    r = session.get(url, timeout=TIMEOUT)
    r.raise_for_status()
    time.sleep(WAIT_SECONDS)
    return BeautifulSoup(r.text, "html.parser")


# ------------------------------------------------------------------
# PARSE ONE DEBARRED / BLACKLIST / ARCHIVE TABLE
#
# These tables all share one shape: Sr No | Drug Code | Drug Name |
# Manufactured by | Reason | From Date | To Date (the archive tables add
# one more, unused "Download/View" text column after that - harmless,
# never read).
#
# Two quirks the parser has to handle, found by reading the real pages,
# not guessed:
#   1. When ONE action covers several products, only ONE row of the group
#      carries the firm/reason/dates - the other rows repeat just the Drug
#      Code/Name with those cells blank. The blanks are filled in from the
#      most recent row that DID have them (forward-fill).
#   2. Some rows are not product rows at all - they are an EXTRA order-
#      letter link for the row just above (seen on the archive page: three
#      separate PDF links for one Syncom Healthcare entry, each its own
#      near-empty row). Those are folded into the previous row's own list
#      of links instead of becoming their own entry.
#
# "Entire Firm", "-" and "NA" are three different placeholders PMBI itself
# uses, in different places, for "no specific product" - all three are
# treated the same way here.
# ------------------------------------------------------------------
FIRM_PLACEHOLDERS = {"", "-", "na", "entire firm"}


def firm_level(code):
    return code.strip().lower() in FIRM_PLACEHOLDERS


def parse_table(soup):
    table = soup.find("table")
    if table is None:
        return []
    entries = []
    last_firm = last_reason = last_frm = last_to = ""
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        links = [urljoin(BASE, a["href"]) for a in tr.find_all("a", href=True)]
        if not cells:
            continue
        first = cells[0].strip().lower()
        if first.startswith("sr") or first.startswith("from"):
            continue                                       # header rows
        if cells == [""]:
            # a pure extra-document-link row - belongs to the entry just above
            if links and entries:
                entries[-1]["links"].extend(links)
            continue
        # Blacklist.aspx collapses Drug Code + Drug Name into ONE cell for
        # "Entire Firm" rows (confirmed by counting cells: 6, not the normal
        # 7), so Firm/Reason/From/To all shift one position earlier. Other
        # pages (debar.aspx, the archive pages) keep Code and Name as two
        # separate cells even for a firm-level row ("-"/"NA" each on their
        # own), so they keep the normal 7-cell shape.
        if len(cells) == 6 and len(cells) > 1 and firm_level(cells[1]):
            code = cells[1].strip()
            name = ""
            firm = cells[2].strip()
            reason = cells[3].strip()
            frm = cells[4].strip()
            to = cells[5].strip()
        else:
            code = cells[1].strip() if len(cells) > 1 else ""
            name = cells[2].strip() if len(cells) > 2 else ""
            firm = cells[3].strip() if len(cells) > 3 else ""
            reason = cells[4].strip() if len(cells) > 4 else ""
            frm = cells[5].strip() if len(cells) > 5 else ""
            to = cells[6].strip() if len(cells) > 6 else ""
        if firm:
            last_firm, last_reason, last_frm, last_to = firm, reason, frm, to
        else:
            firm, reason, frm, to = last_firm, last_reason, last_frm, last_to
        if not firm:
            continue                                       # nothing usable at all - skip
        entries.append({"code": code, "name": name, "firm": firm, "reason": reason,
                         "frm": frm, "to": to, "links": links})
    return entries


def ymd(text):
    """dd.mm.yyyy / dd-mm-yyyy -> yyyymmdd, for building a date-based Entry ID."""
    m = DATE_RE.search(str(text or ""))
    if not m:
        return "UNKNOWNDATE"
    d, mth, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    return f"{y}{int(mth):02d}{int(d):02d}"


def pmbi_id(prefix, frm, code):
    code_part = "FIRM" if firm_level(code) else (re.sub(r"[^0-9A-Za-z]", "", code) or "UNKNOWNCODE")
    return f"PMBI-{prefix}-{ymd(frm)}-{code_part}"


def build_row(entry, prefix, page_url, page_label):
    code, name, firm = entry["code"], entry["name"], entry["firm"]
    is_firm = firm_level(code)
    level = ("Firm" if is_firm else "Product") + " (inferred)"
    product = "-" if is_firm else one_line(name)
    code_desc = "Entire Firm" if is_firm else f"Drug Code {code}"
    where = f"{code_desc}, {firm}, From {entry['frm']} ({page_label})"
    source_link = entry["links"][0] if entry["links"] else page_url
    entry_id = pmbi_id(prefix, entry["frm"], code)
    return [entry_id, one_line(firm), standard_name(firm), "Not stated", "PMBI", "National",
            level, product, clean_date(entry["frm"]), clean_date(entry["to"]), "Not stated",
            one_line(entry["reason"]), CATEGORY_PENDING, "Not stated",
            "Not stated", source_link, where, "Auto - check", TODAY]


def download_evidence_links(entries):
    for e in entries:
        for link in e["links"]:
            fname = safe_filename(link.rsplit("/", 1)[-1])
            download_evidence(link, fname)


def collect_current():
    rows = []
    for prefix, page in (("DEBAR", DEBAR_PAGE), ("BLACK", BLACK_PAGE)):
        soup = fetch_page(page)
        if soup is None:
            continue
        entries = parse_table(soup)
        if prefix == "BLACK":
            download_evidence_links(entries)                # order letters linked from Blacklist.aspx
        page_url = urljoin(BASE, page)
        for e in entries:
            rows.append(build_row(e, prefix, page_url, page))
    return rows


def collect_archive():
    rows = []
    for page in (ARCH_DEBAR_PAGE, ARCH_BLACK_PAGE):
        soup = fetch_page(page)
        if soup is None:
            continue
        entries = parse_table(soup)
        page_url = urljoin(BASE, page)
        for e in entries:
            rows.append(build_row(e, "ARCH", page_url, page))
    return rows


def collect_revocations(all_rows):
    """Match each revocation to an existing row by standardised firm name.
    A match updates that row's 'Court stay or revoked?' field. A revocation
    that cannot be matched to anything is returned separately - it is kept,
    not dropped."""
    soup = fetch_page(REVOKE_PAGE)
    unmatched = []
    if soup is None:
        return unmatched
    table = soup.find("table")
    if table is None:
        return unmatched
    page_url = urljoin(BASE, REVOKE_PAGE)
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if not cells:
            continue
        first = cells[0].strip().lower()
        if first.startswith("sr") or first.startswith("from"):
            continue
        if len(cells) < 4 or not cells[3].strip():
            continue
        firm = cells[3].strip()
        reason = cells[4].strip() if len(cells) > 4 else ""
        date_cell = cells[5].strip() if len(cells) > 5 else ""
        target = standard_name(firm)
        matches = [r for r in all_rows if r[2] == target]
        if matches:
            note = f"Revoked - {date_cell}"
            for r in matches:
                r[14] = note
        else:
            unmatched.append(["PMBI", one_line(firm), standard_name(firm), one_line(reason),
                               clean_date(date_cell), page_url,
                               "Could not match to any current or archived PMBI Debarred/Blacklist "
                               "record (checked current Debarred List + Blacklist + both archive "
                               "pages, 59 rows total). See output/source_issues.csv."])
    return unmatched


def main():
    rows = collect_current()
    rows += collect_archive()
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    unmatched = collect_revocations(rows)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [16, 34, 30, 14, 10, 10, 16, 30, 13, 13, 12, 34, 22, 14, 20, 40, 50, 16, 12])
    wb.active.title = "Master list"
    write_sheet(wb.create_sheet("Unmatched revocations"),
                ["Buyer", "Firm name (as written)", "Firm name (standard)", "Reason (as written)",
                 "Revocation date", "Source link", "Note"],
                unmatched, [10, 34, 30, 34, 16, 40, 60])
    out = BUYERS / "PMBI.xlsx"
    wb.save(out)

    if download_log:
        log_path = OUT / "pmbi_download_log.csv"
        with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
            w.writeheader()
            w.writerows(download_log)
        print(f"Download log saved: {log_path}")

    print(f"Saved: {out}")
    print(f"  Master list rows      : {len(rows)}")
    print(f"  Unmatched revocations : {len(unmatched)}")


if __name__ == "__main__":
    main()
