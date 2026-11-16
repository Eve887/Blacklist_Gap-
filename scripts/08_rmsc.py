"""
08_rmsc.py - Rajasthan Medical Services Corporation Ltd (RMSC) buyer file.

Fetches:
  https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/library/Banned-firms0/Banned_Product.html

Ordinary server-rendered HTML (AEM site), one table, no pagination needed
(confirmed: all rows are in the single GET response).

Column shape, confirmed by reading the real page, not guessed:
  S.NO | Company Name | Product Name | Reason | Up To | Download
S.NO is blank on every row. There is NO "From" date column at all on this
page - only "Up To" - so "Blacklisted from" is honestly "Not stated" here,
not fabricated. The "Download" column's text is actually an order
reference (e.g. "LB13/328/29.02.2024"), not a URL - the real download link
is a separate <a> on the same cell. "Product Name" sometimes lists SEVERAL
products in one cell (e.g. Ridley Life Sciences: "NRD-71 (...), NRD-477
(...), ..."), which means the SOURCE itself treats that as one banning
action covering several products - kept as one row with the full text, not
split, since splitting would invent a per-product action that was not
actually recorded separately.

"Up To" sometimes also carries a "(For N years)" phrase - extracted as the
Period, as written, when present.

IDs: RMSC-DEBAR-<sanitized order reference>, e.g. RMSC-DEBAR-LB13-328-29-02-2024.
RMSC gives an order-like reference for every row, so that is the ID basis,
per the project's "order number if available" rule.

Writes:
  output/buyers/RMSC.xlsx  with one sheet:
    - Master list   : RMSC rows, in the 19 master-list columns
"""

import re
import time
from datetime import date
from urllib.parse import urljoin, urlparse
from urllib import robotparser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import BUYERS, COLUMNS, standard_name, clean_date, one_line, \
    dedup_entry_ids, write_sheet, looks_like_date, CATEGORY_PENDING

BASE = "https://rmsc.health.rajasthan.gov.in/"
PAGE_URL = ("https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-"
            "corporation-ltd-/en/library/Banned-firms0/Banned_Product.html")

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")

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


PERIOD_RE = re.compile(r"\(for\s+[^)]+\)", re.I)
STAY_RE = re.compile(r"\b(stay|revok)", re.I)


def order_id_part(text):
    cleaned = re.sub(r"[^0-9A-Za-z]+", "-", text or "").strip("-")
    return cleaned[:60] or "UNKNOWNORDER"


def parse_rows(soup):
    table = soup.find("table")
    if table is None:
        return []
    out = []
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        links = [urljoin(BASE, a["href"]) for a in tr.find_all("a", href=True)]
        if len(cells) < 6:
            continue
        sno, firm, product, reason, upto, order = cells[:6]
        if not firm or firm.strip().upper() == "COMPANY NAME":
            continue                                        # header row
        out.append({"firm": firm, "product": product, "reason": reason, "upto": upto,
                     "order": order, "links": links})
    return out


# A targeted, hand-verified Entry-ID override used to live here for one
# case: reason_category() auto-categorizes from the ENGLISH text on this
# live page alone ("Violation of conditions" -> "To check", since it
# matches neither QUALITY_WORDS nor SUPPLY_WORDS), but reading the actual
# order letter (Hindi, OCR'd - see PROGRESS_LOG.md Entry 13-14) revealed
# what that vague English summary doesn't say: RMSC-DEBAR-1837-12-08-2025
# is a BID-RATE CORRECTION dispute (the firm tried to fix a mistaken BOQ
# rate AFTER the financial bid opened, ruled a tender-condition
# violation), not a drug-quality finding at all. Moved to
# 10_merge_master.py 2026-11-01 (PROGRESS_LOG.md Entry 19, checklist item
# 13) along with MPPHSCL's overrides, so every override and the generic
# keyword classifier run from the same central place on every merge - see
# 10_merge_master.py for the override itself.

# RMSC's own live page never names a lab (no such column exists on the
# page) - "Lab that declared NSQ" is honestly "Not stated" for every row
# by default, same as GMSCL (see 05_gmscl.py LAB_OVERRIDES for the same
# pattern). Where the actual Hindi order letter was read (hin+eng OCR,
# then hand-translated - see build_pdfs_to_read.py RMSC_HINDI_READINGS
# and PROGRESS_LOG.md) and names a real lab, that hand-verified fact is
# recorded here instead of staying "Not stated" - this directly feeds
# quality_level()'s Q1/Q2 split (common.py), since a named statutory lab
# is exactly what Q1 means. Added 2026-11-01 per the user's own document
# check (document_checks_2026-11-01.xlsx, "Hindi letter
# corrections" sheet) - these 6 rows were wrongly Q2 because this field
# was never populated, not because the letters don't name a lab.
RMSC_LAB_OVERRIDES = {
    # Zee Laboratories - Itraconazole Cap 100mg[118], order letter 1846
    # dt 16-11-2025 (10-11-2025 is the RMSC office-order date on the live
    # page; the letter itself is numbered 1846 dt 16-11-2025 - see
    # source_issues.csv for this date discrepancy, not resolved here).
    "RMSC-DEBAR-1846-10-11-2025-2472-16-01-2026":
        "State Drug Testing Laboratory (dissolution failure); objection sample confirmed by "
        "Central Drugs Laboratory, Kolkata (Not of Standard Quality, Grossly Substandard) - "
        "read from the Hindi order letter, hin+eng OCR, hand-translated.",
    # ANG Life Sciences - Dobutamine Inj IP 50mg/ml/250mg vial[192]. NOTE:
    # the order letter itself is numbered 3099 dt 21-02-2025; this
    # project's own Entry ID uses "3078", the reference shown on RMSC's
    # live page - see source_issues.csv for this unresolved discrepancy.
    "RMSC-DEBAR-3078-20-02-2025-b":
        "State Drug Testing Laboratory (Description failure, confirmed on retest) - read from "
        "the Hindi order letter (numbered 3099 in the letter itself), hin+eng OCR, "
        "hand-translated.",
    # Vee Excel Drugs and Pharmaceuticals - Cefpodoxime Proxetil Oral
    # Suspension IP 100mg/5ml [NRD-485].
    "RMSC-DEBAR-1453-24-08-2026":
        "State Drug Testing Laboratory (Description failure, Minor Defects) - read from the "
        "Hindi order letter, hin+eng OCR, hand-translated.",
    # Medipol Pharmaceutical India Pvt Ltd - Acyclovir Tab IP 200mg[63].
    "RMSC-DEBAR-554-5-6-2026":
        "RMSC's own registered lab (Dissolution failure), confirmed by the State Drug Testing "
        "Laboratory (Grossly Substandard) - read from the Hindi order letter, hin+eng OCR, "
        "hand-translated.",
    # Medipol Pharmaceutica - Multivitamin Tablets NFI Formula [394], 13 batches.
    "RMSC-DEBAR-3085-20-02-2025":
        "RMSC's own registered lab (Assay of Calcium Pantothenate failure), retested at the "
        "State Drug Testing Laboratory with the same result; the Central Drugs Laboratory "
        "separately confirmed batch TMVN-163 failing too - read from the Hindi order letter, "
        "hin+eng OCR, hand-translated.",
    # Unicure India Ltd - Glipizide And Metformin Hydrochloride Tablets USP [452].
    "RMSC-DEBAR-3093-20-02-2025":
        "State Drug Testing Laboratory (Dissolution failure, confirmed on retest for batches "
        "URDT0327/URDT0328) - read from the Hindi order letter, hin+eng OCR, hand-translated.",
}


def build_row(e):
    period_m = PERIOD_RE.search(e["upto"])
    period = period_m.group(0) if period_m else "Not stated"
    court = "Not stated"
    for text in (e["reason"], e["upto"]):
        if STAY_RE.search(text or ""):
            court = one_line(text)
            break
    source_link = e["links"][0] if e["links"] else PAGE_URL
    entry_id = f"RMSC-DEBAR-{order_id_part(e['order'])}"
    where = f"Order {e['order']} (Banned_Product.html)"
    # some rows' "Up To" cell is ONLY a period phrase like "(For 2 year)" with no actual
    # date at all - don't let that non-date text masquerade as a date.
    to_date = clean_date(e["upto"]) if looks_like_date(e["upto"]) else "Not stated"
    lab = RMSC_LAB_OVERRIDES.get(entry_id, "Not stated")
    return [entry_id, one_line(e["firm"]), standard_name(e["firm"]), "Not stated", "RMSC",
            "Rajasthan", "Product (inferred)", one_line(e["product"]), "Not stated",
            to_date, period, one_line(e["reason"]), CATEGORY_PENDING,
            lab, court, source_link, where, "Auto - check", TODAY]


def collect():
    if not robots_allows(PAGE_URL):
        print(f"  BLOCKED by robots.txt - open in a browser: {PAGE_URL}")
        return []
    r = session.get(PAGE_URL, timeout=TIMEOUT)
    r.raise_for_status()
    time.sleep(WAIT_SECONDS)
    soup = BeautifulSoup(r.text, "html.parser")
    entries = parse_rows(soup)
    print(f"  {len(entries)} rows found")
    return [build_row(e) for e in entries]


LAB_IDX = COLUMNS.index("Lab that declared NSQ")


def main():
    rows = collect()
    dupe_warnings = dedup_entry_ids(rows)
    # RMSC_LAB_OVERRIDES is keyed by the FINAL (post-dedup) Entry ID (e.g.
    # "...-b" for ANG, which shares its undeduped order reference with
    # Apple Formulation) - build_row() computes the override too early
    # (before dedup_entry_ids() can add that suffix), so it is re-applied
    # here, once IDs are final. Found 2026-11-01 - without this, ANG's
    # override silently never matched and it stayed wrongly Q2.
    for r in rows:
        if r[0] in RMSC_LAB_OVERRIDES:
            r[LAB_IDX] = RMSC_LAB_OVERRIDES[r[0]]
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [30, 30, 30, 14, 10, 12, 18, 40, 13, 13, 20, 34, 22, 14, 20, 40, 40, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "RMSC.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")


if __name__ == "__main__":
    main()
