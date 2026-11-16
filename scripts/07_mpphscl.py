"""
06_mpphscl.py - Madhya Pradesh Public Health Services Corporation Limited
(MPPHSCL) buyer file.

Fetches all three of MPPHSCL's lists:
  - BlackListFirm.aspx     (Firms)   - Level "Firm"
  - BlackListLabs.aspx     (Labs)    - Level "Lab"    (kept, per instruction,
                                        but EXCLUDED from the Matches sheet -
                                        labs are not manufacturers, so
                                        matching a lab name against a firm
                                        name would be a false "cross-buyer
                                        case")
  - BlackListProduct.aspx  (Products) - Level "Product"

Two real site quirks, confirmed by testing, not guessed:
  1. The server needs legacy TLS renegotiation enabled to connect at all
     (confirmed error: "[SSL: UNSAFE_LEGACY_RENEGOTIATION_DISABLED]").
     Handled with a custom SSLContext that sets ssl.OP_LEGACY_SERVER_CONNECT.
     Certificate verification is NOT disabled anywhere - ssl.create_default_
     context() keeps verify_mode=CERT_REQUIRED, confirmed by checking it
     directly before writing this script.
  2. Fetching a list page cold (no prior request in the session) sometimes
     behaves differently - fetching the homepage first, in the same
     session, avoids any issue. This is normal session/cookie behavior, not
     a deliberate anti-bot measure to route around.

All three pages are small enough to load as ONE table each (confirmed: 140,
2, and 236 data rows respectively) - no pagination needed.

Court stays / revocations: the "Corrigendum" column sometimes carries this
in free text (e.g. "As per the order of Hon'ble High Court of MP under
WP-36448-2026, the debarring is stayed till the final decision") - copied
AS WRITTEN into "Court stay or revoked?".

IDs: MPPHSCL-<FIRM/LAB/PRODUCT>-<order serial>-<year>, e.g.
MPPHSCL-FIRM-1373158-2026, pulled out of the "Blacklisted Order" text (e.g.
"Sno 1373158/F 0397/26/MPPHSCL/TECH/2026"). The full order text is kept in
"Page / row in source" regardless.

Writes:
  output/buyers/MPPHSCL.xlsx  with one sheet:
    - Master list   : MPPHSCL rows (Firms + Labs + Products), in the 19
                       master-list columns
"""

import re
import ssl
import time
from datetime import date
from urllib.parse import urljoin, urlparse
from urllib import robotparser

import requests
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import BUYERS, COLUMNS, standard_name, clean_date, one_line, \
    dedup_entry_ids, write_sheet, strip_leaked_address, CATEGORY_PENDING

BASE = "https://mpphscl.mp.gov.in/"
PAGES = [("BlackListFirm.aspx", "Firm", "FIRM"),
         ("BlackListLabs.aspx", "Lab", "LAB"),
         ("BlackListProduct.aspx", "Product", "PRODUCT")]

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")


class LegacyRenegotiationAdapter(HTTPAdapter):
    """Allows the TLS legacy renegotiation this server needs, WITHOUT
    turning off certificate verification (ssl.create_default_context()
    keeps verify_mode=CERT_REQUIRED and check_hostname=True by default)."""
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.options |= ssl.OP_LEGACY_SERVER_CONNECT
        kwargs["ssl_context"] = ctx
        super().init_poolmanager(*args, **kwargs)


session = requests.Session()
session.headers.update(HEADERS)
session.mount("https://", LegacyRenegotiationAdapter())

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


# Older orders use "Sn<num>" instead of "Sno <num>" (e.g. "Sn3599/F292/2023/..."),
# and some have no prefix at all, just a bare leading number
# (e.g. "5281/MPPHSCL/TECH/2019/F-325/19") - both confirmed by reading the
# actual data, not guessed.
ORDER_SNO_RE = re.compile(r"(?:sno|sn|sr)\s*/?\s*(\d+)", re.I)
YEAR_RE = re.compile(r"\b(20\d{2})\b")
FULL_YEAR_RE = re.compile(r"\A20\d{2}\Z")


def order_id_part(order_text):
    order_text = order_text or ""
    m = ORDER_SNO_RE.search(order_text)
    if m:
        sno = m.group(1)
    else:
        sno = None
        for num in re.findall(r"\d+", order_text):
            if FULL_YEAR_RE.match(num):
                continue                                    # skip a bare year, not a serial
            sno = num
            break
    years = YEAR_RE.findall(order_text)
    year = years[-1] if years else None
    if sno and year:
        return f"{sno}-{year}"
    if sno:
        return sno
    if year:
        return year
    cleaned = re.sub(r"[^0-9A-Za-z-]", "", order_text or "")[:30]
    return cleaned or "UNKNOWNORDER"


def parse_table(soup, has_product_col):
    table = soup.find("table")
    if table is None:
        return []
    out = []
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        links = [urljoin(BASE, a["href"]) for a in tr.find_all("a", href=True)]
        if not cells or not cells[0].strip().isdigit():
            continue                                        # header row
        min_cols = 11 if has_product_col else 10
        if len(cells) < min_cols:
            continue
        if has_product_col:
            sr, _cat, product, firm, frm, to, order, reason, remark, _doc, corrigendum = cells[:11]
        else:
            sr, _cat, firm, frm, to, order, reason, remark, _doc, corrigendum = cells[:10]
            product = "-"
        out.append({"sr": sr, "product": product, "firm": firm, "frm": frm, "to": to,
                     "order": order, "reason": reason, "remark": remark,
                     "corrigendum": corrigendum, "links": links})
    return out


# MPPHSCL's live blacklist page often cites only a bare clause number
# ("TENDER TERMS AND CONDITON 12.5", "Annexure IX clause No 7") with no
# further explanation. The glossary of what each clause actually MEANS
# (read from the real tender/order documents) used to live here as
# CLAUSE_TEXT_OVERRIDES/ENTRY_ID_OVERRIDES, applied by this script's own
# build_row(). Moved to 10_merge_master.py 2026-11-01 (PROGRESS_LOG.md
# Entry 19, checklist item 13) so classification - the generic
# reason_category() keyword match AND these hand-verified overrides alike -
# runs ONCE, centrally, on every merge, rather than needing this script
# re-run separately whenever a shared word list changes (exactly how the
# Meril/GMSCL bug happened - see Entries 16-17). This script now only
# writes the raw "Reason (as written)" text (remark folded in, see
# build_row() below) and the lab name; see 10_merge_master.py for the
# glossary itself and the override logic.

# MPPHSCL's live page never names a lab (no such column exists) - "Lab
# that declared NSQ" is honestly "Not stated" for every row by default,
# same pattern as GMSCL (05_gmscl.py LAB_OVERRIDES) and RMSC (08_rmsc.py
# RMSC_LAB_OVERRIDES). Where the full order PDF was actually read and
# names a specific lab, that hand-verified fact is recorded here instead
# of staying "Not stated" - this directly feeds quality_level()'s Q1/Q2
# split. Added 2026-11-01 per the user's own document check
# (document_checks_2026-11-01.xlsx, "Case verdicts" sheet) - these
# 4 Zee rows were wrongly Q2 because this field was never populated, even
# though the order letters themselves name a statutory lab.
MPPHSCL_LAB_OVERRIDES = {
    # Zee Laboratories - Dexamethasone Sodium Phosphate Inj, batch 323-885.
    "MPPHSCL-PRODUCT-4102-2024":
        "Central Drugs Laboratory (CDL), Kolkata - Government Analyst report KOL/LS/2024-25/218 "
        "dt 07-05-2024 (read from the full order PDF - not on the live page itself).",
    # Zee Laboratories - Gentamicin Inj 40mg/ml 2ml, batch 323-216.
    "MPPHSCL-PRODUCT-4790-2024":
        "Central Drugs Laboratory (CDL), Kolkata - Government Analyst report KOL/LS/2024-25/370 "
        "dt 30-05-2024 (read from the full order PDF - not on the live page itself).",
    # Zee Laboratories - Ceftriaxone Inj IP 1g, batch 324-476.
    "MPPHSCL-PRODUCT-7882-2024":
        "Central Drugs Laboratory (CDL), Kolkata - Government Analyst report KOL/LS/2024-25/952 "
        "dt 20-09-2024 (read from the full order PDF - not on the live page itself).",
    # Zee Laboratories - Ranitidine Tab 150mg, batch 419-776.
    "MPPHSCL-PRODUCT-8078-2021":
        "Government Analyst, Controller Food & Drug Administration, Madhya Pradesh - report "
        "42N dt 19-11-2020 (read from the full order PDF - not on the live page itself).",
}

# Astra Life Care's MPPHSCL ban is a REVOKED price-above-ceiling case, not
# an active one - order 2335 dt 26-07-2018 LIFTS the firm's earlier 1-year
# debarment after it agreed to NPPA/DPCO rates. The live page's own
# "Corrigendum" column (already copied into "Court stay or revoked?" as
# written, see build_row() below) does not carry this - it was found only
# by reading the full order PDF. Added 2026-11-01 per the user's own
# document check (document_checks_2026-11-01.xlsx, "Random-20
# check" sheet, row 1).
MPPHSCL_COURT_OVERRIDES = {
    "MPPHSCL-FIRM-1507-2018":
        "REVOKED - order 2335 dt 26-07-2018 lifts the earlier 1-year debarment after the firm "
        "agreed to NPPA/DPCO rates (read from the full order PDF - not on the live page itself).",
    # Meril Endo Surgery - RESOLVED 2026-11-07 (source_issues.csv, 07-11-2026 entry):
    # MPPHSCL-PRODUCT-166-2024 is confirmed the SAME real action as this order (not a second,
    # separate product ban) - MPPHSCL-PRODUCT-166-2024.pdf and MPPHSCL-FIRM-7876-2024.pdf are
    # SHA-256 IDENTICAL files (not an OCR read - a hash fact); the "-b" dedup twin's own PDF is a
    # different scan of the same dispatch letter (OCR cross-reference: same order 7876 dt
    # 06.11.2024, item MIS93, EMD Rs 2,00,000, firm/address all match - OCR-derived, PROJECT_RULES.md
    # rule F, pending the user's own look). This row's firm-level listing is therefore SUPERSEDED
    # by the appeal outcome, which survives only as MPPHSCL-PRODUCT-166-2024's own item-level
    # entry - kept here as the ORIGINAL raw Firm-page listing (Level/Product(s) no longer
    # overridden to impersonate a second product row - see MPPHSCL_LEVEL_OVERRIDES/
    # MPPHSCL_PRODUCT_OVERRIDES below, both now empty), excluded from every active-ban count via
    # this non-"Not stated" Court field (build_rc_vs_bans.py classify() already treats that as
    # group "X" - stayed/revoked/superseded - unconditionally).
    "MPPHSCL-FIRM-7876-2024":
        "SUPERSEDED BY APPEAL - see MPPHSCL-PRODUCT-166-2024 (same action - SHA-256-identical "
        "PDF, confirmed 2026-11-07). Firm-level ban SET ASIDE on appeal, Commissioner order 167 "
        "dt 20-01-2025 (page 2) - the firm remains blacklisted only for item MIS93, 2 years from "
        "06-11-2024, for quoting a single-foil rate instead of per pack; EMD Rs 2,00,000 "
        "forfeited. Read from the full order PDF, checked by the user 2026-11-02.",
}

# Meril's appeal (above) means order 7876's own ORIGINAL firm-level listing
# on MPPHSCL's live Firm page no longer reflects what actually happened -
# the ban survived only as an ITEM-level debarment. RESOLVED 2026-11-07
# (source_issues.csv): that surviving item-level debarment IS
# MPPHSCL-PRODUCT-166-2024 - MPPHSCL's own, separately-scraped Product-list
# row for the same item (MIS93), same dates (06-Nov-2024 to 05-Nov-2026),
# confirmed via a SHA-256-identical PDF. So this row is NOT also turned
# into a second, duplicate product-level row - Level/Product(s) are left
# as the real, raw Firm-page data (Level "Firm", Product(s) "-"); the
# supersession itself is recorded in MPPHSCL_COURT_OVERRIDES above, which
# already excludes it from every active-ban count (see
# build_rc_vs_bans.py's classify()). These two dicts are kept (now empty)
# so the next override added here follows the same pattern/comment.
MPPHSCL_LEVEL_OVERRIDES = {}
MPPHSCL_PRODUCT_OVERRIDES = {}


def build_row(e, level, list_type, page_url, page_label):
    court = one_line(e["corrigendum"]) if e["corrigendum"].strip() else "Not stated"
    source_link = e["links"][0] if e["links"] else page_url
    entry_id = f"MPPHSCL-{list_type}-{order_id_part(e['order'])}"
    if entry_id in MPPHSCL_COURT_OVERRIDES:
        court = MPPHSCL_COURT_OVERRIDES[entry_id]
    where = f"Sr {e['sr']}, Order {e['order']} ({page_label})"
    reason_display = one_line(e["reason"])
    # Remark often carries the real reason (e.g. Zee's MPPHSCL ban: the
    # visible Reason column says "Violation of Tender Terms and Condition",
    # but Remark says "Blacklisted/Debarred for Quality Failure" - checklist
    # item 2). Folded into the one stored raw field, rather than left to
    # silently affect classification invisibly as before, so it is both
    # visible to a human reading this column AND available to
    # 10_merge_master.py's centralized classify_row() - which needs exactly
    # this combined text to classify these rows the same way this script
    # used to.
    remark_clean = one_line(e["remark"]) if e["remark"] else ""
    if remark_clean and remark_clean.lower() != reason_display.lower():
        reason_display = f"{reason_display} [Remark: {remark_clean}]"
    # MPPHSCL's own "Company Name" field often has the full address appended
    # directly (no separate address column exists to split from) - e.g.
    # "M/s Alpa Laboratries Ltd., Indore". Trim that off before standardizing.
    firm = strip_leaked_address(e["firm"])
    lab = MPPHSCL_LAB_OVERRIDES.get(entry_id, "Not stated")
    level = MPPHSCL_LEVEL_OVERRIDES.get(entry_id, level)
    product_display = MPPHSCL_PRODUCT_OVERRIDES.get(entry_id, one_line(e["product"]))
    return [entry_id, one_line(firm), standard_name(firm), "Not stated", "MPPHSCL",
            "Madhya Pradesh", level, product_display, clean_date(e["frm"]), clean_date(e["to"]),
            "Not stated", reason_display, CATEGORY_PENDING, lab,
            court, source_link, where, "Auto - check", TODAY]


def collect():
    rows = []
    session.get(BASE, timeout=TIMEOUT)            # establish a session before the data pages
    time.sleep(WAIT_SECONDS)
    for page, level, list_type in PAGES:
        url = urljoin(BASE, page)
        if not robots_allows(url):
            print(f"  BLOCKED by robots.txt - open in a browser: {url}")
            continue
        r = session.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        time.sleep(WAIT_SECONDS)
        soup = BeautifulSoup(r.text, "html.parser")
        entries = parse_table(soup, has_product_col=(level == "Product"))
        print(f"  {page}: {len(entries)} rows")
        for e in entries:
            rows.append(build_row(e, level, list_type, url, page))
    return rows


def main():
    rows = collect()
    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, rows, [20, 34, 30, 14, 10, 16, 14, 30, 13, 13, 12, 34, 22, 14, 40, 40, 50, 16, 12])
    wb.active.title = "Master list"
    out = BUYERS / "MPPHSCL.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows : {len(rows)}")
    for level in ("Firm", "Lab", "Product"):
        print(f"    {level}: {sum(1 for r in rows if r[6] == level)}")


if __name__ == "__main__":
    main()
