"""
15_rate_contracts.py - RMSC's "Approved Rate List - Drugs, Surgical and
Sutures" (Step 4b, RMSC only - the one confirmed-complete source; see
output/rate_contract_sources.xlsx for why GMSCL/KMSCL/TNMSC/APMSIDC/
MPPHSCL/PMBI are not built yet).

This is a DIFFERENT kind of data than the 19-column blacklist master list
(drug-level pricing, not a firm-ban record), so it gets its own output
shape rather than being forced into COLUMNS.

Confirmed columns on the live page (read directly, not guessed): DRUG NAME
| CODE | TENDER NO | SUPPLIER | SUPPLIER CONTACT | UNIT | RATE | GST RATE |
LEVEL | CONTRACT FROM DATE | CONTRACT TO DATE. There is NO separate
Manufacturer field anywhere on this page - "Manufacturer (if stated)" is
honestly "Not stated" for every row, not guessed, because the RC holder
(SUPPLIER) may be a distributor/marketer rather than the actual factory,
and this page never says which.

"Supplier Contact" (a personal name/phone/email) is deliberately NOT
captured - company-level data only.

IDs: RMSC-RC-<tender code>-<drug code>-<supplier code>.
  - <tender code> pulls the "NIB-17"/"NIT-xx"-style reference and the
    year range out of the full Tender No text (e.g.
    "F.02(209)/RMSCL/PROC./SS/NIB-17/2025-26/927" -> "NIB17-2025-26") -
    short and still traceable back to the real tender number, which is
    kept in full in the "Tender No" column regardless.
  - <drug code> is the site's own short code (e.g. "S119") - already
    unique per drug.
  - <supplier code> is the first word of the STANDARDIZED supplier name,
    uppercased, alnum only, first 8 characters (e.g. "Sterimed Medical
    Devices Pvt Ltd" -> "STERIMED") - needed because the SAME drug under
    the SAME tender can have MULTIPLE approved suppliers at different
    price/L-levels (confirmed on the real page: drug code S122 has both
    "Syana India Associater Pvt Ltd" at L1 and "Vitromed Healthcare
    Jaipur" at L4, same tender). The "-b"/"-c" rule is still the final
    safety net if two suppliers happen to reduce to the same short code.

"Supplier (standard)" uses the SAME `standard_name()` and
`strip_leaked_address()` functions used for blacklist data, so names join
correctly against the master list (e.g. for 16_rc_vs_bans.py).

Every run saves the raw HTML to data/raw/RMSC_RC/DrugsRC_<YYYY-MM-DD>.html
and logs its SHA-256 in output/download_log.csv (the SAME log
01_download.py writes to - one shared evidence trail, not a separate
file) - repeated runs build a dated history of the page's content over
time, the same way the project already does for downloaded PDFs.

Also checks (report only, does not download): the RMSC Archives page for
older rate-contract tender documents, and the Wayback Machine for earlier
captures of this exact page - both printed to the console, per the user's
"report which dates are available, don't download them yet" instruction.

Writes:
  output/rate_contracts/RMSC.xlsx   - one sheet, "Rate contracts"
  data/raw/RMSC_RC/DrugsRC_<date>.html  - raw HTML evidence, one per run
  output/download_log.csv           - one new row logging that HTML's SHA-256
"""

import re
import csv
import time
import hashlib
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib import robotparser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import OUT, RAW, standard_name, strip_leaked_address, clean_date, dedup_entry_ids, \
    write_sheet

URL = ("https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-"
       "corporation-ltd-/en/Approved-Rate-Lists/DrugsRC.html")
ARCHIVES_URL = ("https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-"
                "services-corporation-ltd-/en/library/Archives.html")

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")
TODAY_ISO = date.today().strftime("%Y-%m-%d")

RC_OUT_DIR = OUT / "rate_contracts"
RC_OUT_DIR.mkdir(parents=True, exist_ok=True)
EVIDENCE_DIR = RAW / "RMSC_RC"
EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)

session = requests.Session()
session.headers.update(HEADERS)
_robots_cache = {}


def robots_allows(url):
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if root not in _robots_cache:
        rp = robotparser.RobotFileParser()
        try:
            r = session.get(root + "/robots.txt", timeout=20)
            rp.parse(r.text.splitlines() if r.status_code == 200 else [])
        except requests.RequestException:
            rp.parse([])
        _robots_cache[root] = rp
    return _robots_cache[root].can_fetch(HEADERS["User-Agent"], url)


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def log_download(status, file_path, sha256, url):
    log_path = OUT / "download_log.csv"
    is_new = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
        if is_new:
            w.writeheader()
        w.writerow({"date_time": datetime.now().strftime("%d-%m-%Y %H:%M"), "status": status,
                    "file": str(file_path), "sha256": sha256, "url": url})


TENDER_SHORT_RE = re.compile(r"(NI[BT]-?\d+)", re.I)
YEAR_RANGE_RE = re.compile(r"\b(\d{4}-\d{2,4})\b")


def tender_id_part(tender_no):
    tender_no = tender_no or ""
    m = TENDER_SHORT_RE.search(tender_no)
    code = re.sub(r"[^A-Za-z0-9]", "", m.group(1)).upper() if m else ""
    y = YEAR_RANGE_RE.search(tender_no)
    year = y.group(1) if y else ""
    parts = [p for p in (code, year) if p]
    if parts:
        return "-".join(parts)
    cleaned = re.sub(r"[^0-9A-Za-z]+", "-", tender_no).strip("-")
    return cleaned[:30] or "UNKNOWNTENDER"


def supplier_code_part(supplier_standard):
    if not supplier_standard:
        return "UNKNOWN"
    first_word = supplier_standard.split()[0]
    code = re.sub(r"[^A-Z0-9]", "", first_word)
    return code[:8] or "UNKNOWN"


def parse_rows(soup):
    table = soup.find("table")
    if table is None:
        return []
    out = []
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if len(cells) < 11 or cells[0].strip().upper() == "DRUG NAME":
            continue                                        # header row
        (drug_name, code, tender_no, supplier, _contact, unit, rate, gst, level,
         frm, to) = cells[:11]
        out.append({"drug_name": drug_name, "code": code, "tender_no": tender_no,
                    "supplier": supplier, "unit": unit, "rate": rate, "gst": gst,
                    "level": level, "frm": frm, "to": to})
    return out


def build_row(e):
    supplier = strip_leaked_address(e["supplier"])
    supplier_std = standard_name(supplier)
    entry_id = f"RMSC-RC-{tender_id_part(e['tender_no'])}-{e['code']}-{supplier_code_part(supplier_std)}"
    return [entry_id, e["drug_name"], e["code"], e["tender_no"], supplier, supplier_std,
            "Not stated", e["rate"], e["gst"], e["level"], clean_date(e["frm"]), clean_date(e["to"]),
            TODAY, URL]


def report_archive_and_wayback():
    print("\n--- Checking for older RC data (report only, nothing downloaded) ---")
    try:
        if robots_allows(ARCHIVES_URL):
            r = session.get(ARCHIVES_URL, timeout=TIMEOUT)
            time.sleep(WAIT_SECONDS)
            soup = BeautifulSoup(r.text, "html.parser")
            rc_links = [a for a in soup.find_all("a", href=True)
                        if "rate contract" in a.get_text(" ", strip=True).lower()]
            print(f"RMSC Archives page: {len(rc_links)} links mentioning 'Rate Contract' found "
                  f"(older NIT/NIB tender documents, e.g. NIT-10 from the 2012-14 period through "
                  f"NIB-12x+ more recently) - these are individual TENDER NOTICES, not periodic "
                  f"snapshots of this same DrugsRC.html table. Not downloaded.")
        else:
            print("RMSC Archives page blocked by robots.txt.")
    except requests.RequestException as e:
        print(f"Could not check RMSC Archives page: {e}")

    try:
        r = requests.get("https://web.archive.org/cdx/search/cdx", params={
            "url": URL.replace("https://", "").replace("http://", ""),
            "output": "json", "matchType": "prefix", "collapse": "timestamp:8", "limit": "10000",
        }, timeout=60)
        if r.status_code == 200 and r.text.strip().startswith("["):
            import json
            rows = json.loads(r.text)
            n = len(rows) - 1
            if n > 0:
                print(f"Wayback Machine: {n} distinct-day snapshots of this exact page found, "
                      f"from {rows[1][1][:8]} to {rows[-1][1][:8]}. Not downloaded.")
            else:
                print("Wayback Machine: no snapshots found for this page.")
        else:
            print(f"Wayback Machine check did not return usable data (status {r.status_code}).")
    except requests.RequestException as e:
        print(f"Could not check Wayback Machine: {e}")


def main():
    if not robots_allows(URL):
        print(f"BLOCKED by robots.txt - open in a browser: {URL}")
        return

    r = session.get(URL, timeout=TIMEOUT)
    r.raise_for_status()
    time.sleep(WAIT_SECONDS)

    html_path = EVIDENCE_DIR / f"DrugsRC_{TODAY_ISO}.html"
    html_path.write_text(r.text, encoding="utf-8")
    sha = sha256_bytes(r.content)
    log_download("captured", html_path, sha, URL)
    print(f"Saved evidence: {html_path}")
    print(f"  SHA-256 logged to output/download_log.csv: {sha}")

    soup = BeautifulSoup(r.text, "html.parser")
    entries = parse_rows(soup)
    rows = [build_row(e) for e in entries]

    dupe_warnings = dedup_entry_ids(rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    header = ["Entry ID", "Drug Name", "Drug Code", "Tender No", "Supplier (as written)",
              "Supplier (standard)", "Manufacturer (if stated)", "Rate", "GST Rate", "Level",
              "Contract From", "Contract To", "Date captured", "Source link"]

    wb = Workbook()
    write_sheet(wb.active, header, rows, [30, 44, 10, 30, 30, 26, 20, 10, 10, 8, 14, 14, 12, 60])
    wb.active.title = "Rate contracts"
    out = RC_OUT_DIR / "RMSC.xlsx"
    wb.save(out)

    print(f"\nSaved: {out}")
    print(f"  Rows: {len(rows)}")
    suppliers = {r[5] for r in rows}
    print(f"  Distinct suppliers (standard name): {len(suppliers)}")
    froms = [r[10] for r in rows if r[10] and r[10] != "Not stated"]
    tos = [r[11] for r in rows if r[11] and r[11] != "Not stated"]
    if froms and tos:
        print(f"  Contract date range: {min(froms)} to {max(tos)}")
    stated_mfr = sum(1 for r in rows if r[6] != "Not stated")
    print(f"  Rows with a manufacturer stated: {stated_mfr} (0 expected - no such field exists "
          f"on this page, confirmed, not guessed)")

    report_archive_and_wayback()


if __name__ == "__main__":
    main()
