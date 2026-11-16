"""
16_ksmscl.py - Karnataka State Medical Supplies Corporation Limited
(KSMSCL) buyer file.

KSMSCL splits its blacklisting into THREE separate PDFs by category:
  - Block_Listed_Firms_281125.pdf  - firms blacklisted (dated 28.11.2025)
  - product_141125.pdf             - "LIST OF PRODUCTS BLACKLISTED" (14.11.2025)
  - lab141025.pdf                  - labs debarred for product testing (14.10.2025)

All three are scanned images (confirmed: 0 extractable characters on
every page). Tesseract's plain OCR does NOT preserve these dense
multi-column tables' alignment - confirmed by testing (see the project's
own history, PROGRESS_LOG.md) - so an EARLIER version of this script
only extracted firm/product/lab NAME FRAGMENTS via OCR, honestly leaving
every other field "Not stated".

REPLACED 2026-11-01 (PROGRESS_LOG.md, per the user's own document check,
document_checks_2026-11-01.xlsx "Karnataka full lists" sheet): the
user read all 3 PDFs directly from the page IMAGES (not OCR) and
transcribed every row in full - 21 rows total (5 Firm, 13 Product, 3
Lab), replacing the earlier 11 OCR name-fragment rows entirely. This is
the SAME kind of hand-verified-fact table this project already uses for
RMSC's Hindi order letters (build_pdfs_to_read.py RMSC_HINDI_READINGS)
and GMSCL/MPPHSCL's Lab overrides (05_gmscl.py/07_mpphscl.py
LAB_OVERRIDES) - a human read the source directly and the result is
recorded as data, not derived by a parser. The 3 PDFs are still
downloaded and SHA-256'd for evidence preservation (PROJECT_RULES.md), but no
longer OCR'd - there is nothing left for OCR to add once a full,
confident hand transcription exists.

Level = "Lab" for the lab file, consistent with MPPHSCL's existing
convention - common.py's find_matches() already excludes Level="Lab"
from cross-buyer Matches (labs are not manufacturers).

Court stays: 7 of the 13 product blacklistings are stayed by the
Karnataka High Court (recorded in "Court stay or revoked?", as written
in the source PDF).

Entry IDs: taken directly from the user's own "Entry ID (suggested)"
column - these are source-derived (buyer + level + year + a running
index within that year, not an arbitrary global counter), so they stay
stable across reruns the same way every other buyer's IDs do.

robots.txt: none on this domain (404, confirmed).

Writes:
  output/buyers/KSMSCL.xlsx   - one sheet, "Master list", 19 columns
  data/raw/KSMSCL/*.pdf       - raw PDF evidence (downloaded if not
                                 already present locally)
"""

import time
import hashlib
import csv
from datetime import date, datetime
from urllib.parse import urlparse
from urllib import robotparser

import requests
from openpyxl import Workbook

from common import (BUYERS, RAW, OUT, COLUMNS, standard_name, one_line,
                    dedup_entry_ids, write_sheet, CATEGORY_PENDING)

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")

BASE = "https://www.ksmscl.in"
FILES = {
    "Firm": ("/images/Block_Listed_Firms_281125.pdf", "Block_Listed_Firms_281125.pdf"),
    "Product": ("/files/product_141125.pdf", "product_141125.pdf"),
    "Lab": ("/files/lab141025.pdf", "lab141025.pdf"),
}

EVIDENCE_DIR = RAW / "KSMSCL"
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


def log_download(status, file_path, sha256, url):
    log_path = OUT / "download_log.csv"
    is_new = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
        if is_new:
            w.writeheader()
        w.writerow({"date_time": datetime.now().strftime("%d-%m-%Y %H:%M"), "status": status,
                    "file": str(file_path), "sha256": sha256, "url": url})


def ensure_downloaded(path_fragment, filename):
    """Downloads (and SHA-256-logs) the source PDF as evidence, even
    though this script no longer OCRs it - the hand transcription below
    still needs the original PDF kept on disk so a reader can check it
    directly (PROJECT_RULES.md evidence-preservation rule)."""
    target = EVIDENCE_DIR / filename
    if target.exists() and target.stat().st_size > 0:
        return target
    url = BASE + path_fragment
    if not robots_allows(url):
        print(f"  BLOCKED by robots.txt - open in a browser: {url}")
        return None
    r = session.get(url, timeout=TIMEOUT)
    r.raise_for_status()
    target.write_bytes(r.content)
    sha = hashlib.sha256(r.content).hexdigest()
    log_download("captured", target, sha, url)
    time.sleep(WAIT_SECONDS)
    return target


# Hand-transcribed directly from the page images by the user, 2026-11-01
# (document_checks_2026-11-01.xlsx, "Karnataka full lists" sheet) -
# NOT derived by OCR or regex. Fields: entry_id, firm_as_written, address,
# level, product, from_date, to_date, order_no_and_date, tender,
# reason_as_written, court_stay, source_pdf_key, source_page.
KARNATAKA_FULL_LIST = [
    ("KSMSCL-F-2021-22-1", "M/s Dinesh Pharma",
     "No:388/22, 8th Main Road, Hanumanthanagar, Bangalore-560019", "Firm",
     "Hand Sanitizer 5000 ml", "25-10-2021", "24-10-2026",
     "KSMSCL/NSQ/06/BLACKLIST(A)2021-22 dt 25-10-2021", "Quotations (COVID-19 drug supply)",
     "NSQ (order series KSMSCL/NSQ)", "Not stated", "Firm", "page 1"),
    ("KSMSCL-F-2021-22-2", "M/s SPY Agro Industries Ltd",
     "Nandiyal, Kannur District, Andhra Pradesh-518501", "Firm",
     "Hand Sanitizer 180 ml; 5000 ml", "25-10-2021", "24-10-2026",
     "KSMSCL/NSQ/01/BLACKLIST(A)2021-22 dt 25-10-2021", "Quotations (COVID-19 drug supply)",
     "NSQ", "Not stated", "Firm", "page 1"),
    ("KSMSCL-F-2021-22-3", "M/s Glint Cosmetics Pvt Ltd",
     "C-216-217-218 Indl. Area, MIDC Turbhe, Navi Mumbai-400705, Maharashtra", "Firm",
     "Hand Sanitizer 5000 ml", "25-10-2021", "24-10-2026",
     "KSMSCL/NSQ/01/BLACKLIST(B)2021-22 dt 25-10-2021", "Quotations (COVID-19 drug supply)",
     "NSQ", "Not stated", "Firm", "page 1"),
    ("KSMSCL-F-2023-24-1", "M/s Jackson Laboratories Pvt Ltd",
     "22-245, Manjunatha Road, Bypass Po Khannanagar, Amrithsar-14", "Firm",
     "1) Hydrocortisone Sodium Succinate Inj 60mg 2) Inj Lignocaine & Adrenaline",
     "13-09-2023", "12-09-2028", "KDL/QC/PAC/4-2/Jackson/2023-24 dt 13-09-2023",
     "KDL/e/TND/33-DRUGS-20/2013-14 dt 21-09-2013", "Quality (QC order)",
     "Stayed by Karnataka High Court, WP 26927/2023 dt 12-12-2023, till next hearing",
     "Firm", "page 1"),
    ("KSMSCL-F-2023-24-2", "M/s Sontex Rolled Bandage Works",
     "55, 2nd floor, Dawa Bazar, R.N.T Marg, Indore", "Firm",
     "Bandage cloth individually packed, Sch F-II D&C Rules, 1m x 20m",
     "13-09-2023", "12-09-2028", "KDL/QC/PAC/4-2/Sontex/2023-24 dt 13-09-2023",
     "KDL/e/TNMED/RE-TND/23/2014-15", "Quality (QC order)", "Not stated", "Firm", "page 1"),

    ("KSMSCL-P-2020-21-1", "M/s Mahendra Labs Pvt Ltd",
     "371 & 372, 10th cross, 4th Phase, Peenya Indl Area, Bangalore-560068", "Product",
     "Vitamin C chewable (Ascorbic acid) Tablet 500mg", "08-02-2021", "07-02-2026",
     "KDL/QC/28/2020-21 dt 08-02-2021",
     "KDL/AD/TND/Capsule & Tablets 2/06/2018-19 dt 11-01-2019 (IND-534)", "Quality",
     "Not stated", "Product", "page 1"),
    ("KSMSCL-P-2023-24-1", "M/s Paschim Banga Pharmaceuticals",
     "Dwarika Tower, 2nd Floor, Burdwan Road, Siliguri, Darjeeling, West Bengal-734005",
     "Product", "26.2.9 Ringer Lactate Infusion FFS IP 1x500ml", "22-03-2024", "21-03-2027",
     "KSMSCL/NSQ/09/2023-24 dt 22-03-2024",
     "KSMSCL DRUGS/TND CAT-26/24/2023-24 dt 26-05-2023 IND-0006", "NSQ",
     "Stayed by Karnataka High Court, WP 10393/2024 dt 04-04-2024", "Product", "page 1"),
    ("KSMSCL-P-2023-24-2", "M/s Vivek Pharmachem of India Pvt Ltd",
     "A1, Sethi Colony, Jaipur 302004", "Product",
     "6.7.33 Norfloxacin (Film Coated) Tablet 400mg 1x10x10", "26-03-2024", "25-03-2027",
     "KSMSCL/NSQ/03/2023-24 dt 26-03-2024",
     "HFW/KSMSCL/TND/Drugs/Tab-02/2022-23 dt 18-05-2022 IND-948", "NSQ",
     "Stayed, WP 12167/2024 dt 28-05-2024", "Product", "page 1"),
    ("KSMSCL-P-2023-24-3", "M/s Martin & Brown Bio Sciences",
     "Mallachu Majna, Opp Annapurna Hostel, PO Bhud, Tehsil Baddi, Solan-173212", "Product",
     "2.3.1 Atropine Sulphate Injection 0.6mg/ml 1x1ml Amp", "26-03-2024", "25-03-2027",
     "KSMSCL/NSQ/08/2023-24 dt 26-03-2024",
     "HFW/KSMSCL/TND/Drugs/Injection/05/2022-23 dt 20-05-2022 IND-951", "NSQ",
     "Stayed, WP 22526/2024 dt 22-08-2024", "Product", "page 1"),
    ("KSMSCL-P-2024-25-1a", "M/s Paschim Banga Pharmaceuticals", "(as above)", "Product",
     "26.2.4 Dextrose with Sodium Chloride FFS IP/BP/USP Infusion 5%+0.9% 1x500ml",
     "12-12-2024", "11-12-2027", "KSMSCL/NSQ/09/A/-Blacklist/2023-24 dt 12-12-2024",
     "IND-0006 (as above)", "NSQ", "Stayed, WP 34455/2024 dt 17-12-2024", "Product", "page 2"),
    ("KSMSCL-P-2024-25-1b", "M/s Paschim Banga Pharmaceuticals", "(as above)", "Product",
     "26.2.1 Dextrose FFS Injection 5% w/v 1x500ml", "13-02-2025", "12-02-2028",
     "KSMSCL/NSQ/02/2024-25 dt 14-02-2025", "IND-0006", "NSQ",
     "Stayed, WP 6832/2025 dt 19-03-2025", "Product", "page 2"),
    ("KSMSCL-P-2024-25-1c", "M/s Paschim Banga Pharmaceuticals", "(as above)", "Product",
     "26.2.13 Sodium Chloride (Normal Saline) FFS Infusion 0.9% w/v 1x100ml",
     "12-03-2025", "11-03-2028", "KSMSCL/NSQ/02-Blacklist/2024-25 dt 12-03-2025",
     "IND-0006", "NSQ", "Stayed, WP 12056/2025 dt 22-04-2025", "Product", "page 2"),
    ("KSMSCL-P-2024-25-1d", "M/s Paschim Banga Pharmaceuticals", "(as above)", "Product",
     "26.2.12 Sodium Chloride Injection FFS 0.9% 1x500ml", "12-03-2025", "11-03-2028",
     "KSMSCL/NSQ/02A-Blacklist/2024-25 dt 12-03-2025", "IND-0006", "NSQ",
     "Stayed, WP 12056/2025 dt 22-04-2025", "Product", "page 2"),
    ("KSMSCL-P-2024-25-2", "M/s Medivision Healthcare",
     "84/4, Span IND complex, Phase I GIDC Vatva, Ahmedabad, Gujarat-380015", "Product",
     "28.12.1 Sterile disposable perfusion (infusion) set with filter NA adult 1x1",
     "11-03-2025", "10-03-2028", "KSMSCL/NSQ/08/2024-25 dt 11-03-2025",
     "KSMSCL/Drugs/TND/Annexure-c/03/2023-24 dt 04-05-2023 IND-1062-5", "NSQ",
     "Not stated", "Product", "page 2"),
    ("KSMSCL-P-2024-25-3", "M/s Biogenetic Drugs Pvt Ltd",
     "No-66, Mathur Vaishya Nagar, Tonk Road, Jaipur-302011", "Product",
     "20.1.20 Ferrous Salts + Folic Acid 45mg elemental iron + 400mcg Pink Tablet 1x10x10",
     "28-03-2025", "27-03-2028", "KSMSCL/NSQ/15/2024-25 dt 28-03-2025",
     "KSMSCL/Drugs/TND/CAT-19/16/2023-24 dt 08-05-2025 IND-1075", "NSQ",
     "Not stated", "Product", "page 2"),
    ("KSMSCL-P-2025-26-1", "M/s Biogenetic Drugs Pvt Ltd", "(as above)", "Product",
     "1.3.11 Ibuprofen Tablet 400mg 1x10x10", "20-08-2025", "19-08-2028",
     "KSMSCL/NSQ/05-B/2025-26 dt 20-08-2025", "Annexure-A/01/2023-24 dt 04-05-2023 IND-1060",
     "NSQ", "Not stated", "Product", "page 3"),
    ("KSMSCL-P-2025-26-2a", "M/s Vivek Pharmachem of India Pvt Ltd",
     "A1, Sethi Colony, Jaipur-302004", "Product",
     "6.7.13 Ciprofloxacin Hydrochloride Tablet IP 250mg 1x10x10", "20-08-2025", "19-08-2028",
     "KSMSCL/NSQ/03/2025-26 dt 20-08-2025", "Annexure-B/02/2023-24 dt 04-05-2023 IND-1061-1",
     "NSQ", "Stayed, WP 27367/2025 dt 12-09-2025", "Product", "page 3"),
    ("KSMSCL-P-2025-26-2b", "M/s Vivek Pharmachem of India Pvt Ltd", "(as above)", "Product",
     "1.3.20 Paracetamol Tablet 650mg 1x10x10", "20-08-2025", "19-08-2028",
     "KSMSCL/NSQ/10/2024-25 dt 20-08-2025", "Annexure-A/01/2023-24 dt 04-05-2023 IND-1060",
     "NSQ", "Not stated", "Product", "page 3"),
    ("KSMSCL-P-2025-26-3", "M/s Gnosis Pharmaceuticals Pvt Ltd",
     "#619, 6th Floor, Ring Road Mall, Near Kali Mata Mandir, Delhi-110085", "Product",
     "6.7.24 Ketoconazole Tablet IP 200mg 1x10x10", "28-10-2025", "27-10-2028",
     "KSMSCL/NSQ/01-Blacklist/2025-26 dt 28-10-2025", "Annexure-B/02/2023-24 dt 04-05-2023 IND 1061-1",
     "NSQ", "Not stated", "Product", "page 3"),

    ("KSMSCL-L-2025-26-1", "M/s Standard Analytical Laboratory (ND) Pvt Ltd",
     "#69, Functional Industrial Estate, Patparganj, Delhi-110092", "Lab",
     "Testing of Ringer Lactate Infusion IP 500 ml", "30-06-2025", "29-06-2028",
     "HFW/KSMSCL/QCMIS/01-02-B/2025-26 dt 30-06-2025",
     "HFW/KSMSCL/TND/EMPL/14/2021-22 (IND-918) dt 16-02-2022", "Lab blacklisted",
     "Not stated", "Lab", "page 1"),
    ("KSMSCL-L-2025-26-2", "M/s Devansh Testing and Research Lab Pvt Ltd",
     "#94, Shiv Ganga Industrial Estate, Lakeshari, Bhagwanpur, Roorkee, Haridwar, "
     "Uttarakhand-247661", "Lab", "Testing of Ringer Lactate Infusion IP 500 ml",
     "30-06-2025", "29-06-2028", "HFW/KSMSCL/QCMIS/01-01-A/2025-26 dt 30-06-2025",
     "IND-918 (as above)", "Lab blacklisted", "Not stated", "Lab", "page 2"),
    ("KSMSCL-L-2025-26-3", "M/s Vorick Analytical Research Pvt Ltd",
     "Plot No.252, Sector-K, DSIDC, Bawana-II, Bhorgarh, Delhi-110040", "Lab",
     "Testing of drugs, chemicals and miscellaneous products", "30-06-2025", "29-06-2028",
     "HFW/KSMSCL/QCSMPL/01-03/2025-26",
     "HFW/KSMSCL/TND/LAB-EMPANELMENT/12/2022-23 (IND-958) dt 30-06-2022", "Lab debarred",
     "Not stated", "Lab", "page 3"),
]

SOURCE_URL_BY_KEY = {"Firm": BASE + FILES["Firm"][0], "Product": BASE + FILES["Product"][0],
                     "Lab": BASE + FILES["Lab"][0]}
SOURCE_FILE_BY_KEY = {"Firm": FILES["Firm"][1], "Product": FILES["Product"][1],
                      "Lab": FILES["Lab"][1]}


def build_rows():
    rows = []
    for (entry_id, firm, address, level, product, frm, to, order, tender, reason, court,
         source_key, page) in KARNATAKA_FULL_LIST:
        reason_display = f"{reason} (order {order}; tender {tender})"
        rows.append([entry_id, one_line(firm), standard_name(firm), one_line(address), "KSMSCL",
                    "Karnataka", level, one_line(product), frm, to, "Not stated",
                    reason_display, CATEGORY_PENDING, "Not stated", court,
                    SOURCE_URL_BY_KEY[source_key], f"{page} ({SOURCE_FILE_BY_KEY[source_key]})",
                    "Checked (hand-transcribed from the page images by the user, "
                    "document_checks_2026-11-01.xlsx)", TODAY])
    return rows


def main():
    for level, (fragment, filename) in FILES.items():
        print(f"\n--- {level}: {filename} ---")
        path = ensure_downloaded(fragment, filename)
        if path is None:
            print("  (hand transcription used regardless - see KARNATAKA_FULL_LIST)")

    all_rows = build_rows()
    dupe_warnings = dedup_entry_ids(all_rows)
    if dupe_warnings:
        print("\nWARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, all_rows,
               [20, 40, 34, 50, 10, 12, 10, 70, 13, 13, 12, 50, 14, 14, 50, 50, 24, 50, 12])
    wb.active.title = "Master list"
    out = BUYERS / "KSMSCL.xlsx"
    wb.save(out)

    print(f"\nSaved: {out}")
    print(f"  Total rows: {len(all_rows)}")
    for level in ("Firm", "Product", "Lab"):
        print(f"  {level}: {sum(1 for r in all_rows if r[6] == level)}")
    stayed = sum(1 for r in all_rows if "stayed" in str(r[14]).lower())
    print(f"  Court-stayed: {stayed}")


if __name__ == "__main__":
    main()
