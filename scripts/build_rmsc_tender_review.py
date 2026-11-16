"""
build_rmsc_tender_review.py - Step 4b regulatory layer: download the RMSC
drug tender document(s) behind the CURRENT rate contracts in
output/rate_contracts/RMSC.xlsx, and report whether each one contains a
clause on bidders debarred/blacklisted by OTHER states or agencies, and
whether it states estimated annual quantities per drug.

Not part of the numbered pipeline - a one-off read/report script, run by
hand. Does NOT compute any rupee values (per the user's instruction) -
only reports whether quantity figures exist in the document.

Scope decision (explained to the user before running): RMSC.xlsx's 2,084
rate-contract rows cite 72 distinct raw tender-number strings, but most of
that variation is formatting noise (spacing, "PROC." vs "PROCUREMENT",
etc.) around a much smaller set of real NIB tenders. The four NIB numbers
that account for the most rows - NIB-14/2024/69 (535 rows), NIB-08/2024
(417), NIB-01/2025 (364), NIB-10/2024 (131) - cover about 70% of all rows
between them, so they were used as "the drug tender document(s) behind
the current RCs" rather than chasing all 72 raw strings individually.

Found on RMSC's own Drug-Tenders listing page (not the Archives page,
which turned out to hold OLDER/historical tender notices, not these
current ones - see PROGRESS_LOG.md):
  NIB-14/2024/69 -> Tender14.pdf.pdf
  NIB-01/2025    -> NIB 01.pdf
  NIB-10/2024    -> Tender_NIb10_DM.pdf
  NIB-08/2024    -> NOT FOUND as a distinct main tender notice on this
                     listing page (only a "Corrigendum III for Drugs
                     NIB-08/2024" exists there, implying an earlier
                     original tender notice that either isn't listed on
                     this page or uses a URL slug this script's search
                     didn't anticipate). Logged as an open item, not
                     guessed at.

robots.txt: confirmed earlier in this project that
rmsc.health.rajasthan.gov.in has no robots.txt (404) - no restriction.
Re-checked here before every download anyway, the same as 15_rate_contracts.py.

Writes:
  data/raw/RMSC_Tenders/<slug>.pdf           - raw PDF evidence
  output/download_log.csv                    - one row per download (shared log)
  output/rmsc_tender_review.xlsx              - findings (clause text + quantity check)
"""

import csv
import time
import hashlib
from datetime import datetime
from urllib.parse import urlparse
from urllib import robotparser

import requests
import pdfplumber
from openpyxl import Workbook

from common import OUT, RAW, write_sheet

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 5
TIMEOUT = 60
BASE = "https://rmsc.health.rajasthan.gov.in"

TENDERS = [
    {"tender": "NIB-14/2024/69", "rc_rows": 535, "slug": "NIB14_2024_69_Tender14",
     "pdf_path": "/content/dam/doitassets/Medical-and-Health-Portal/rajasthan-medical-corporation/"
                 "pdf/DRUG_TENDERS/Tender14.pdf.pdf",
     "notice_page": "/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/"
                     "Tenders/Drug-Tenders/drug---medicine-tender-nib-14.html"},
    {"tender": "NIB-01/2025", "rc_rows": 364, "slug": "NIB01_2025_NIB01",
     "pdf_path": "/content/dam/doitassets/Medical-and-Health-Portal/rajasthan-medical-corporation/"
                 "pdf/DRUG_TENDERS/NIB 01.pdf",
     "notice_page": "/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/"
                     "Tenders/Drug-Tenders/NIB_01_871_652025.html"},
    {"tender": "NIB-10/2024", "rc_rows": 131, "slug": "NIB10_2024_Tender_NIb10_DM",
     "pdf_path": "/content/dam/doitassets/Medical-and-Health-Portal/rajasthan-medical-corporation/"
                 "pdf/DRUG_TENDERS/Tender_NIb10_DM.pdf",
     "notice_page": "/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/"
                     "Tenders/Drug-Tenders/NIb10_11092024.html"},
]

EVIDENCE_DIR = RAW / "RMSC_Tenders"
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


# Keywords used to find the debarment/blacklist clause - looked for
# together with "debar"/"blacklist" word stems rather than a single exact
# phrase, since tender wording varies.
DEBAR_KEYWORDS = ["debar", "black list", "blacklist", "blacklisted"]

# A per-drug quantity TABLE (the real thing Task 4 is asking about) is
# distinguished from generic procurement boilerplate like "procuring
# entity's right to vary quantity" by requiring BOTH an "estimated" word
# AND a "qty"/"quantity" word to appear together with a drug-code-style
# table header ("code" or "s.no") on the SAME page - a plain mention of
# the word "quantity" in a general clause is not enough evidence on its
# own (confirmed as a false signal during testing: NIB-14/NIB-10 mention
# "quantity" 36-60 times each in clauses unrelated to any per-drug figure).
QTY_TABLE_HINTS = ["estimated", "qty"]


def _is_toc_context(text, idx, window=150):
    """A Table-of-Contents / Annexure-index entry, not the real clause -
    recognisable because several "Annexure -" list markers cluster right
    before the match (confirmed false-positive during testing: the index
    page lists "Annexure - IX Guidelines for Blacklisting/Debarring..."
    which matches the same keywords as the real clause 5 pages later)."""
    before = text[max(0, idx - window):idx]
    return before.count("Annexure") >= 2 or before.count("eruxennA") >= 1


def search_pdf(pdf_path):
    """Return (clause_text_or_None, clause_page_or_None, quantity_found_bool,
    quantity_page_or_None) by scanning every page's extracted text. The
    clause quote is extended onto the NEXT page too if the match sits near
    the end of its page - the real clause runs on to a second page in both
    PDFs tested, and cutting there mid-sentence would misquote it."""
    clause_text, clause_page = None, None
    qty_found, qty_page = False, None
    page_texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            page_texts.append(text)
            lower = text.lower()
            if clause_text is None:
                for kw in DEBAR_KEYWORDS:
                    idx = lower.find(kw)
                    if idx != -1 and not _is_toc_context(text, idx):
                        start = max(0, idx - 400)
                        end = min(len(text), idx + 1100)
                        quote = text[start:end].strip()
                        if end >= len(text) - 50 and i < len(pdf.pages):
                            next_text = pdf.pages[i].extract_text() or ""  # pdf.pages[i] = page i+1
                            quote += "\n[continues to next page]\n" + next_text[:700].strip()
                        clause_text = quote
                        clause_page = i
                        break
            if not qty_found and all(h in lower for h in QTY_TABLE_HINTS) and \
                    ("code" in lower or "s.no" in lower.replace(" ", "")):
                qty_found = True
                qty_page = i
    return clause_text, clause_page, qty_found, qty_page


def main():
    results = []
    for t in TENDERS:
        url = BASE + t["pdf_path"]
        print(f"\n--- {t['tender']} ({t['rc_rows']} RC rows) ---")
        print(f"Notice page: {BASE + t['notice_page']}")
        print(f"PDF: {url}")

        if not robots_allows(url):
            print("  BLOCKED by robots.txt - skipping, open by hand:", url)
            results.append({**t, "status": "blocked_by_robots", "clause": None,
                            "clause_page": None, "qty_found": None, "qty_page": None, "sha256": None})
            continue

        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()
        except requests.RequestException as e:
            print(f"  DOWNLOAD FAILED: {e}")
            results.append({**t, "status": f"download_failed: {e}", "clause": None,
                            "clause_page": None, "qty_found": None, "qty_page": None, "sha256": None})
            time.sleep(WAIT_SECONDS)
            continue

        ctype = r.headers.get("Content-Type", "")
        if "pdf" not in ctype.lower() and not r.content[:5].startswith(b"%PDF-"):
            print(f"  WARNING: response does not look like a PDF (Content-Type={ctype}) - "
                  f"saving anyway for inspection, logging to source_issues.csv recommended.")

        local_path = EVIDENCE_DIR / f"{t['slug']}.pdf"
        local_path.write_bytes(r.content)
        sha = sha256_bytes(r.content)
        log_download("captured", local_path, sha, url)
        print(f"  Saved: {local_path}  ({len(r.content):,} bytes)")
        print(f"  SHA-256 logged to output/download_log.csv: {sha}")

        try:
            with pdfplumber.open(local_path) as pdf:
                total_chars = sum(len(p.extract_text() or "") for p in pdf.pages)
                page_count = len(pdf.pages)
        except Exception as e:
            print(f"  PDF COULD NOT BE OPENED: {e}")
            results.append({**t, "status": "unreadable_pdf", "clause": None,
                            "clause_page": None, "qty_found": None, "qty_page": None, "sha256": sha})
            time.sleep(WAIT_SECONDS)
            continue

        if total_chars == 0:
            print(f"  SCANNED/IMAGE PDF - zero extractable text across all {page_count} pages. "
                  f"Same OCR blocker already logged for RMSC's Medipol PDF, KSMSCL and PMBI "
                  f"(this project's OCR environment is confirmed broken, and the user asked not "
                  f"to try to fix it). Cannot search this document for the clause or quantities.")
            results.append({**t, "status": f"scanned_pdf_no_ocr ({page_count} pages, 0 chars)",
                            "clause": None, "clause_page": None, "qty_found": None,
                            "qty_page": None, "sha256": sha})
            time.sleep(WAIT_SECONDS)
            continue

        try:
            clause, clause_page, qty_found, qty_page = search_pdf(local_path)
        except Exception as e:
            print(f"  PDF TEXT EXTRACTION FAILED: {e}")
            results.append({**t, "status": "extraction_failed", "clause": None,
                            "clause_page": None, "qty_found": None, "qty_page": None, "sha256": sha})
            time.sleep(WAIT_SECONDS)
            continue

        if clause:
            print(f"  Debarment/blacklist clause FOUND on page {clause_page}:")
            print(f"    {clause[:200]}...")
        else:
            print("  No debarment/blacklist clause found (searched every page's extracted text).")
        print(f"  Estimated-quantity wording found: {qty_found}"
              + (f" (page {qty_page})" if qty_found else ""))

        results.append({**t, "status": "ok", "clause": clause, "clause_page": clause_page,
                        "qty_found": qty_found, "qty_page": qty_page, "sha256": sha})
        time.sleep(WAIT_SECONDS)

    # NIB-08/2024: not found - record as an explicit gap, not a guess.
    results.append({"tender": "NIB-08/2024", "rc_rows": 417, "slug": None, "pdf_path": None,
                    "notice_page": None, "status": "not_found_on_listing_page", "clause": None,
                    "clause_page": None, "qty_found": None, "qty_page": None, "sha256": None})

    header = ["Tender", "RC rows covered", "Status", "PDF URL", "SHA-256",
              "Debarment/blacklist clause (exact quote)", "Clause page",
              "Estimated-quantity wording found?", "Quantity page"]
    rows = []
    for r in results:
        pdf_url = (BASE + r["pdf_path"]) if r.get("pdf_path") else "Not found"
        rows.append([r["tender"], r["rc_rows"], r["status"], pdf_url, r["sha256"] or "",
                    r["clause"] or ("Not found" if r["status"] == "ok" else "N/A"),
                    r["clause_page"] or "", r["qty_found"], r["qty_page"] or ""])

    wb = Workbook()
    write_sheet(wb.active, header, rows, [16, 14, 24, 60, 68, 80, 12, 14, 12])
    wb.active.title = "Tender review"
    out = OUT / "rmsc_tender_review.xlsx"
    wb.save(out)
    print(f"\nSaved: {out}")
    print(f"Tenders checked: {len(results)}")
    print(f"  With a debarment/blacklist clause found: "
          f"{sum(1 for r in results if r['clause'])}")
    print(f"  With estimated-quantity wording found: "
          f"{sum(1 for r in results if r['qty_found'])}")


if __name__ == "__main__":
    main()
