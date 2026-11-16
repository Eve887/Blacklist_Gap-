"""
build_pdfs_to_read.py - reading list of the order PDFs behind every
candidate gap, for manual checking.

Not part of the numbered pipeline (not auto-chained after
10_merge_master.py) - run it by hand whenever you want a fresh reading
list after candidate_gaps.xlsx changes.

Reads:
  output/candidate_gaps.xlsx   (made by 12_candidate_gaps.py)
  output/master_list.xlsx      (for each entry's Source link)
Downloads:
  Any order PDF not already saved locally, into data/raw/<BUYER>/ - same
  download safety pattern as the buyer scripts (robots.txt check, 3-second
  pause, SHA-256 not needed here since each buyer script already logs its
  own downloads; this just needs a local copy to point you at).
Writes:
  output/pdfs_to_read.xlsx  - one row per distinct entry (PDF) behind a
  candidate gap, Zee Laboratories / Nestor / ANG first (as asked), then
  everything else. Plus Medipol's two RMSC order letters (EXTRA_ENTRY_IDS)
  even though Medipol isn't in candidate_gaps.xlsx (its two order letters
  use different standardized names, never auto-merged) - added because
  the user asked for them specifically.

As of 2026-10-31, OCR IS used (Tesseract reinstalled - see
PROGRESS_LOG.md) for RMSC rows whose PDF is a scanned image. RMSC's
order letters turned out to be written in HINDI - English-only OCR
could only make out isolated Latin-script fragments, not which field
each belonged to (see Entry 13). Hindi language support
(tessdata_fast/hin.traineddata, official repo, hash-verified) was added
2026-10-31 and the 7 RMSC letters were re-OCR'd with `-l hin+eng`
(`lang="hin+eng"` below) - MUCH more readable, but still OCR, so
PROJECT_RULES.md rule F still applies: nothing from it counts as confirmed
until the user checks it against the PDF.

Batch number / Manufacturing date / Supply date for RMSC rows now come
from RMSC_HINDI_READINGS below - a hand-curated dict, NOT a regex
parser. Reading Hindi-English mixed OCR text and deciding which number
is a batch vs a purchase-order vs a notice date is a semantic judgment
call this script cannot make reliably on its own (and should not
pretend to) - so these values were read and translated by hand (one
pass, 2026-10-31), each with its own uncertainty notes for digits OCR
plausibly misread (1/7, 4/7 and similar look-alikes in this scan
quality), and are reproduced here as data, not derived by code. A
genuine "Manufacturing date" field does not appear in ANY of the 7
letters (not an OCR gap - these debarment orders cite tender dates,
purchase-order dates, and disciplinary-committee dates, never a
drug-batch manufacturing date) - that field stays honestly "Not stated
in letter" for all 7, not guessed. Each translation is marked "machine
translation - check" per the user's instruction.

Non-RMSC rows are untouched by any of this (still fill-in-by-hand).
"""

import ssl
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlparse
from urllib import robotparser

import requests
import pdfplumber
from requests.adapters import HTTPAdapter
from openpyxl import Workbook, load_workbook

from common import OUT, RAW, write_sheet, ocr_page_text, save_ocr_text

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")

PRIORITY_FIRMS = ["ZEE LABORATORIES LTD", "NESTOR PHARMACEUTICALS LTD", "ANG LIFE SCIENCES INDIA LTD"]

# Medipol is NOT in candidate_gaps.xlsx - its two RMSC order letters use
# different standardized names ("Medipol Pharmaceutica" vs "Medipol
# Pharmaceutical India Pvt Ltd", never auto-merged, per the project's
# standing rule) - but the user asked for these specifically ("Medipol
# first"), so they're added here explicitly rather than pulled from
# candidate_gaps.xlsx like everything else.
EXTRA_ENTRY_IDS = ["RMSC-DEBAR-3085-20-02-2025", "RMSC-DEBAR-554-5-6-2026"]

MAX_OCR_PAGES = 4   # order letters are short - cap so a long PDF can't run away

# Hand-read and translated 2026-10-31 from the hin+eng OCR text (NOT a
# regex parser's output - see the module docstring). Every digit the OCR
# rendered ambiguously (visually similar Devanagari/Latin shapes in this
# scan quality - 1/7/4 confusions were the most common) is flagged in
# "uncertain" rather than silently picked one way.
#
# CORRECTED 2026-11-01 (document_checks_2026-11-01.xlsx, "Hindi
# letter corrections" sheet): the user re-read 4 of these 6 letters
# directly from the page IMAGES (not OCR) and found wrong numbers in
# batch/purchase-order/tender-date fields on all 4 - every correction
# below is marked with the date it was fixed; fields marked "Correct" in
# that sheet (554, 3085) are left untouched. The ANG order-number
# discrepancy (the letter itself is numbered 3099, not 3078) is NOT
# resolved here - logged to output/source_issues.csv instead, per the
# "log source problems instead of fixing them silently" rule, since
# which number is authoritative is still an open question.
RMSC_HINDI_READINGS = {
    "RMSC-DEBAR-1846-10-11-2025-2472-16-01-2026": {   # Zee Laboratories
        "batch": "1722-285 (confirmed 2026-11-01 by the user reading the page image directly - "
                 "corrects the earlier OCR-uncertain '722-285'/'i722.285').",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase order 1654, dated 14.06.2022 (confirmed 2026-11-01 - corrects "
                       "the earlier OCR-uncertain '654'/'74.06.2022')",
        "translation": "Tender F.02(345)/RMSCL/PROCUREMENT/DRUG/NIB-02/2022/62 dt 07.01.2022 "
                       "(confirmed 2026-11-01 - corrects the earlier OCR-misread tender date "
                       "'07.07.2022') for Itraconazole Cap 100mg[118]. Supplied batch failed a "
                       "Dissolution test at the State Drug Testing Laboratory; firm's objection "
                       "sample sent to the Central Drugs Laboratory, Kolkata, which also found "
                       "it Not of Standard Quality (Grossly Substandard). Disciplinary "
                       "committee met 28.10.2025; notice to firm dated approx. 17.10.2025 "
                       "(uncertain digit, not re-checked); no firm representative attended.",
    },
    "RMSC-DEBAR-3078-20-02-2025-b": {   # ANG Life Sciences
        "batch": "I472006, I472007 (confirmed 2026-11-01 - corrects the earlier OCR reading "
                 "that dropped the leading 'I').",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase order 3917, dated 11.10.2022 (date confirmed 2026-11-01 - "
                       "corrects the earlier 'date partly illegible'); and 2793, dated "
                       "23.08.2022 (unchanged - already correct).",
        "translation": "Tender F.02(332)/RMSCL/Proc./Drug/NIB-22/2021/1380 dt 04-08-2021 "
                       "(confirmed 2026-11-01 - corrects the earlier OCR-misread tender number "
                       "'380') for Dobutamine Inj IP 50mg/ml/250mg (vial)[192] (confirmed "
                       "2026-11-01 - corrects the earlier OCR-misread drug code '492'). Both "
                       "batches failed a 'Description' check (not a potency/assay failure) at "
                       "RMSC's registered lab, confirmed by the State Drug Testing Laboratory. "
                       "Disciplinary committee met approx. 22.01.2025 (partly illegible date, "
                       "not re-checked). Debarred for 1 year from the order date. NOTE: the "
                       "order letter itself is numbered 3099 dt 21-02-2025, not 3078 - this "
                       "project's Entry ID uses the reference ('3078') shown on RMSC's own live "
                       "page; which number is authoritative is an open question, logged to "
                       "output/source_issues.csv, not resolved here.",
    },
    "RMSC-DEBAR-1837-12-08-2025": {   # Ultra Drugs
        "batch": "Not applicable - see translation",
        "mfg_date": "Not applicable",
        "supply_date": "Not applicable",
        "translation": "NOT a quality/batch case. The firm quoted wrong rates (Rs 10.00 vs a "
                       "mistakenly entered 0.1, and Rs 56.00 vs 1.56) for 2 items in a bid, "
                       "then tried to correct them AFTER the financial bid was opened "
                       "(letter dated 17.07.2024) - ruled a violation of the no-modification-"
                       "after-opening bid condition, not a drug-quality failure.",
    },
    "RMSC-DEBAR-3093-20-02-2025": {   # Unicure India
        "batch": "GE1TC005, URDT0327, URDT0328 initially tested; URDT0327 and URDT0328 "
                 "confirmed failing on retest by the government lab (confirmed 2026-11-01 - "
                 "corrects the earlier OCR-ambiguous 'URDT0005' for the first batch).",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase orders 1404, dated 01.06.2022; 4644, dated 16.07.2022; 190, "
                       "dated 21.12.2022; 1171, dated 17.05.2023; 5190, dated 21.12.2022 (all "
                       "confirmed 2026-11-01 - corrects several earlier OCR-uncertain numbers/"
                       "dates: 1104->1404, 1611->4644, 490(?)->190, 5490(?)->5190, and both "
                       "'27.12.2022'/'24.12.2022' dates -> 21.12.2022).",
        "translation": "Tender F.02(330)/RMSCL/procurement/drug/NIB-20/2021/1266 (16.07.2021) "
                       "for Glipizide And Metformin Hydrochloride Tablets USP [452]. Batches "
                       "failed a Dissolution test at the State Drug Testing Laboratory. "
                       "Disciplinary committee met 22.04.2025. Debarred for 2 years.",
    },
    "RMSC-DEBAR-1453-24-08-2026": {   # Vee Excel
        "batch": "C260005, C260006",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase order 4988, dated 25.11.2025 (confirmed 2026-11-01 - corrects "
                       "the earlier OCR-misread '25.01.2025'); and 6034, dated 24.12.2025.",
        "translation": "Tender F.02(400)/RMSCL/PROCUREMENT/DRUG/NIB-05/2024/546 (13.03.2024) "
                       "for Cefpodoxime Proxetil Oral Suspension [NRD-485]. Both batches failed "
                       "a 'Description' check (Minor defects) at the State Drug Testing "
                       "Laboratory. Notice to firm 24.07.2026; disciplinary committee met "
                       "30.07.2026, firm's representative heard but presented no new facts.",
    },
    "RMSC-DEBAR-3085-20-02-2025": {   # Medipol Pharmaceutica
        # Checked against the page image 2026-11-01 (document_checks_2026-11-01.xlsx,
        # "Hindi letter corrections" sheet) - every field below confirmed CORRECT, no change.
        "batch": "13 batches cited for Assay-of-Calcium-Pantothenate failure: TMVN-143, 154, "
                 "157, 158, 161, 162, 163, 168, 169, 170, 171, 172, 173 (OCR renders one as "
                 "'TMVN-243' once vs 'TMVN-143' twice elsewhere - likely the same batch, "
                 "digit uncertain). TMVN-163 specifically also failed at the Central lab.",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase orders 4055 (03.02.2022) and 4621 (16.03.2022), under tender "
                       "F.02(311)/RMSCL/Proc./Drug/NIB-02/2021/183 (12.02.2021)",
        "translation": "For Multivitamin Tablets NFI Formula [394]. 13 batches failed an Assay "
                       "of Calcium Pantothenate test at RMSC's registered lab; retested at the "
                       "State Drug Testing Laboratory with the same result; the Central Drugs "
                       "Laboratory separately confirmed batch TMVN-163 failing too. "
                       "Disciplinary committee met 22.07.2025.",
    },
    "RMSC-DEBAR-554-5-6-2026": {   # Medipol Pharmaceutical India Pvt Ltd
        # Checked against the page image 2026-11-01 (document_checks_2026-11-01.xlsx,
        # "Hindi letter corrections" sheet) - every field below confirmed CORRECT, no change.
        "batch": "TACR-010 (consistent across all mentions - high confidence)",
        "mfg_date": "Not stated in letter",
        "supply_date": "Purchase order 4146, dated 30.09.2025",
        "translation": "Tender F.02(392)/RMSCL/PROCUREMENT/Drug/NIB-01/2024/96 (12.01.2024) "
                       "for Acyclovir Tab IP 200mg[63]. Batch failed a Dissolution test at "
                       "RMSC's own lab, confirmed by the State Drug Testing Laboratory. Notice "
                       "to firm 08.05.2026; disciplinary committee met 13.05.2026, no firm "
                       "representative attended.",
    },
}


def extract_or_ocr(local_path, entry_id):
    """Returns (text, status, pages_used). Uses native PDF text when
    available; OCRs with hin+eng (and saves via save_ocr_text, rule F)
    when the PDF is a scanned image - RMSC's letters are Hindi (see
    module docstring). Does NOT guess which value is which field -
    that's RMSC_HINDI_READINGS' job (hand-curated), not this function's."""
    if not local_path:
        return "", "", []
    path = Path(local_path)
    if not path.exists():
        return "", "", []
    try:
        with pdfplumber.open(path) as pdf:
            native = "\n".join((p.extract_text() or "") for p in pdf.pages)
            if native.strip():
                return native, "", []
            pages_used, parts = [], []
            for i, page in enumerate(pdf.pages[:MAX_OCR_PAGES], start=1):
                text = ocr_page_text(page, lang="hin+eng")
                save_ocr_text(path.stem + "_hineng" + path.suffix, i, text)
                parts.append(text)
                pages_used.append(i)
            return "\n".join(parts), "OCR - check", pages_used
    except Exception as e:
        return f"[could not read: {e}]", "", []


class LegacyRenegotiationAdapter(HTTPAdapter):
    """MPPHSCL needs this TLS quirk to connect at all (see 07_mpphscl.py) -
    certificate verification stays ON (ssl.create_default_context() keeps
    verify_mode=CERT_REQUIRED by default)."""
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
            rp.parse(r.text.splitlines() if r.status_code == 200 else [])
        except requests.RequestException:
            rp.parse([])
        _robots_cache[root] = rp
    return _robots_cache[root].can_fetch(HEADERS["User-Agent"], url)


def download(url, buyer, entry_id):
    if not url or not url.lower().startswith("http") or not url.lower().endswith(".pdf"):
        return ""                                           # not a direct PDF link - nothing to save
    safe_name = "".join(c if c.isalnum() or c in "-_." else "_" for c in entry_id)[:100] + ".pdf"
    target = RAW / buyer / "candidate_gap_pdfs" / safe_name
    if target.exists() and target.stat().st_size > 0:
        return str(target)
    if not robots_allows(url):
        print(f"  BLOCKED by robots.txt -> save by hand: {target}  link: {url}")
        return ""
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        r = session.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        if not r.content.startswith(b"%PDF"):
            return ""
        target.write_bytes(r.content)
        time.sleep(WAIT_SECONDS)
        return str(target)
    except requests.RequestException as e:
        print(f"  failed to download {url}: {e}")
        return ""


def main():
    cand_wb = load_workbook(OUT / "candidate_gaps.xlsx", read_only=True)
    cand_rows = list(cand_wb["Candidate gaps"].iter_rows(values_only=True))[1:]

    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_rows = list(master_wb["Master list"].iter_rows(values_only=True))
    header = master_rows[0]
    by_id = {r[0]: r for r in master_rows[1:]}
    firm_idx = header.index("Firm name (as written)")
    buyer_idx = header.index("Buyer")
    link_idx = header.index("Source link")

    # collect every distinct (firm, entry_id) pair referenced as A or B
    seen = {}
    for row in cand_rows:
        firm, a_buyer, a_id, b_buyer, b_id = row[0], row[1], row[2], row[6], row[7]
        for eid in (a_id, b_id):
            if eid not in seen:
                seen[eid] = firm
    for eid in EXTRA_ENTRY_IDS:          # Medipol - see EXTRA_ENTRY_IDS comment above
        if eid not in seen and eid in by_id:
            seen[eid] = by_id[eid][firm_idx]

    out_rows = []
    for entry_id, firm in seen.items():
        r = by_id.get(entry_id)
        if r is None:
            continue
        link = r[link_idx]
        buyer = r[buyer_idx]
        local_path = download(link, buyer, entry_id) if link else ""

        batch, mfg_date, supply_date, status, letter_note, pages_used = "", "", "", "", "", []
        if buyer == "RMSC" and local_path:
            text, status, pages_used = extract_or_ocr(local_path, entry_id)
            if status == "OCR - check":
                reading = RMSC_HINDI_READINGS.get(entry_id)
                if reading:
                    batch = reading["batch"] + " [OCR - check]"
                    mfg_date = reading["mfg_date"]
                    supply_date = reading["supply_date"] + " [OCR - check]" \
                                  if reading["supply_date"] not in ("Not applicable",) else reading["supply_date"]
                    letter_note = reading["translation"] + " [machine translation - check]"
                else:
                    # No hand-curated reading exists for this entry yet (a
                    # future candidate gap not covered by the 2026-10-31
                    # reading pass) - left blank rather than guessed.
                    letter_note = (f"Hindi-language scanned letter, not yet hand-read/"
                                   f"translated. Raw hin+eng OCR text saved for pages "
                                   f"{pages_used} in data/text/ocr/{Path(local_path).stem}_"
                                   f"hineng_p<n>.txt.")
            elif text.strip():
                letter_note = text.strip()[:300].replace("\n", " ")

        out_rows.append([firm, buyer, entry_id, local_path, link,
                         batch, mfg_date, supply_date, letter_note, status,
                         ",".join(str(p) for p in pages_used)])

    priority_rank = {f: i for i, f in enumerate(PRIORITY_FIRMS)}
    out_rows.sort(key=lambda r: (priority_rank.get(r[0], len(PRIORITY_FIRMS)), r[0], r[1], r[2]))

    header_out = ["Firm", "Buyer", "Entry ID", "Local file path", "Link",
                  "Batch number", "Manufacturing date", "Supply/sample date", "What the letter says",
                  "Status", "Page / row in source"]

    wb = Workbook()
    write_sheet(wb.active, header_out, out_rows, [30, 10, 24, 50, 50, 50, 24, 60, 90, 16, 18])
    wb.active.title = "PDFs to read"
    out = OUT / "pdfs_to_read.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Rows (distinct entries) : {len(out_rows)}")
    print(f"  RMSC rows OCR'd (hin+eng): {sum(1 for r in out_rows if r[9] == 'OCR - check')}")
    print(f"    - with a hand-read/translated reading: {sum(1 for r in out_rows if r[2] in RMSC_HINDI_READINGS)}")
    print(f"  Downloaded locally      : {sum(1 for r in out_rows if r[3])}")


if __name__ == "__main__":
    main()
