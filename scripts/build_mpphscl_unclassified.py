"""
build_mpphscl_unclassified.py - one-off: for every MPPHSCL row still
"To check" (no specific clause number cited, so 07_mpphscl.py's own
CLAUSE_TEXT_OVERRIDES glossary never applied), download and OCR (hin+eng)
the linked order letter and try to classify it from the actual letter's
content, the same way Entries 13-16 did by hand for a handful of cases -
just at scale, for all ~164 distinct letters, not one at a time.

EVERY row this touches gets Status = "OCR - check" (rule F) - nothing
here is treated as confirmed. Classification is a best-effort keyword
read (reusing the same clause-citation patterns as 07_mpphscl.py's
glossary, PLUS broader Hindi/English quality/non-supply/fraud signal
words) - genuinely unclear letters are left "To check" rather than
forced into a category, and the user asked for a random 20-row sample
to check by hand, which this script also produces.

This does NOT edit master_list.xlsx directly (that file is regenerated
by 10_merge_master.py from the buyer files) - it writes its OWN
category/status back into 07_mpphscl.py's ENTRY_ID_OVERRIDES dict area
is NOT done automatically either, to avoid an unreviewed bulk edit to a
pipeline script; instead it writes output/mpphscl_unclassified_batch.xlsx
with a proposed category per row, and a SEPARATE override dict the user
(or a follow-up, explicitly-approved step) can paste into 07_mpphscl.py
once reviewed - consistent with "nothing is ever auto-merged".

Writes:
  data/raw/MPPHSCL/unclassified_pdfs/*.pdf  - raw evidence
  data/text/ocr/*.txt                       - OCR text, one per page (rule F)
  output/mpphscl_unclassified_batch.xlsx    - every row's proposed category,
                                              "OCR - check" status, random-20 sample sheet
"""

import re
import ssl
import time
import hashlib
import csv
import gc
import random
from datetime import datetime
from pathlib import Path

import requests
import pdfplumber
from requests.adapters import HTTPAdapter
from openpyxl import Workbook, load_workbook

from common import OUT, RAW, OCR_TEXT, write_sheet, ocr_page_text, save_ocr_text, QUALITY_WORDS, SUPPLY_WORDS

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
MAX_PAGES = 3   # order letters are short - cap so one long PDF can't run away

EVIDENCE_DIR = RAW / "MPPHSCL" / "unclassified_pdfs"
EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)

# Checkpoint - written one row at a time as soon as a row is classified,
# not only at the very end. Found necessary 2026-10-31: the first run of
# this script was killed mid-batch by the OS for low system memory, and
# since the old version only wrote its xlsx once at the very end, all 21
# rows already processed (downloaded + OCR'd, evidence safe on disk) were
# lost from the *results* side purely because nothing had been persisted
# yet. This checkpoint is the fix - PROJECT_RULES.md's recoverability priority,
# not a style preference. On restart, rows already in the checkpoint are
# skipped entirely (no re-download, no re-OCR).
CHECKPOINT_PATH = OUT / "mpphscl_unclassified_checkpoint.csv"
CHECKPOINT_HEADER = ["Entry ID", "Firm", "Source link", "Proposed category (OCR - check)",
                     "Matched on", "Note"]


def load_checkpoint():
    done = {}
    if CHECKPOINT_PATH.exists():
        with open(CHECKPOINT_PATH, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                done[row["Entry ID"]] = [row[h] for h in CHECKPOINT_HEADER]
    return done


def append_checkpoint(row):
    is_new = not CHECKPOINT_PATH.exists()
    with open(CHECKPOINT_PATH, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if is_new:
            w.writerow(CHECKPOINT_HEADER)
        w.writerow(row)
        f.flush()


class LegacyRenegotiationAdapter(HTTPAdapter):
    """MPPHSCL needs this TLS quirk to connect at all (see 07_mpphscl.py) -
    certificate verification stays ON (ssl.create_default_context() keeps
    verify_mode=CERT_REQUIRED by default). Missed in the first version of
    this script - caught by the small-test-first run before the full
    164-row batch, not after."""
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.options |= ssl.OP_LEGACY_SERVER_CONNECT
        kwargs["ssl_context"] = ctx
        super().init_poolmanager(*args, **kwargs)


session = requests.Session()
session.headers.update(HEADERS)
session.mount("https://", LegacyRenegotiationAdapter())

# Same clause-citation patterns as 07_mpphscl.py's CLAUSE_TEXT_OVERRIDES
# (kept in sync by hand, not imported, since 07_mpphscl.py's own copy
# also carries ENTRY_ID_OVERRIDES this script does not need). Expanded
# 2026-11-01 after the user's "Random-20 check" (document_checks_
# 2026-11-01.xlsx) found the computer classified only 3 of 20 correctly -
# 9 of the 20 misses were "agreement not executed" cases citing clause
# 8(c) or an Annexure V/IX withdrawal, neither previously recognized.
CLAUSE_PATTERNS = [
    (re.compile(r"annexure\s*ix\s*clause\s*no\.?\s*7\b", re.I), "Quality"),
    (re.compile(r"annexure\s*ix\s*clause\s*no\.?\s*12\b", re.I), "Non-supply"),
    (re.compile(r"(?:t-213\s*)?(?:clause\s*(?:no\.?\s*)?)?10\s*\.?\s*6\b", re.I), "Other"),
    (re.compile(r"(?:t-213\s*)?(?:clause\s*(?:no\.?\s*)?)?12\s*\.?\s*5\b", re.I), "Other"),
    (re.compile(r"clause\s*(?:no\.?\s*)?8\s*\(\s*c\s*\)", re.I), "Other"),
    (re.compile(r"annexure[\s-]*v\b.{0,60}withdraw", re.I | re.S), "Other"),
    (re.compile(r"annexure[\s-]*ix\b.{0,60}withdraw", re.I | re.S), "Other"),
]

# Hindi signal words, found while hand-reading letters in Entries 13-16 -
# a translation aid, not an authoritative dictionary. Kept separate from
# QUALITY_WORDS/SUPPLY_WORDS (common.py) since those are English-only by
# design; mixing languages into one shared list would be confusing there.
HINDI_QUALITY_WORDS = ["अमानक", "गुणवत्ता", "nosq"]
HINDI_SUPPLY_WORDS = ["सप्लाई नहीं", "सप्लाई की पालना नहीं", "विलम्ब", "आपूर्ति नहीं", "समय-सीमा में पूर्ण नहीं"]
FRAUD_WORDS = ["forged", "edited/forged", "जाली", "फर्जी", "fabricated", "misleading", "wrong information",
              "गलत सूचना", "collusion", "coercive"]

# Added 2026-11-01 per the Random-20 review - a letter saying the
# "agreement" (security deposit / contract) was never executed, in prose,
# with no clause number cited at all (the ones ABOVE still catch a bare
# clause-number citation; this catches the same real-world case when the
# letter only describes it in words). The most common single pattern in
# the sample (9 of 20 rows) - still "Other" (a contract-execution
# failure, never a drug-quality finding), same bucket clause 12.5/8(c)
# already map to.
AGREEMENT_NOT_EXECUTED_WORDS = ["agreement not executed", "fail to execute the agreement",
                                "fails to execute the agreement", "failed to execute the agreement",
                                "failure to execute the agreement", "non-execution of agreement",
                                "non execution of the agreement", "did not execute the agreement",
                                "withdrawal of tender", "withdrawal from the tender",
                                "withdrawn the tender", "emd forfeited", "security deposit forfeited"]

# Hospital furniture/equipment found NOT per specification (clause 10.9,
# one of its 6 possible grounds per the existing comment above
# CLAUSE_PATTERNS' module - NOT safe to override on the bare clause
# number alone, so this instead requires an EQUIPMENT noun actually
# present in the text, not just a clause citation). Added 2026-11-01 -
# 3 of 20 Random-20 rows were this exact case, misread as "Quality" by
# the old classifier because MPPHSCL's own live-page category label
# ("Quality") is itself wrong for non-drug goods.
FURNITURE_EQUIPMENT_WORDS = ["bedside locker", "examination table", "hospital furniture",
                            "operation table", "operating table", "instrument trolley",
                            "stretcher trolley", "patient trolley", "wheelchair", "hospital bed",
                            "surgical table"]
NOT_PER_SPEC_WORDS = ["not as per specification", "not as per the specification",
                      "not according to specification", "non-conforming", "non conforming"]

# Firm hid/did not declare that ANOTHER buyer had already blacklisted it -
# added 2026-11-01 (the Radiant Parenterals case, Random-20 row 4 - the
# old classifier read this as "Quality" because the letter mentions a
# quality blacklisting, but the GROUND for THIS buyer's own action is the
# concealment, not a quality test of its own).
CONCEALMENT_WORDS = ["conceal", "suppress", "non-disclosure", "non disclosure", "did not declare",
                     "not declared", "failed to declare", "false affidavit", "withheld information"]
CONCEALMENT_CONTEXT_WORDS = ["blacklist", "debar", "banned", "another state", "other state"]


def classify_text(text):
    """Returns (category, matched_on) - "To check"/"" if nothing
    recognized (never forced)."""
    lower = text.lower()
    for pattern, category in CLAUSE_PATTERNS:
        if pattern.search(lower):
            return category, f"clause citation matched"
    if any(w in lower for w in AGREEMENT_NOT_EXECUTED_WORDS):
        return "Other", "agreement not executed / tender withdrawal keyword found"
    if (any(w in lower for w in CONCEALMENT_WORDS)
            and any(w in lower for w in CONCEALMENT_CONTEXT_WORDS)):
        return "Other", "concealment of another buyer's ban keyword found"
    if (any(w in lower for w in FURNITURE_EQUIPMENT_WORDS)
            and any(w in lower for w in NOT_PER_SPEC_WORDS)):
        return "Other", "furniture/equipment not per specification keyword found"
    if any(w in text for w in HINDI_QUALITY_WORDS) or any(w in lower for w in QUALITY_WORDS):
        return "Quality", "quality keyword found"
    if any(w in text for w in HINDI_SUPPLY_WORDS) or any(w in lower for w in SUPPLY_WORDS):
        return "Non-supply", "supply-delay keyword found"
    if any(w in lower for w in FRAUD_WORDS) or any(w in text for w in FRAUD_WORDS):
        return "Other", "fraud/misrepresentation keyword found"
    return "To check", "no recognized signal"


def robots_allows():
    # Already confirmed project-wide (PROGRESS_LOG.md, multiple entries):
    # mpphscl.mp.gov.in has no robots.txt (404) - no restriction. Re-checked
    # once per run anyway, not assumed silently.
    try:
        r = session.get("https://mpphscl.mp.gov.in/robots.txt", timeout=20)
        return r.status_code != 200 or not r.text.strip()
    except requests.RequestException:
        return True


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def log_download(file_path, sha256, url):
    log_path = OUT / "download_log.csv"
    is_new = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
        if is_new:
            w.writeheader()
        w.writerow({"date_time": datetime.now().strftime("%d-%m-%Y %H:%M"), "status": "captured",
                    "file": str(file_path), "sha256": sha256, "url": url})


def reclassify_from_cache():
    """Re-runs classify_text() (with the expanded signal words above)
    against every row already in CHECKPOINT_PATH, WITHOUT any network
    call and WITHOUT re-running OCR - added 2026-11-01 so an improved
    classifier can be applied to the existing 164-row batch without
    redoing the heaviest, slowest part of this script. Re-reads native
    PDF text fresh (cheap, no OCR engine involved) for every row whose
    local PDF is still on disk; for a row that needed OCR the first time,
    reads the OCR text ALREADY SAVED at data/text/ocr/<file>_p<n>.txt
    (PROJECT_RULES.md rule F - that text was preserved precisely so it would not
    need to be regenerated) instead of calling Tesseract again. A row
    whose PDF/OCR text is missing keeps its PREVIOUS checkpoint category
    unchanged (never silently downgraded to "To check") and is listed as
    skipped."""
    done = load_checkpoint()
    if not done:
        print(f"No checkpoint found at {CHECKPOINT_PATH} - nothing to reclassify. Run the normal "
              f"(downloading) pass first.")
        return

    out_rows, changed, unchanged, skipped_no_text = [], 0, 0, 0
    for entry_id, row in done.items():
        firm, link, old_category = row[1], row[2], row[3]
        safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", entry_id)[:100] + ".pdf"
        target = EVIDENCE_DIR / safe_name

        text = ""
        if target.exists() and target.stat().st_size > 0:
            try:
                with pdfplumber.open(target) as pdf:
                    native = "\n".join((p.extract_text() or "") for p in pdf.pages[:MAX_PAGES])
            except Exception:
                native = ""
            if len(native.strip()) >= 150:
                text = native
            else:
                ocr_parts = []
                for pg_i in range(1, MAX_PAGES + 1):
                    ocr_file = OCR_TEXT / f"{target.stem}_p{pg_i}.txt"
                    if ocr_file.exists():
                        ocr_parts.append(ocr_file.read_text(encoding="utf-8"))
                text = "\n".join(ocr_parts)

        if not text.strip():
            out_rows.append(row)                   # keep the previous result unchanged
            skipped_no_text += 1
            continue

        category, matched_on = classify_text(text)
        if category == old_category:
            out_rows.append(row)
            unchanged += 1
        else:
            out_rows.append([entry_id, firm, link, category, matched_on,
                             f"Reclassified 2026-11-01 with the expanded signal words (was "
                             f"{old_category!r}) - {row[5]}"])
            changed += 1

    header = ["Entry ID", "Firm", "Source link", "Proposed category (OCR - check)",
             "Matched on", "Note"]
    wb = Workbook()
    write_sheet(wb.active, header, out_rows, [26, 34, 70, 18, 24, 50])
    wb.active.title = "All processed"
    sample = random.sample(out_rows, min(20, len(out_rows)))
    write_sheet(wb.create_sheet("Random 20 sample"), header, sample, [26, 34, 70, 18, 24, 50])
    out = OUT / "mpphscl_unclassified_batch.xlsx"
    wb.save(out)

    from collections import Counter
    counts = Counter(r[3] for r in out_rows)
    print(f"Reclassified {len(out_rows)} rows from cached text (0 downloads, 0 new OCR calls).")
    print(f"  Category changed from the expanded classifier : {changed}")
    print(f"  Category unchanged                             : {unchanged}")
    print(f"  Skipped (no cached PDF/OCR text found)         : {skipped_no_text}")
    print(f"  Proposed categories (new totals): {dict(counts)}")
    print(f"Saved: {out}")


def main():
    import sys
    if "--reclassify" in sys.argv:
        reclassify_from_cache()
        return

    if not robots_allows():
        print("BLOCKED by robots.txt - stopping.")
        return

    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_ws = master_wb["Master list"]
    master_rows = list(master_ws.iter_rows(values_only=True))
    m_header, m_data = master_rows[0], master_rows[1:]
    m_idx = {h: i for i, h in enumerate(m_header)}

    targets = [r for r in m_data if r[m_idx["Buyer"]] == "MPPHSCL"
              and r[m_idx["Reason category (Quality / Non-supply / Other / To check)"]] == "To check"]

    done = load_checkpoint()
    remaining = [r for r in targets if r[m_idx["Entry ID"]] not in done]
    print(f"Found {len(targets)} unclassified MPPHSCL rows total - "
          f"{len(done)} already checkpointed from a prior run, {len(remaining)} left to do.",
          flush=True)

    out_rows = list(done.values())
    for i, r in enumerate(remaining, start=1):
        entry_id = r[m_idx["Entry ID"]]
        firm = r[m_idx["Firm name (standard)"]]
        link = r[m_idx["Source link"]]
        print(f"[{i}/{len(remaining)}] {entry_id} ({firm})", flush=True)

        if not link or not link.lower().endswith(".pdf"):
            row = [entry_id, firm, link, "To check", "no signal", "No PDF link available"]
            append_checkpoint(row)
            out_rows.append(row)
            continue

        safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", entry_id)[:100] + ".pdf"
        target = EVIDENCE_DIR / safe_name
        try:
            if not target.exists() or target.stat().st_size == 0:
                resp = session.get(link, timeout=TIMEOUT)
                resp.raise_for_status()
                if not resp.content.startswith(b"%PDF"):
                    row = [entry_id, firm, link, "To check", "no signal",
                           "Downloaded content is not a valid PDF"]
                    append_checkpoint(row)
                    out_rows.append(row)
                    time.sleep(WAIT_SECONDS)
                    continue
                target.write_bytes(resp.content)
                sha = sha256_bytes(resp.content)
                log_download(target, sha, link)
                time.sleep(WAIT_SECONDS)
        except requests.RequestException as e:
            row = [entry_id, firm, link, "To check", "no signal", f"Download failed: {e}"]
            append_checkpoint(row)
            out_rows.append(row)
            continue

        # A SHORT native-text layer (e.g. just a stamped page number, like
        # "1373158 0397/26 02 09") is not the same as a readable letter -
        # found via the small-test-first run, not assumed. Require a
        # minimum length before trusting native text over OCR.
        MIN_NATIVE_CHARS = 150
        try:
            with pdfplumber.open(target) as pdf:
                pages = pdf.pages[:MAX_PAGES]
                native = "\n".join((p.extract_text() or "") for p in pages)
                if len(native.strip()) >= MIN_NATIVE_CHARS:
                    text, used_ocr = native, False
                else:
                    parts = []
                    for pg_i, page in enumerate(pages, start=1):
                        ocr_text = ocr_page_text(page, lang="hin+eng")
                        save_ocr_text(target.name, pg_i, ocr_text)
                        parts.append(ocr_text)
                    text, used_ocr = "\n".join(parts), True
        except Exception as e:
            row = [entry_id, firm, link, "To check", "no signal", f"Could not read PDF: {e}"]
            append_checkpoint(row)
            out_rows.append(row)
            continue

        category, matched_on = classify_text(text)
        note = f"{'OCR hin+eng' if used_ocr else 'native text'}; {matched_on}"
        row = [entry_id, firm, link, category, matched_on, note]
        append_checkpoint(row)
        out_rows.append(row)

        # OCR (page.to_image at 300dpi, hin+eng) is the heaviest step here -
        # nudge garbage collection periodically so freed PDF/image memory is
        # actually reclaimed promptly rather than lingering, after the first
        # run was killed by the OS for low system memory mid-batch.
        if i % 5 == 0:
            gc.collect()

    header = ["Entry ID", "Firm", "Source link", "Proposed category (OCR - check)",
             "Matched on", "Note"]
    wb = Workbook()
    write_sheet(wb.active, header, out_rows, [26, 34, 70, 18, 24, 50])
    wb.active.title = "All processed"

    sample = random.sample(out_rows, min(20, len(out_rows)))
    write_sheet(wb.create_sheet("Random 20 sample"), header, sample, [26, 34, 70, 18, 24, 50])

    out = OUT / "mpphscl_unclassified_batch.xlsx"
    wb.save(out)

    from collections import Counter
    counts = Counter(r[3] for r in out_rows)
    print(f"\nSaved: {out}")
    print(f"  Rows processed: {len(out_rows)}")
    print(f"  Proposed categories: {dict(counts)}")
    print(f"  Random 20 sample included for hand-checking.")


if __name__ == "__main__":
    main()
