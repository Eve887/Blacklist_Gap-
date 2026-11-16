"""
10_merge_master.py - combine every buyer file into the final MASTER LIST.

Reads:
  output/buyers/*.xlsx   - one file per buyer (AP.xlsx, AS.xlsx, PMBI.xlsx, ...),
                           each written by its own buyer script, each using
                           the same 19 master-list columns
Writes:
  output/master_list.xlsx  with four sheets:
    - Master list           : every buyer's rows, combined
    - Matches               : firms that appear more than once - now checked
                               ACROSS buyers too, not just within one
    - Skipped rows          : combined from every buyer that has its own
                               Skipped rows sheet, with a Buyer column added
    - Unmatched revocations : combined from every buyer that has its own
                               Unmatched revocations sheet

This script does no extraction or cleaning itself - it only combines what
the buyer scripts already built. Running it again from scratch (after
re-running the buyer scripts) must give the same result: buyer files are
read in a fixed (sorted by filename) order, and nothing here depends on
timing or on what ran before.

Classification (reason_category(), the MPPHSCL/RMSC hand-verified
overrides, and quality_level()) ALSO happens here, once, centrally -
moved out of the individual buyer scripts 2026-11-01 (PROGRESS_LOG.md
Entry 19, per checklist item 13). Every buyer script writes
common.CATEGORY_PENDING into the "Reason category" column instead of
computing it - this script overwrites that placeholder on every row, on
every merge, so a QUALITY_WORDS/SUPPLY_WORDS change (or an override table
edit) takes effect everywhere the next time this script runs, with no
buyer script needing a separate re-run. (The Meril/GMSCL bug - Entries
16-17 - happened exactly because a buyer script HADN'T been re-run after
a shared word list changed; this removes that failure mode entirely.)

After saving master_list.xlsx, this script ALSO runs 11_cross_buyer_pairs.py
and 12_candidate_gaps.py automatically, in that order, so both always
reflect the latest set of buyers without a separate manual step. (Imported
by file path, not a normal "import 11_..." statement - their filenames
start with a digit, which Python's import statement does not allow.)
"""

import re
import importlib.util
from pathlib import Path

from openpyxl import Workbook, load_workbook

from common import OUT, BUYERS, COLUMNS, find_matches, write_sheet, quality_level, reason_category

ENTRY_ID_IDX = COLUMNS.index("Entry ID")
BUYER_IDX = COLUMNS.index("Buyer")
REASON_IDX = COLUMNS.index("Reason (as written)")
CATEGORY_IDX = COLUMNS.index("Reason category (Quality / Non-supply / Other / To check)")
LAB_IDX = COLUMNS.index("Lab that declared NSQ")
STATUS_IDX = COLUMNS.index("Status")


# ------------------------------------------------------------------
# CENTRALIZED CLASSIFICATION (moved from 07_mpphscl.py and 08_rmsc.py,
# 2026-11-01 - PROGRESS_LOG.md Entry 19)
# ------------------------------------------------------------------

# MPPHSCL's live blacklist page often cites only a bare clause number
# ("TENDER TERMS AND CONDITON 12.5", "Annexure IX clause No 7") with no
# further explanation - reason_category()'s keyword matching can't see
# what a clause number MEANS. Read the actual tender/order documents
# behind several of these citations (hin+eng OCR where scanned, native
# text where not - see PROGRESS_LOG.md) to build this glossary:
#
#   Annexure IX clause 7  - "In case any one batch is found NSQ..." -
#     QUALITY, confirmed by the clause's own quoted text (universal -
#     applies to every row citing this clause, not just the one read).
#   Annexure IX clause 12 - "Blacklisting for Non Supply... supplier
#     fails to supply the ordered quantity..." - NON-SUPPLY, same
#     universal-text reasoning.
#   T-213 clause 10.6 - "price quoted... shall not exceed the DPCO/NPPA
#     ceiling price... bid will be disqualified" - OTHER (a pricing/
#     bid-compliance violation, not a quality or supply failure).
#   T-213 clause 12.5 - "successful bidder fails to execute the
#     agreement and/or deposit the security deposit... EMD forfeited" -
#     OTHER (contract-execution failure, not a quality finding).
#   T-557 clause 10.9 is explicitly NOT a single-meaning clause - its
#     own text lists 6 different possible grounds (delay in agreement,
#     poor after-sales, delay in supply, collusion/fraud, misleading,
#     "any other violation") - so a bare "clause 10.9" citation is NOT
#     safe to override in general. The one row actually read (Unicare
#     Industries, forged technical-bid documents) IS overridden by its
#     own Entry ID, since that specific case's ground is known.
#
# The ~250 MPPHSCL rows with NO clause number cited at all ("Violation
# of Tender Terms and Condition" alone) are NOT covered here - see
# build_mpphscl_unclassified.py instead (a separate, OCR-based report).
MPPHSCL_CLAUSE_TEXT_OVERRIDES = [
    (re.compile(r"annexure\s*ix\s*clause\s*no\.?\s*7\b", re.I),
     "Quality", "Annexure IX clause 7 (NSQ batch failure) - see 10_merge_master.py MPPHSCL_CLAUSE_TEXT_OVERRIDES"),
    (re.compile(r"annexure\s*ix\s*clause\s*no\.?\s*12\b", re.I),
     "Non-supply", "Annexure IX clause 12 (non-supply) - see 10_merge_master.py MPPHSCL_CLAUSE_TEXT_OVERRIDES"),
    (re.compile(r"(?:t-213\s*)?(?:clause\s*(?:no\.?\s*)?)?10\s*\.?\s*6\b", re.I),
     "Other", "T-213 clause 10.6 (price above DPCO/NPPA ceiling) - see 10_merge_master.py MPPHSCL_CLAUSE_TEXT_OVERRIDES"),
    (re.compile(r"(?:t-213\s*)?(?:clause\s*(?:no\.?\s*)?)?12\s*\.?\s*5\b", re.I),
     "Other", "T-213 clause 12.5 (failed to execute agreement/EMD forfeited) - see 10_merge_master.py MPPHSCL_CLAUSE_TEXT_OVERRIDES"),
]
# Entry-ID-specific override - clause 10.9's meaning depends on the
# individual case, so only this ONE row (read directly) is overridden,
# not every row citing "clause 10.9".
MPPHSCL_ENTRY_ID_OVERRIDES = {
    "MPPHSCL-FIRM-5690-2024": ("Other", "T-557 clause 10.9, this case: forged/edited purchase "
                               "order submitted in technical bid - see 10_merge_master.py MPPHSCL_ENTRY_ID_OVERRIDES"),
}

# Hand-verified overrides from the "Random-20 check" sheet of the user's
# own document check (document_checks_2026-11-01.xlsx) - the user
# read each of these 20 letters directly from the page images (not OCR,
# not a keyword guess) as part of checking build_mpphscl_unclassified.py's
# 164-row auto-classified batch (checklist item 12). Per PROJECT_RULES.md rule F,
# an OCR-derived value is never trusted until the user has checked it
# against the PDF - these 20 rows have now had exactly that check, so they
# move from "OCR - check" (unconfirmed) to a genuinely confirmed category,
# unlike the other ~144 rows in the same batch (see
# build_mpphscl_unclassified.py's improved classify_text(), which covers
# those with a better auto-guess, still unconfirmed, still "OCR - check").
# Kept as its OWN dict (not merged into MPPHSCL_ENTRY_ID_OVERRIDES above)
# so the two kinds of override - a few hand-read one-offs found while
# building the clause glossary, vs this batch-review pass - stay separately
# traceable to where each one came from.
MPPHSCL_HAND_VERIFIED_OVERRIDES = {
    "MPPHSCL-FIRM-1507-2018": ("Other",
        "Order 2335 dt 26-07-2018 LIFTS an earlier 1-year ban (rates above NPPA/DPCO) after the "
        "firm agreed to NPPA rates - ban REVOKED, not active (see also MPPHSCL_COURT_OVERRIDES "
        "in 07_mpphscl.py)"),
    "MPPHSCL-FIRM-1586-2022-b": ("Other",
        "Glucometer strips: agreement not executed; 2-year debar; appeal heard 10-10-2022"),
    "MPPHSCL-FIRM-2282-2017": ("Other",
        "Manufacturing licence submitted was fake; barred 5 years - forged document"),
    "MPPHSCL-FIRM-273-2025": ("Other",
        "Hid that Tamil Nadu (TNMSC) had blacklisted its Phenytoin Sodium Inj on quality "
        "grounds; false affidavit; debarred 5 years - concealed another state's quality "
        "blacklisting (see also output/concealment_cases.xlsx)"),
    "MPPHSCL-FIRM-2793-2024": ("Other",
        "Rifampicin 450/300mg quoted above DPCO/NPPA ceiling; clause 10.6; 3 months"),
    "MPPHSCL-FIRM-3370-2022": ("Other",
        "4 items: agreement not executed; 2 years"),
    "MPPHSCL-FIRM-3589-2023": ("Other",
        "Salbutamol 4mg above DPCO; clause 10.6; 3 months"),
    "MPPHSCL-FIRM-3591-2023": ("Other",
        "Neostigmine inj above DPCO; clause 10.6; 3 months"),
    "MPPHSCL-FIRM-5283-2019": ("Other",
        "Silver Sulphadiazine cream above DPCO; clause 10.6; 3 months"),
    # Both the base ID and its "-b" dedup twin (MPPHSCL's live Firm page lists this order
    # 7876 twice - see 07_mpphscl.py) used to need an override here too, for the appeal outcome
    # (found 2026-11-02 while verifying - the "-b" row was silently left at "To check" the first
    # time). Moved out to MPPHSCL_SUPERSEDED_OVERRIDES below, 2026-11-07 - that fix's own
    # provenance (the automated SHA-256/OCR check) is NOT "the user, hand-verified against
    # the letter image, document_checks_2026-11-01.xlsx" like every entry in THIS dict,
    # and the fixed wrapper text immediately below classify_row()'s MPPHSCL_HAND_VERIFIED_
    # OVERRIDES branch would have misstated that if left here - see MPPHSCL_SUPERSEDED_OVERRIDES's
    # own comment and classify_row()'s separate branch for it.
    "MPPHSCL-FIRM-9445-2023": ("Other",
        "Surgical sutures: agreement not executed; 2 years; EMD forfeited"),
    "MPPHSCL-PRODUCT-133943-2026": ("Other",
        "Carboplatin 450mg: agreement not executed (raw material price rise); 6 months"),
    "MPPHSCL-PRODUCT-3361-2023": ("Other",
        "Hospital furniture (bedside locker, bed) not as per specification; clause 10.9 - "
        "non-conforming goods, not a drug-quality finding"),
    "MPPHSCL-PRODUCT-3367-2023": ("Other",
        "Hospital furniture not as per specification - non-conforming goods, not a drug-quality "
        "finding"),
    "MPPHSCL-PRODUCT-5040-2023": ("Other",
        "Hospital furniture (examination table) not as per specification - non-conforming "
        "goods, not a drug-quality finding"),
    "MPPHSCL-PRODUCT-5066-2023": ("Other",
        "Lactulose: agreement not executed (wrong rate quoted); clause 12.5"),
    "MPPHSCL-PRODUCT-7157-2020": ("Other",
        "Human Insulin: agreement not executed (force majeure claimed); 2 years"),
    "MPPHSCL-PRODUCT-9115-2021": ("Other",
        "Calcium + Vit D tabs: agreement not executed; clause 12.5; 2 years"),
    "MPPHSCL-PRODUCT-921-2021": ("Other",
        "Cotton crepe bandages: agreement not executed; 6 months"),
    "MPPHSCL-PRODUCT-927-2021": ("Other",
        "Latex examination gloves: agreement not executed; 6 months"),
}

# Meril Endo Surgery's MPPHSCL order 7876 (Commissioner's appeal, order 167 dt 20-01-2025 -
# see MPPHSCL_HAND_VERIFIED_OVERRIDES above for how that appeal result was found) survives
# only as an ITEM-level debarment (MIS93). RESOLVED 2026-11-07, by the automated check
# (NOT the user reading document_checks_2026-11-01.xlsx, unlike every entry above -
# kept in its OWN dict for exactly that reason, so classify_row() can give it its own, honest
# status wording instead of MPPHSCL_HAND_VERIFIED_OVERRIDES's fixed "hand-verified against the
# letter image by the user" wrapper): MPPHSCL-PRODUCT-166-2024 - MPPHSCL's own Product-list
# scrape of the SAME item (MIS93), SAME dates (06-Nov-2024 to 05-Nov-2026), under its own
# separate order-reference scheme (SR/166/F791/2024, vs the Firm page's 7876/F605/2024) - IS
# that surviving item-level debarment, confirmed via a SHA-256 hash: MPPHSCL-PRODUCT-166-2024.pdf
# and MPPHSCL-FIRM-7876-2024.pdf are byte-identical files (a hash fact, not an OCR read - see
# source_issues.csv, 07-11-2026). The "-b" dedup twin's own PDF is a different scan (the
# dispatch/forwarding copy) - not hash-identical, but OCR'd and cross-referenced against the
# same order number/date/item/EMD figure/firm/address (OCR-derived, PROJECT_RULES.md rule F, so this
# one specific identification is flagged "OCR - check" below, pending the user's own look, even
# though the base-pair hash match is not in doubt). Per the user's instruction: do NOT turn
# MPPHSCL-FIRM-7876-2024/-b into a second product-level ban (07_mpphscl.py no longer does -
# MPPHSCL_LEVEL_OVERRIDES/MPPHSCL_PRODUCT_OVERRIDES are now empty) - mark them superseded and
# exclude them from every active-ban count instead. That exclusion itself is carried by
# 07_mpphscl.py's MPPHSCL_COURT_OVERRIDES (any non-"Not stated" Court text makes
# build_rc_vs_bans.py's classify() return group "X" - stayed/revoked/superseded -
# unconditionally); this dict only fixes Category (both rows' true original ground is "Other" -
# the rate-quoting/agreement-execution error - never "Quality") and Status text.
MPPHSCL_SUPERSEDED_OVERRIDES = {
    "MPPHSCL-FIRM-7876-2024": ("Other",
        "Checked (verified by script via a SHA-256 hash match against "
        "MPPHSCL-PRODUCT-166-2024.pdf, 2026-11-07 - not an OCR read, see source_issues.csv) - "
        "SUPERSEDED BY APPEAL, see MPPHSCL-PRODUCT-166-2024 (same action). Commissioner's appeal "
        "order 167 dt 20-01-2025 (page 2, checked by the user 2026-11-02) SET ASIDE the "
        "firm-level blacklisting (order 7876 dt 06-11-2024) - the firm was blacklisted only for "
        "item MIS93, 2 years from 06-11-2024, for quoting a single-foil rate instead of a "
        "per-pack rate (EMD Rs 2,00,000 forfeited) - rate error, not a quality finding. Excluded "
        "from every active-ban count (see 07_mpphscl.py MPPHSCL_COURT_OVERRIDES); Level/"
        "Product(s) are the real, raw Firm-page data, not overridden to impersonate a second "
        "product row."),
    "MPPHSCL-FIRM-7876-2024-b": ("Other",
        "OCR - check (the automated check, 2026-11-07, pending the user's own look per "
        "PROJECT_RULES.md rule F - this copy's own PDF is a different scan from MPPHSCL-FIRM-7876-2024's "
        "hash-identical pair, so this specific identification rests on an OCR cross-reference "
        "against order 7876 dt 06.11.2024/item MIS93/EMD Rs 2,00,000/firm+address - not a hash "
        "match - see source_issues.csv) - SUPERSEDED BY APPEAL, see MPPHSCL-PRODUCT-166-2024 "
        "(same action). Commissioner's appeal order 167 dt 20-01-2025 (page 2, checked by the "
        "user 2026-11-02) SET ASIDE the firm-level blacklisting (order 7876 dt 06-11-2024) - the "
        "firm was blacklisted only for item MIS93, 2 years from 06-11-2024, for quoting a "
        "single-foil rate instead of a per-pack rate (EMD Rs 2,00,000 forfeited) - rate error, "
        "not a quality finding. Excluded from every active-ban count (see 07_mpphscl.py "
        "MPPHSCL_COURT_OVERRIDES); Level/Product(s) are the real, raw Firm-page data, not "
        "overridden to impersonate a second product row."),
    "MPPHSCL-PRODUCT-166-2024": ("Other",
        "Checked (verified by script via a SHA-256 hash match against "
        "MPPHSCL-FIRM-7876-2024.pdf, 2026-11-07 - not an OCR read, see source_issues.csv) - this "
        "is the AUTHORITATIVE, currently-active item-level debarment (MIS93, 2 years from "
        "06-11-2024, rate-quoting/agreement-execution error, EMD Rs 2,00,000 forfeited) that "
        "Meril Endo Surgery's MPPHSCL appeal (order 167 dt 20-01-2025) survives as. "
        "MPPHSCL-FIRM-7876-2024 and its -b twin are the superseded, duplicate Firm-page listing "
        "of this SAME action (see their own Court-stay field) - not a separate finding."),
}

# RMSC: reason_category() auto-categorizes from the ENGLISH text on the
# live page alone ("Violation of conditions" -> "To check"), but reading
# the actual order letter (Hindi, OCR'd - PROGRESS_LOG.md Entry 13-14)
# revealed this ONE row is a BID-RATE CORRECTION dispute (not a drug-
# quality finding at all) - a hand-verified, Entry-ID-keyed override, not
# a general rule.
RMSC_REASON_CATEGORY_OVERRIDES = {
    "RMSC-DEBAR-1837-12-08-2025": "Other - tender violation (bid-rate change after financial bid opening)",
}

# KSMSCL (16_ksmscl.py's hand-transcribed KARNATAKA_FULL_LIST, Entry 20):
# 3 rows' own "Reason" column literally says "Quality" or "Quality (QC
# order)" - a confident, hand-verified quality finding - but
# reason_category()'s QUALITY_WORDS list requires a more specific phrase
# ("quality failure", "quality ground", etc.), not the bare word
# "Quality" alone (deliberately, to avoid matching unrelated uses of
# "quality" elsewhere - e.g. "quality assurance certificate not
# submitted" would not actually be a product-quality finding). Found
# 2026-11-01 while verifying the Karnataka rewrite: these 3 rows fell to
# "To check" instead. A hand-verified, Entry-ID-keyed override instead of
# broadening the shared global word list.
KSMSCL_REASON_CATEGORY_OVERRIDES = {
    "KSMSCL-F-2023-24-1": "Quality",    # Jackson Laboratories - "Quality (QC order)"
    "KSMSCL-F-2023-24-2": "Quality",    # Sontex Rolled Bandage Works - "Quality (QC order)"
    "KSMSCL-P-2020-21-1": "Quality",    # Mahendra Labs Pvt Ltd - "Quality"
}


def classify_row(buyer, entry_id, reason_as_written):
    """Returns (category, status_override_or_None). Applies the hand-
    verified, buyer-specific overrides above FIRST (they are facts read
    directly from a specific order letter, not guesses from keywords),
    then falls back to the generic reason_category() keyword match.
    `status_override` replicates the exact "Auto - check (category
    overridden - ...)" wording the old per-buyer-script version used to
    write directly, so moving this here does not change that text."""
    if buyer == "MPPHSCL":
        if entry_id in MPPHSCL_SUPERSEDED_OVERRIDES:
            # Own branch, own status text - see MPPHSCL_SUPERSEDED_OVERRIDES's comment: this
            # finding's provenance (the automated SHA-256/OCR check, 2026-11-07) is NOT the
            # user hand-verifying against the letter image, document_checks_2026-11-01.xlsx,
            # so it must not go through MPPHSCL_HAND_VERIFIED_OVERRIDES's fixed wrapper below -
            # the note text itself already states its own status ("Checked"/"OCR - check").
            return MPPHSCL_SUPERSEDED_OVERRIDES[entry_id]
        if entry_id in MPPHSCL_HAND_VERIFIED_OVERRIDES:
            cat, note = MPPHSCL_HAND_VERIFIED_OVERRIDES[entry_id]
            return cat, (f"Checked (hand-verified against the letter image by the user, "
                         f"document_checks_2026-11-01.xlsx, 2026-11-01 - {note})")
        if entry_id in MPPHSCL_ENTRY_ID_OVERRIDES:
            cat, note = MPPHSCL_ENTRY_ID_OVERRIDES[entry_id]
            return cat, f"Auto - check (category overridden - {note})"
        for pattern, cat, note in MPPHSCL_CLAUSE_TEXT_OVERRIDES:
            if pattern.search(reason_as_written or ""):
                return cat, f"Auto - check (category overridden - {note})"
    if buyer == "RMSC" and entry_id in RMSC_REASON_CATEGORY_OVERRIDES:
        return RMSC_REASON_CATEGORY_OVERRIDES[entry_id], None
    if buyer == "KSMSCL" and entry_id in KSMSCL_REASON_CATEGORY_OVERRIDES:
        return KSMSCL_REASON_CATEGORY_OVERRIDES[entry_id], None
    return reason_category(reason_as_written), None


def _run_module(filename):
    path = Path(__file__).parent / filename
    spec = importlib.util.spec_from_file_location(path.stem, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.main()

SKIPPED_HEADER_PREFIX = ["Buyer", "Where", "Number of cells"]
UNMATCHED_HEADER = ["Buyer", "Firm name (as written)", "Firm name (standard)", "Reason (as written)",
                    "Revocation date", "Source link", "Note"]


def buyer_name_from_filename(path):
    return path.stem                      # "AP.xlsx" -> "AP"


def read_master_list(wb, buyer):
    if "Master list" not in wb.sheetnames:
        print(f"  WARNING - {buyer}.xlsx has no 'Master list' sheet - skipped entirely.")
        return []
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header, data = rows[0], rows[1:]
    if list(header) != COLUMNS:
        print(f"  WARNING - {buyer}.xlsx 'Master list' header does not match the 19 master "
              f"columns - skipped entirely so it cannot corrupt the combined file. "
              f"Expected: {COLUMNS}  Found: {list(header)}")
        return []
    return [list(r) for r in data if any(c is not None and str(c).strip() for c in r)]


def read_skipped(wb, buyer):
    if "Skipped rows" not in wb.sheetnames:
        return [], 0
    ws = wb["Skipped rows"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], 0
    header, data = rows[0], rows[1:]
    out = []
    for r in data:
        if not any(c is not None and str(c).strip() for c in r):
            continue
        out.append([buyer] + [c if c is not None else "" for c in r])
    return out, len(header) - 2          # how many "cell_i" columns this buyer used


def read_unmatched(wb, buyer):
    if "Unmatched revocations" not in wb.sheetnames:
        return []
    ws = wb["Unmatched revocations"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header, data = rows[0], rows[1:]
    out = []
    for r in data:
        if not any(c is not None and str(c).strip() for c in r):
            continue
        # the buyer sheet already has its own "Buyer" column first - keep it as written
        out.append([c if c is not None else "" for c in r])
    return out


def main():
    buyer_files = sorted(BUYERS.glob("*.xlsx"))          # fixed order = reproducible result
    if not buyer_files:
        print(f"No buyer files found in {BUYERS} - run the buyer scripts first "
              f"(03_build_master.py, 03_assam_master.py, 04_pmbi.py, ...).")
        return

    all_rows = []
    all_skipped = []
    all_unmatched = []
    max_skipped_width = 0
    per_buyer_counts = {}

    for path in buyer_files:
        buyer = buyer_name_from_filename(path)
        wb = load_workbook(path, read_only=True)
        rows = read_master_list(wb, buyer)
        skipped, width = read_skipped(wb, buyer)
        unmatched = read_unmatched(wb, buyer)

        all_rows.extend(rows)
        all_skipped.extend(skipped)
        all_unmatched.extend(unmatched)
        max_skipped_width = max(max_skipped_width, width)
        per_buyer_counts[buyer] = len(rows)

    # sanity check: Entry IDs should already be unique (buyer-prefixed) - warn, don't silently fix
    ids = [r[0] for r in all_rows]
    dupes = sorted({i for i in ids if ids.count(i) > 1})
    if dupes:
        print("WARNING - duplicate Entry IDs found ACROSS buyer files (should not happen, since IDs "
              "are buyer-prefixed) - fix the source buyer script, this script will not rename them:")
        for d in dupes:
            print("  ", d)

    # Centralized classification, added 2026-11-01 (Entry 19, checklist
    # item 13) - see classify_row() above. Every buyer script now writes
    # common.CATEGORY_PENDING instead of a computed category; this
    # overwrites it with the real value on every row, every merge, BEFORE
    # quality_level() runs (which needs the real category, not the
    # placeholder).
    category_changes = 0
    for r in all_rows:
        cat, status_override = classify_row(r[BUYER_IDX], r[ENTRY_ID_IDX], r[REASON_IDX])
        if r[CATEGORY_IDX] != cat:
            category_changes += 1
        r[CATEGORY_IDX] = cat
        if status_override:
            r[STATUS_IDX] = status_override

    matches = find_matches(all_rows)

    # Q1/Q2 quality-level split, added 2026-10-31 - see quality_level()
    # in common.py for the legal basis (Drugs & Cosmetics Act 1940
    # Section 25) and PROGRESS_LOG.md section 3.5 for the full record.
    # Computed HERE (once, centrally) rather than inside every buyer
    # script, since it only needs 3 columns each buyer file already has.
    all_rows = [r + [quality_level(r[REASON_IDX], r[LAB_IDX], r[CATEGORY_IDX])] for r in all_rows]
    master_header = COLUMNS + ["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"]

    # pad every Skipped-rows row to the same width before combining
    padded_skipped = []
    for row in all_skipped:
        fixed = row[:3]                                   # Buyer, Where, Number of cells
        cells = row[3:]
        cells = cells[:max_skipped_width] + [""] * max(0, max_skipped_width - len(cells))
        padded_skipped.append(fixed + cells)
    skipped_header = SKIPPED_HEADER_PREFIX + [f"cell_{i}" for i in range(1, max_skipped_width + 1)]

    wb = Workbook()
    write_sheet(wb.active, master_header, all_rows,
                [10, 30, 30, 34, 10, 14, 12, 26, 13, 13, 14, 34, 22, 18, 26, 40, 30, 16, 12, 20])
    wb.active.title = "Master list"
    write_sheet(wb.create_sheet("Matches"),
                ["Firm name (standard)", "Times listed", "Entry IDs", "Buyers", "What it means"],
                matches, [36, 8, 30, 30, 40])
    write_sheet(wb.create_sheet("Skipped rows"), skipped_header, padded_skipped,
                [12, 40, 8] + [18] * max_skipped_width)
    write_sheet(wb.create_sheet("Unmatched revocations"), UNMATCHED_HEADER, all_unmatched,
                [10, 34, 30, 34, 16, 40, 60])
    out = OUT / "master_list.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Rows where classify_row() changed the category away from the buyer's own written "
          f"value: {category_changes} (expected to be every row with real buyer data the first "
          f"time this runs, since every buyer script now writes CATEGORY_PENDING instead of a "
          f"real category - compare against a BEFORE snapshot of master_list.xlsx to check "
          f"whether the FINAL category actually changed, not just whether this overwrote a "
          f"placeholder)")
    print(f"  Buyer files combined   : {len(buyer_files)}  ({', '.join(p.stem for p in buyer_files)})")
    for buyer, count in per_buyer_counts.items():
        print(f"    {buyer:8s} : {count} rows")
    print(f"  Master list rows       : {len(all_rows)}")
    print(f"  Matches                : {len(matches)}")
    print(f"  Skipped rows           : {len(all_skipped)}")
    print(f"  Unmatched revocations  : {len(all_unmatched)}")

    print("\nRunning 11_cross_buyer_pairs.py ...")
    _run_module("11_cross_buyer_pairs.py")
    print("\nRunning 12_candidate_gaps.py ...")
    _run_module("12_candidate_gaps.py")


if __name__ == "__main__":
    main()
