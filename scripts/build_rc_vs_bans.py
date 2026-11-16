"""
build_rc_vs_bans.py - READ-ONLY report: RMSC rate-contract suppliers whose
standard name EXACTLY matches a firm banned by ANOTHER buyer, with a
strength classification for the overlapping cases.

Not part of the numbered pipeline - run it by hand whenever
rate_contracts/RMSC.xlsx or master_list.xlsx change. Creates NO new data -
every fact in the output already exists in those two files; this only
cross-references and classifies them.

"Another buyer" means any buyer OTHER than RMSC itself (RMSC is the one
issuing the rate contract being checked - a ban BY RMSC on the same firm
is a different, same-buyer question, not what this report is for).

Only EXACT standard-name matches are used (same as the confirmed network
layer in network_firms_buyers.xlsx) - no fuzzy matching on FIRM names
here. A supplier whose name only fuzzy-matches a banned firm belongs in
Part C's match_candidates.xlsx for a human decision, not in this report.

Three sheets:
  - RC vs bans : every (RC drug line, matching ban entry) pair, with an
                 overlap check (yes/no/unknown) - unchanged in shape from
                 the first version of this script.
  - Strength   : only the OVERLAPPING ones (overlap = yes), classified:
      A = the other buyer's ban is FIRM-level, reason category Quality,
          and active during the RC period
      B = the ban is PRODUCT-level, reason category Quality, AND it is
          the SAME MEDICINE AND SAME DOSAGE FORM as the RC line (word-
          overlap heuristic on the medicine name, PLUS a dosage-form
          match - Tablet/Injection/Syrup/etc. must agree; strength
          (e.g. 500mg vs 250mg) is allowed to differ, but noted when it
          does) - NOT exact matching, so both names are always shown
          for you to check by eye
      C = product-level Quality ban, but a DIFFERENT drug OR a
          different dosage form from the RC line (both are shown, and
          the "Why" column says which one failed to match)
      D = Non-supply, "Both - check", or any other KNOWN reason category
      U = reason category is "To check" - genuinely unknown, not the
          same as D (added 2026-10-31 per the user's review, so an
          unclassified reason no longer hides inside "Non-supply or
          other")
      X = the ban was stayed by a court or has been revoked - EXCLUDED
          from A-D/U (shown in its own group instead of silently vanishing)
  - By firm    : rolled up to firm | group | banning buyer | ban from/to |
                 RC drug lines in that group (count) | strongest example
                 row - "strongest" meaning the row with the fewest
                 judgment calls (an exact drug-name substring match beats
                 a looser word-overlap match, alphabetically first as the
                 final tie-break for determinism).

The drug-name match for group B/C is a HEURISTIC (strip trailing "[code]"
and dosage/unit words, then require at least one shared significant
word, AND a matching dosage form) - it is not a confident identity
check, which is exactly why both drug names are always shown rather
than silently deciding for you. Word matching also compares each side's
words FUSED together with no spaces, so a name split across several
words on one side (e.g. "CARBOXY METHYL CELLULOSE") still matches the
same name written as one compound word on the other side (e.g.
"Carboxymethylcellulose") - added 2026-10-31 after the user found this
exact case was being missed.

IMPORTANT, found when adding the dosage-form check: all 3 of the
previously-reported Group B matches (Nestor x2, Zee x1) turned out to
be DIFFERENT dosage forms on close reading - Nestor's RC Tablet vs the
APMSIDC ban's Syrup (both times), and Zee's RC oral Suspension/Syrup vs
the PMBI ban's Tablet. All 3 move to Group C under this stricter rule -
reported plainly, not smoothed over, since it changes a result the
user had already been given.

Group A/B rows also get a clause-based re-classification, added after
actually reading RMSC's own tender documents (see
output/rmsc_tender_review.xlsx): RMSC's eligibility clause (page 9 of
both readable tenders) only treats an OTHER state/agency's ban as
disqualifying if its ground was conviction-by-court or Not of Standard
Quality - a ban elsewhere for any other reason does not disqualify under
this clause. Three new "Strength" columns capture this: whether the
OTHER ban's own reason text is NSQ/conviction, whether that ban was
active on the RC's own tender's bid-submission deadline (only known for
the 2 tenders whose PDFs were readable - "unknown" otherwise, never
guessed), and a combined classification ("Possible breach of RMSC tender
rule" / "Allowed by RMSC rule (policy gap)" / an explicit "unknown"
variant). C/D/X rows are marked N/A for these three columns, since this
check is scoped to Group A/B only.

Reads:
  output/rate_contracts/RMSC.xlsx
  output/master_list.xlsx
Writes:
  output/rc_vs_bans.xlsx
"""

import re
from datetime import datetime, date
from collections import defaultdict

from openpyxl import load_workbook, Workbook

from common import OUT, write_sheet

DATE_FORMATS = ["%d-%m-%Y", "%d-%b-%Y", "%d-%B-%Y"]

# Generic pharmaceutical/formulation words that say nothing about WHICH
# drug this is - found a real false positive during testing: "Cefuroxime
# Axetil 500mg film coated Tablet" vs "Capecitabine Tablet" matched on
# "FILM"/"COATED" alone (completely different drugs) before these were
# added.
DOSAGE_WORDS = {"TABLET", "TABLETS", "TAB", "INJECTION", "INJ", "SYRUP", "SYP", "CAPSULE",
                "CAPSULES", "CAP", "SOLUTION", "SUSPENSION", "DROP", "DROPS", "CREAM", "GEL",
                "OINTMENT", "LOTION", "POWDER", "SACHET", "VIAL", "AMPOULE", "AMP", "STRIP",
                "IP", "BP", "USP", "EACH", "PIECE", "WITH", "FOR", "AND", "THE",
                "FILM", "COATED", "UNCOATED", "ENTERIC", "DELAYED", "RELEASE", "EXTENDED",
                "SUSTAINED", "PLAIN", "PURE", "DRIED", "FORM", "CONTAINS", "CONTAINING",
                "EQUIVALENT", "ELEMENTAL", "SIZE", "TYPE", "USE", "SINGLE", "STERILE"}

# Canonical dosage FORM per word - several spellings/abbreviations map to
# the same real-world form (e.g. TAB/TABLET/TABLETS are all "tablet").
# Added 2026-10-31 (Group B now requires the SAME form, not just a
# shared word - see module docstring for the 3 cases this demoted).
DOSAGE_FORM_MAP = {
    "TABLET": "TABLET", "TABLETS": "TABLET", "TAB": "TABLET",
    "INJECTION": "INJECTION", "INJ": "INJECTION",
    "SYRUP": "SYRUP", "SYP": "SYRUP",
    "SUSPENSION": "SUSPENSION",
    "SOLUTION": "SOLUTION",
    "CAPSULE": "CAPSULE", "CAPSULES": "CAPSULE", "CAP": "CAPSULE",
    "DROP": "DROPS", "DROPS": "DROPS",
    "CREAM": "CREAM", "GEL": "GEL", "OINTMENT": "OINTMENT", "LOTION": "LOTION",
    "POWDER": "POWDER", "SACHET": "SACHET", "VIAL": "VIAL", "AMPOULE": "AMPOULE", "AMP": "AMPOULE",
}


def dosage_form(text):
    """Returns the canonical dosage form found in `text` (e.g. "TABLET"),
    or None if no recognized form word is present - never guessed."""
    text = re.sub(r"[^\w\s]", " ", (text or "").upper())
    for w in text.split():
        if w in DOSAGE_FORM_MAP:
            return DOSAGE_FORM_MAP[w]
    return None


def parse_date(text):
    text = (text or "").strip()
    if not text or text.lower() == "not stated":
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def overlap(rc_from, rc_to, ban_from, ban_to):
    if None in (rc_from, rc_to, ban_from, ban_to):
        return "unknown"
    return "yes" if (rc_from <= ban_to and ban_from <= rc_to) else "no"


def drug_words(text):
    """Significant words for the group B/C drug-name heuristic - strips a
    trailing [code], units/dosage-form words, and pure numbers. Also
    adds the SAME significant words fused together with no spaces, so
    a name split across several words on one side (e.g. "CARBOXY
    METHYL CELLULOSE") still matches the same name written as one
    compound word on the other side (e.g. "Carboxymethylcellulose") -
    added 2026-10-31 per the user's review."""
    text = re.sub(r"\[.*?\]\s*$", "", text or "")
    text = re.sub(r"[^\w\s]", " ", text.upper())
    kept = []
    for w in text.split():
        if len(w) < 4 or w in DOSAGE_WORDS or w.isdigit() or re.fullmatch(r"\d+(MG|ML|GM|MCG|IU)", w):
            continue
        kept.append(w)
    words = set(kept)
    if len(kept) > 1:
        fused = "".join(kept)
        if len(fused) >= 8:
            words.add(fused)
    return words


# RMSC clause check (added after reading the actual tender documents -
# see output/rmsc_tender_review.xlsx and PROGRESS_LOG.md Entry 9). RMSC's
# own eligibility clause (page 9 of both NIB-14/2024/69 and NIB-10/2024,
# identical wording) only treats an OTHER state/agency's ban as relevant
# if its ground was "conviction by court of law" or the product being
# "Not of Standard Quality" (NoSQ/spurious/adulterated) - a ban elsewhere
# for any OTHER reason (e.g. non-supply) does not disqualify a bidder
# under this clause, by its own wording. This is scoped to Group A/B/C
# (see WHOLE_FIRM_GROUPS above) - D/U/X rows are marked N/A rather than
# left blank, since a blank could be misread as "checked, found no".
#
# CONVICTION_WORDS only (NOT an "NSQ-by-text" word list anymore) - fixed
# 2026-11-02 after the Protech Telelinks case: this used to ALSO treat
# any literal "Not of Standard Quality"/"NoSQ"/etc. wording in the ban's
# own scraped summary text as satisfying the Q1 (headline) ground on its
# own, regardless of quality_level_value - but APMSIDC's own summary
# text for Protech says "Supply of Not of Standard Quality Product"
# while its own "NSQ declared by" field names NO lab at all (quality_
# level() correctly computes Q2 for it). The project's own Q1/Q2 split
# (common.py) is supposed to be the SOLE test for "is this a statutory
# NSQ declaration" - a bare literal "NSQ" word in a scraped summary is
# not itself evidence of that (the summary text and the lab field are
# two different things). A court CONVICTION is still its own, separate,
# always-sufficient ground under the clause's own wording (it needs no
# lab at all), so that part is kept, just no longer bundled with the
# NSQ words.
CONVICTION_WORDS = ["conviction", "convicted"]

# Bid-submission deadlines for RMSC tenders whose PDFs were actually
# readable (see rmsc_tender_review.xlsx). NIB-14/2024 and NIB-10/2024
# came from clean, native PDF text - CONFIRMED, used freely. NIB-08/2024's
# tender notice was never found, so it stays "unknown", not guessed.
# Keyed by the same NIB-code+year pattern used in 15_rate_contracts.py's
# tender_id_part().
TENDER_DEADLINES = {
    "NIB14-2024": date(2025, 2, 24),   # NIB-14/2024/69, page 1 (native PDF text)
    "NIB10-2024": date(2024, 10, 8),   # NIB-10/2024, page 1 (native PDF text)
    # NIB-10/2026 (sutures): moved here from TENDER_DEADLINES_OCR_UNCHECKED 2026-11-07 - the
    # user confirmed 06-05-2026, 6 PM, BY EYE against pages 1-3 of the tender PDF (not OCR -
    # PROJECT_RULES.md rule F is satisfied), removing the leading-digit uncertainty flagged in Entry 20.
    "NIB10-2026": date(2026, 5, 6),
}

# NIB-01/2025's PDF is a scanned image (0 extractable characters) - this
# value came from Tesseract OCR on page 1, read twice with consistent
# results, but per PROJECT_RULES.md rule F an OCR value is NEVER trusted until
# the user has checked it against the PDF themselves. Kept in a SEPARATE
# dict so it can never silently merge into the confirmed one above, and
# clause_check() below always labels anything using it as PENDING.
TENDER_DEADLINES_OCR_UNCHECKED = {
    "NIB01-2025": date(2025, 6, 5),   # NIB-01/2025, OCR p.1 - data/text/ocr/NIB01_2025_NIB01_p1.txt
    # NIB-10/2026 (sutures) moved to TENDER_DEADLINES above, 2026-11-07 - now confirmed, not
    # OCR-unchecked. See that dict's comment.
}

TENDER_CODE_RE = re.compile(r"(NI[BT]-?\d+)", re.I)
TENDER_YEAR_RE = re.compile(r"\b(\d{4})(?:-\d{2,4})?\b")


def tender_deadline(tender_no):
    """Returns (date_or_None, source, short_code). source is "confirmed"
    (native PDF text), "ocr_unchecked" (rule F - needs the user's check),
    or "none" (tender not readable at all, not guessed)."""
    tender_no = tender_no or ""
    m = TENDER_CODE_RE.search(tender_no)
    y = TENDER_YEAR_RE.search(tender_no)
    if not m or not y:
        return None, "none", "UNKNOWN"
    code = re.sub(r"[^A-Za-z0-9]", "", m.group(1)).upper() + "-" + y.group(1)
    if code in TENDER_DEADLINES:
        return TENDER_DEADLINES[code], "confirmed", code
    if code in TENDER_DEADLINES_OCR_UNCHECKED:
        return TENDER_DEADLINES_OCR_UNCHECKED[code], "ocr_unchecked", code
    return None, "none", code


# Fixed label for every Q2 finding, everywhere it appears (strength rows,
# by-firm roll-ups, printed summaries) - per the user's decision
# 2026-11-01 (PROGRESS_LOG.md Entry 19): Q2 must always be described this
# way, never reworded per call site, so a reader always recognizes it as
# the same, weaker, separately-reported case.
Q2_LABEL = "quality failure in substance, not a formal NSQ declaration (no government lab named)"


# Which groups clause 5(l)(II) / 5(k)(ii) (the WHOLE-FIRM clause) reaches,
# vs which groups the PRODUCT-ONLY clauses 5(l)(I)/(III) / 5(k)(iii) reach.
# Re-read 2026-11-01 per the user's own document check
# (document_checks_2026-11-01.xlsx, "Rajasthan clause" sheet, exact
# clause text quoted there): clause II's own words are "The concern/
# company/firm which stands blacklisted/banned/debarred ... by any other
# State/Central Government ... on the ground of ... the products being
# found Not of Standard Quality (NoSQ) ... SHALL ALSO NOT BE ELIGIBLE TO
# PARTICIPATE IN THE BID" - this disqualifies the FIRM from the ENTIRE
# bid once ANY of its products was found NSQ elsewhere, regardless of
# whether THIS tender's own line item is the same product as the one the
# other buyer found NSQ. That is Groups A (firm-level ban), B (same
# product) AND C (different product - the ban is still an NSQ finding on
# THIS firm, by another buyer, even though it was scoped to a different
# product over there) - previously this script only checked A/B, which
# is why Zee/Meril/Nestor's real Group-C cases never surfaced as
# "Possible breach" even after their Q1 lab was correctly recorded (see
# 07_mpphscl.py/08_rmsc.py Lab overrides, added the same day).
#
# Clauses I/III (product-only) are narrower on their own words - "Bid
# should not be submitted for THOSE PRODUCTS ALSO for which the concern/
# company stands blacklisted ... by any other State" - this only bars
# bidding on the SPECIFIC product found NSQ elsewhere, which is exactly
# what Group B already means (same medicine, same dosage form). Group A
# (a bare firm-level ban elsewhere, no product named) does not by itself
# identify which product is barred under this narrower reading, so it is
# NOT counted as a product-only breach either - only Group B is.
WHOLE_FIRM_GROUPS = ("A", "B", "C")
PRODUCT_ONLY_GROUPS = ("B",)


def clause_check(group, reason_as_written, ban_from, ban_to, tender_no, quality_level_value):
    """Returns (allowed, active_on_deadline, q1_classification,
    q2_classification) for the WHOLE-FIRM reading (clause 5(l)(II) /
    5(k)(ii)) - Groups A/B/C. ("N/A", "N/A", "N/A (scope: Group A/B/C
    only)", "N/A (scope: Group A/B/C only)") for D/U/X, which are never a
    live NSQ ground (D = non-quality reason, U = reason genuinely
    unknown, X = stayed/revoked).

    Per the user's decision 2026-11-01 (PROGRESS_LOG.md Entry 19):
    HEADLINE findings use Q1 ONLY. A ban's reason text literally saying
    "NSQ"/"spurious"/etc. is ONE way to satisfy RMSC's clause - but a Q1
    finding (a named STATUTORY lab: Government Analyst, State Drug
    Testing Lab, CDSCO, CDTL, DCL - common.py's quality_level()) IS ALSO
    a formal NSQ declaration under the Drugs and Cosmetics Act even when
    the scraped summary text never uses the word "NSQ" - so either one
    makes the headline "Possible breach (Q1, clause ii)" on its own. Q2
    (a quality signal with NO statutory lab named - the buyer's own/
    empanelled lab, or no lab named at all) is NEVER folded into that
    headline - it is always returned and reported SEPARATELY, with the
    fixed Q2_LABEL text above, so a reader sees both counts side by side
    rather than having to notice a debatable case buried inside a single
    combined string.

    "active_on_deadline" == "No" (the ban's own period does not cover the
    bid deadline, but the RC contract period still overlaps it - the
    caller only calls this for overlap=="yes" rows) is reported as a
    POLICY GAP, not a breach: the firm was not technically ineligible to
    BID, but the ban became active again during the life of the
    contract, with no cancellation mechanism for an NSQ finding (General
    Conditions (o) covers only spurious/adulterated) - see the "Zee - CMC
    eye drops" row in document_checks_2026-11-01.xlsx."""
    if group not in WHOLE_FIRM_GROUPS:
        na = "N/A (scope: Group A/B/C only)"
        return "N/A", "N/A", na, na

    lower = (reason_as_written or "").lower()
    is_conviction_ground = any(w in lower for w in CONVICTION_WORDS)
    is_q1_ground = is_conviction_ground or quality_level_value == "Q1"
    is_q2_ground = (not is_q1_ground) and quality_level_value == "Q2"
    allowed = "Yes" if is_q1_ground else "No"

    deadline, source, code = tender_deadline(tender_no)
    if deadline is None:
        active = "unknown"
    else:
        bf, bt = parse_date(ban_from), parse_date(ban_to)
        active = "unknown" if (bf is None or bt is None) else \
            ("Yes" if bf.date() <= deadline <= bt.date() else "No")
    deadline_note = f" (tender {code} not readable)" if deadline is None else ""

    # --- Q1: the headline, and ONLY the headline ---
    if not is_q1_ground:
        q1_result = "No Q1 ground (other ban's reason text isn't NSQ/conviction, and no statutory lab is named)"
    elif active == "Yes" and source == "confirmed":
        q1_result = "Possible breach (Q1, clause ii - whole firm)"
    elif active == "Yes" and source == "ocr_unchecked":
        q1_result = ("Possible breach (Q1, clause ii - whole firm) - PENDING: deadline came from "
                      "OCR, not yet checked against the PDF by the user (PROJECT_RULES.md rule F)")
    elif active == "No":
        q1_result = f"Policy gap (Q1 ground, not active at bid deadline){deadline_note}"
    else:
        q1_result = f"Allowed ground (Q1), bid-deadline unknown{deadline_note}"

    # --- Q2: always computed and always returned, independent of Q1 -
    # never silently merged into q1_result above.
    if not is_q2_ground:
        q2_result = "N/A (not a Q2 case)"
    elif active == "Yes" and source == "confirmed":
        q2_result = f"Possible breach if Q2 counts as NSQ ground - {Q2_LABEL}"
    else:
        q2_result = (f"Allowed by RMSC rule (policy gap) if Q2 does NOT count as NSQ ground - "
                      f"{Q2_LABEL} - bid-deadline overlap {active}{deadline_note} also needed either way")

    return allowed, active, q1_result, q2_result


def product_only_reading(group, q1_result):
    """Derives the SECOND, narrower classification (clauses 5(l)(I)/(III)
    / 5(k)(iii) - product-only, see PRODUCT_ONLY_GROUPS above) from the
    same clause_check() result, rather than re-running the whole
    computation - the only thing that differs is WHICH groups count as a
    breach at all. Group B keeps the whole-firm result unchanged (it also
    satisfies the narrower reading). Group A/C get an explicit "no breach
    under this narrower reading" instead, so a reader is told WHY two
    columns disagree rather than seeing the same text twice. D/U/X stay N/A."""
    if group not in WHOLE_FIRM_GROUPS:
        return "N/A (scope: Group A/B/C only)"
    if group in PRODUCT_ONLY_GROUPS:
        return q1_result
    if group == "A":
        return ("Not counted under this narrower reading - the other ban is firm-level with no "
                "specific product named, so clause I/III's own wording (barring a bid only on "
                "\"those products\") does not identify which product is barred")
    return ("No breach under this narrower reading - different product/dosage form from the "
            "other buyer's ban (see \"Why\" column)")


def classify(ban_level, ban_cat, ban_court, rc_drug_name, ban_product):
    if ban_court and ban_court.strip().lower() != "not stated":
        return "X", ban_court
    is_firm = ban_level.split()[0].lower() == "firm" if ban_level else False
    is_quality = ban_cat == "Quality"
    if is_firm and is_quality:
        return "A", "Firm-level, Quality"
    if (not is_firm) and is_quality:
        shared = drug_words(rc_drug_name) & drug_words(ban_product)
        rc_form, ban_form = dosage_form(rc_drug_name), dosage_form(ban_product)
        same_form = rc_form is not None and rc_form == ban_form
        if shared and same_form:
            strength_note = "" if _same_strength(rc_drug_name, ban_product) else \
                " - NOTE: strengths differ, see both names"
            return "B", (f"Product-level, Quality, same medicine (shared word(s): "
                         f"{', '.join(sorted(shared))}) AND same dosage form ({rc_form})"
                         f"{strength_note}")
        if shared and not same_form:
            return "C", (f"Product-level, Quality, shared word(s) {', '.join(sorted(shared))} "
                         f"but DIFFERENT dosage form (RC: {rc_form or 'not identified'} vs "
                         f"ban: {ban_form or 'not identified'}) - not treated as the same product")
        return "C", "Product-level, Quality, no shared significant word found"
    if ban_cat == "To check":
        return "U", "Reason category is \"To check\" - unknown, not Non-supply/Other"
    return "D", f"Reason category: {ban_cat}"


STRENGTH_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(MG|ML|GM|MCG|IU|G|%)\b", re.I)


def _same_strength(rc_drug_name, ban_product):
    """Best-effort check of whether the two names cite the same
    strength figure(s) - informational only (strength differing does
    NOT change the A-X group), so a mismatch becomes a note, not a
    reclassification. Returns True if no strength was found on either
    side (nothing to compare, so nothing to flag)."""
    rc_strengths = {(m.group(1), m.group(2).upper()) for m in STRENGTH_RE.finditer(rc_drug_name or "")}
    ban_strengths = {(m.group(1), m.group(2).upper()) for m in STRENGTH_RE.finditer(ban_product or "")}
    if not rc_strengths or not ban_strengths:
        return True
    return bool(rc_strengths & ban_strengths)


def main():
    rc_wb = load_workbook(OUT / "rate_contracts" / "RMSC.xlsx", read_only=True)
    rc_ws = rc_wb["Rate contracts"]
    rc_rows = list(rc_ws.iter_rows(values_only=True))
    rc_header, rc_data = rc_rows[0], rc_rows[1:]
    rc_idx = {name: i for i, name in enumerate(rc_header)}

    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_ws = master_wb["Master list"]
    master_rows = list(master_ws.iter_rows(values_only=True))
    m_header, m_data = master_rows[0], master_rows[1:]
    m_idx = {name: i for i, name in enumerate(m_header)}

    bans_by_name = defaultdict(list)
    for r in m_data:
        if r[m_idx["Buyer"]] == "RMSC":
            continue                                        # "another buyer" excludes RMSC itself
        bans_by_name[r[m_idx["Firm name (standard)"]]].append(r)

    all_rows, strength_rows = [], []
    for rc in rc_data:
        supplier_std = rc[rc_idx["Supplier (standard)"]]
        bans = bans_by_name.get(supplier_std)
        if not bans:
            continue
        rc_from = parse_date(rc[rc_idx["Contract From"]])
        rc_to = parse_date(rc[rc_idx["Contract To"]])
        for ban in bans:
            ban_from = parse_date(ban[m_idx["Blacklisted from"]])
            ban_to = parse_date(ban[m_idx["Blacklisted to"]])
            ov = overlap(rc_from, rc_to, ban_from, ban_to)
            all_rows.append([
                rc[rc_idx["Supplier (as written)"]], rc[rc_idx["Drug Name"]],
                f"{rc[rc_idx['Contract From']]} to {rc[rc_idx['Contract To']]}",
                ban[m_idx["Buyer"]],
                f"{ban[m_idx['Blacklisted from']]} to {ban[m_idx['Blacklisted to']]}", ov,
            ])
            if ov == "yes":
                group, why = classify(ban[m_idx["Level (firm / product)"]],
                                       ban[m_idx["Reason category (Quality / Non-supply / Other / To check)"]],
                                       ban[m_idx["Court stay or revoked?"]],
                                       rc[rc_idx["Drug Name"]], ban[m_idx["Product(s)"]])
                ban_quality_level = ban[m_idx["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"]]
                allowed, active, clause_q1, clause_q2 = clause_check(
                    group, ban[m_idx["Reason (as written)"]],
                    ban[m_idx["Blacklisted from"]], ban[m_idx["Blacklisted to"]],
                    rc[rc_idx["Tender No"]], ban_quality_level)
                clause_product_only = product_only_reading(group, clause_q1)
                strength_rows.append({
                    "supplier": supplier_std, "supplier_written": rc[rc_idx["Supplier (as written)"]],
                    "rc_drug": rc[rc_idx["Drug Name"]], "rc_dates": f"{rc[rc_idx['Contract From']]} to {rc[rc_idx['Contract To']]}",
                    "rc_tender": rc[rc_idx["Tender No"]],
                    "buyer": ban[m_idx["Buyer"]], "ban_id": ban[m_idx["Entry ID"]],
                    "ban_level": ban[m_idx["Level (firm / product)"]], "ban_cat": ban[m_idx["Reason category (Quality / Non-supply / Other / To check)"]],
                    "ban_quality_level": ban_quality_level,
                    "ban_reason_written": ban[m_idx["Reason (as written)"]],
                    "ban_product": ban[m_idx["Product(s)"]],
                    "ban_dates": f"{ban[m_idx['Blacklisted from']]} to {ban[m_idx['Blacklisted to']]}",
                    "ban_court": ban[m_idx["Court stay or revoked?"]], "group": group, "why": why,
                    "clause_allowed": allowed, "clause_active": active,
                    "clause_q1": clause_q1, "clause_q2": clause_q2,
                    "clause_product_only": clause_product_only,
                })

    all_rows.sort(key=lambda r: (r[5] != "yes", r[5] != "unknown", r[0]))

    group_order = {"A": 0, "B": 1, "C": 2, "D": 3, "U": 4, "X": 5}
    strength_rows.sort(key=lambda r: (group_order[r["group"]], r["supplier"], r["rc_drug"]))

    # "By firm" roll-up: firm | group | banning buyer | ban from/to | RC drug lines (count) | strongest example
    by_firm = defaultdict(lambda: {"rc_drugs": [], "example": None, "q1_classes": [], "q2_classes": [],
                                   "product_only_classes": []})
    for r in strength_rows:
        key = (r["supplier"], r["group"], r["buyer"], r["ban_dates"])
        cell = by_firm[key]
        cell["rc_drugs"].append(r["rc_drug"])
        cell["q1_classes"].append(r["clause_q1"])
        cell["q2_classes"].append(r["clause_q2"])
        cell["product_only_classes"].append(r["clause_product_only"])
        # "strongest" = fewest judgment calls: an exact-looking match (group A has none to
        # question; group B's word-overlap is still a judgment call) - here, just take the
        # alphabetically-first RC drug line as the deterministic example.
        if cell["example"] is None or r["rc_drug"] < cell["example"]["rc_drug"]:
            cell["example"] = r

    # Worst-case ordering for the firm-level clause summary (Group A/B only) -
    # a firm with ANY RC line showing a possible breach is flagged as such,
    # even if its other RC lines are only "unknown". Q1 and Q2 are tracked
    # as two SEPARATE worst-cases (2026-11-01 decision) - never combined
    # into one "strongest of either" value, so the headline (Q1) column
    # never silently absorbs a Q2-only finding.
    def worst_q1(classes):
        breaches = [c for c in classes if c.startswith("Possible breach")]
        if breaches:
            return breaches[0]
        allowed = [c for c in classes if c.startswith("Allowed ground")]
        if allowed:
            return allowed[0]
        no_ground = [c for c in classes if c.startswith("No Q1 ground")]
        if no_ground:
            return no_ground[0]
        return classes[0] if classes else "N/A (scope: Group A/B only)"

    def worst_q2(classes):
        breaches = [c for c in classes if c.startswith("Possible breach")]
        if breaches:
            return breaches[0]
        gaps = [c for c in classes if c.startswith("Allowed by RMSC rule")]
        if gaps:
            return gaps[0]
        return classes[0] if classes else "N/A (not a Q2 case)"

    def worst_product_only(classes):
        breaches = [c for c in classes if c.startswith("Possible breach")]
        if breaches:
            return breaches[0]
        no_breach = [c for c in classes if c.startswith(("No breach", "Not counted"))]
        if no_breach:
            return no_breach[0]
        return classes[0] if classes else "N/A (scope: Group A/B/C only)"

    by_firm_rows = []
    for (firm, group, buyer, ban_dates), cell in sorted(by_firm.items(),
            key=lambda kv: (group_order[kv[0][1]], kv[0][0])):
        ex = cell["example"]
        by_firm_rows.append([firm, group, buyer, ban_dates, len(cell["rc_drugs"]),
                              f"{ex['rc_drug']} (vs ban {ex['ban_id']}: {ex['why']})",
                              worst_q1(cell["q1_classes"]), worst_q2(cell["q2_classes"]),
                              worst_product_only(cell["product_only_classes"])])

    print(f"RC rows checked: {len(rc_data)}")
    print(f"RC-vs-ban rows found: {len(all_rows)}")
    print(f"  overlap = yes: {sum(1 for r in all_rows if r[5] == 'yes')}")
    print(f"  overlap = no: {sum(1 for r in all_rows if r[5] == 'no')}")
    print(f"  overlap = unknown: {sum(1 for r in all_rows if r[5] == 'unknown')}")
    print()
    print("Strength groups (DISTINCT FIRMS per group, a firm can appear in more than one group):")
    for g in "ABCDUX":
        firms_in_group = {r[0] for r in by_firm_rows if r[1] == g}
        print(f"  {g}: {len(firms_in_group)} firms")
    print()
    whole_firm_rows = [r for r in by_firm_rows if r[1] in WHOLE_FIRM_GROUPS]
    print("Clause (ii) - WHOLE-FIRM reading, Group A/B/C (DISTINCT FIRMS, worst case per firm) -")
    print("re-read 2026-11-01 per the user's document check: clause 5(l)(II)/5(k)(ii) disqualifies")
    print("the FIRM from the whole bid on ANY other-buyer NSQ finding, not just a matching product -")
    print("both Q1 and Q2 shown separately and always, per the 2026-10-31/11-01 decisions - Q1 is")
    print("the HEADLINE, Q2 is never counted toward it:")
    print("  Q1 (headline):")
    q1_counts = defaultdict(set)
    for r in whole_firm_rows:
        q1_counts[r[6]].add(r[0])
    for cls, firms in q1_counts.items():
        print(f"    {cls}: {len(firms)} firm(s) ({', '.join(sorted(firms))})")
    print(f"  Q2 ({Q2_LABEL}) - reported separately, NOT part of the headline above:")
    q2_counts = defaultdict(set)
    for r in whole_firm_rows:
        q2_counts[r[7]].add(r[0])
    for cls, firms in q2_counts.items():
        print(f"    {cls}: {len(firms)} firm(s) ({', '.join(sorted(firms))})")
    q1_breach_firms = {r[0] for r in whole_firm_rows if r[6].startswith("Possible breach")}
    q2_breach_firms = {r[0] for r in whole_firm_rows if r[7].startswith("Possible breach")}
    policy_gap_firms = {r[0] for r in whole_firm_rows if r[6].startswith("Policy gap")}
    print()
    print(f"SUMMARY - Possible breach (Q1, clause ii - whole firm, HEADLINE): {len(q1_breach_firms)} "
          f"firm(s) ({', '.join(sorted(q1_breach_firms)) if q1_breach_firms else 'none'})")
    print(f"SUMMARY - Possible breach if Q2 counts as NSQ ground ({Q2_LABEL}): "
          f"{len(q2_breach_firms)} firm(s) ({', '.join(sorted(q2_breach_firms)) if q2_breach_firms else 'none'})")
    print(f"SUMMARY - Policy-gap firms (Q1 ground exists, but not active at bid deadline - banned "
          f"product/firm still under an active RC contract): {len(policy_gap_firms)} firm(s) "
          f"({', '.join(sorted(policy_gap_firms)) if policy_gap_firms else 'none'})")
    print()
    product_only_breach_firms = {r[0] for r in by_firm_rows if r[1] == "B" and r[8].startswith("Possible breach")}
    print(f"SUMMARY - Possible breach under the NARROWER product-only reading (clauses 5(l)(I)/"
          f"(III), 5(k)(iii) - Group B only): {len(product_only_breach_firms)} firm(s) "
          f"({', '.join(sorted(product_only_breach_firms)) if product_only_breach_firms else 'none'})")

    header = ["Supplier", "RC Drug", "RC From/To", "Banning Buyer", "Ban From/To", "Overlap?"]
    strength_header = ["Supplier (standard)", "Supplier (as written)", "RC Drug", "RC From/To",
                        "RC Tender No", "Banning Buyer", "Ban Entry ID", "Ban Level",
                        "Ban Reason category", "Ban Quality level (Q1/Q2/N-A)", "Ban Reason (as written)",
                        "Ban Product (as written)",
                        "Ban From/To", "Court stay or revoked?", "Group", "Why",
                        "Other ban reason allowed under RMSC clause? (NSQ/conviction/Q1 = yes)",
                        "Other ban active on bid deadline? (yes/no/unknown)",
                        "Clause ii (whole firm) - Q1 HEADLINE classification",
                        f"Clause ii (whole firm) - Q2 - {Q2_LABEL} (never counted in Q1)",
                        "Clauses i/iii (product-only, narrower) - classification"]
    by_firm_header = ["Firm (standard)", "Group", "Banning Buyer", "Ban From/To",
                       "RC drug lines in this group", "Strongest example row",
                       "Clause ii (whole firm) - Q1 HEADLINE (worst case across this firm's RC lines)",
                       f"Clause ii (whole firm) - Q2 - {Q2_LABEL} (worst case, never counted in Q1)",
                       "Clauses i/iii (product-only, narrower) - worst case"]

    wb = Workbook()
    write_sheet(wb.active, header, all_rows, [34, 50, 26, 16, 26, 10])
    wb.active.title = "RC vs bans"
    write_sheet(wb.create_sheet("Strength"), strength_header,
                [[r["supplier"], r["supplier_written"], r["rc_drug"], r["rc_dates"], r["rc_tender"],
                  r["buyer"], r["ban_id"], r["ban_level"], r["ban_cat"], r["ban_quality_level"],
                  r["ban_reason_written"],
                  r["ban_product"], r["ban_dates"], r["ban_court"], r["group"], r["why"],
                  r["clause_allowed"], r["clause_active"], r["clause_q1"], r["clause_q2"],
                  r["clause_product_only"]]
                 for r in strength_rows],
                [30, 30, 50, 26, 40, 12, 24, 16, 16, 14, 30, 40, 26, 30, 8, 46, 36, 30, 46, 60, 60])
    write_sheet(wb.create_sheet("By firm"), by_firm_header, by_firm_rows,
                [30, 8, 14, 26, 14, 70, 46, 60, 60])
    out = OUT / "rc_vs_bans.xlsx"
    wb.save(out)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
