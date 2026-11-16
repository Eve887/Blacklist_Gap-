"""
14_match_firms.py - Part C: cross-buyer fuzzy name matching (NOT auto-merged).

(Numbered 14, not 05 as originally sketched - 05 was already taken by
05_gmscl.py, and 13 by 13_bmsicl.py; see PROGRESS_LOG.md section 3.)

Reads:
  output/master_list.xlsx   (made by 10_merge_master.py)
Writes:
  output/match_candidates.xlsx  with two sheets:
    - Match candidates   : pairs that clear the score threshold AND share
                            a real (non-generic) word - these are worth
                            your time to look at.
    - Generic words only : pairs that clear the score threshold but share
                            NOTHING besides generic pharma/corporate words -
                            kept, not dropped, but separated out since
                            there is nothing to the Score 11 compare but
                            common filler words ("Pvt Ltd" etc).
  Each row has an empty "My decision" column (Same / Different / Not sure).

THREE design changes from the first version, made after testing against
real data (see PROGRESS_LOG.md section 3.5 for the full history):

1. Compares UNIQUE STANDARDIZED NAMES, not every row. 782 master-list rows
   collapse to 382 unique standard names - many buyers have several rows
   per firm (different products/years). Comparing rows directly produced
   216,891 pairs and a mostly-duplicate 26,723 "candidates" (the same two
   firms' names matching each other once per row-pair). Comparing the 382
   unique names instead is both faster and removes that duplication at the
   source. Each name-pair row lists every Entry ID and buyer behind BOTH
   sides, so nothing about row-level detail is lost.

2. The generic-word check is a real FILTER now (on word length, not just a
   warning label): strip common pharma/corporate words (list below), then
   require a remaining word of 4+ letters that is shared. "Shared" is
   fuzzy, not exact string equality - two words count as the same if
   fuzz.ratio() between them is >= 85 (catches a typo like
   "MERCURY" vs "MERCUREY"). A pair with no such shared word goes to the
   "Generic words only" sheet - still visible, not dropped, just separated
   from the pairs worth a closer look first.

3. The common-words list is shown to the user as a word-frequency report
   (top 40 words across all 382 unique names, with counts) BEFORE being
   finalized - the user picks additions, this script does not guess them.

Two scorers, not one, same reasoning as before: token_sort_ratio alone
scores this project's own "Medipol Pharmaceutica" / "Medipol Pharmaceutical
India Pvt Ltd" test case at 73.7 (misses an 85 threshold); WRatio scores it
90.0 (catches it). A pair is included if EITHER score is >= 85 - "Caught
by" records which.

Two more changes, 2026-10-31, per the user's review:
4. 3-LETTER words (ANG, ZEE, BDR, SAM, ...) now count as content words
   too (previously filtered out with everything under 4 letters) - but
   ONLY on an EXACT match, never fuzzed, since fuzzing something that
   short is unreliable. Also compares each side's content words fused
   together with no spaces (same technique used for drug names in
   build_rc_vs_bans.py), so a name split across words on one side still
   matches a compound word on the other. BHARAT removed from
   GENERIC_WORDS - too common a real distinctive name fragment.
5. Reads output/match_decisions.xlsx (if present) at the start of every
   run and restores the user's own saved "My decision" values by BOTH
   standard names (order-independent) - so decisions survive a re-run
   even as master_list.xlsx changes. This script only READS that file;
   it never writes to it, so the user's own edits there are never at
   risk of being overwritten.

Nothing is merged. Nothing changes any buyer's data. You decide every pair.
"""

from itertools import combinations
from collections import defaultdict
from datetime import date
from pathlib import Path

from rapidfuzz import fuzz
from openpyxl import load_workbook, Workbook

from common import OUT, write_sheet

DECISIONS_PATH = OUT / "match_decisions.xlsx"

THRESHOLD = 85           # token_sort_ratio threshold, as originally specified
WRATIO_THRESHOLD = 90    # WRatio's own threshold - see note below, NOT 85
TODAY = date.today().strftime("%d-%m-%Y")

# WRatio gets its OWN, higher threshold (90, not 85). Found by testing: WRatio
# produces a suspicious FLAT 85.5 score for "Super Formulation(s) Pvt Ltd"
# against dozens of completely unrelated names (confirmed: even a made-up
# "ABC Random Ltd" scores exactly 85.5) - a scorer artifact on short/
# very-different-length strings, not real similarity. Checked the full score
# distribution: there is a hard cluster at exactly 85.5 (59 of the first
# 127 "kept" rows), then nothing between 85.5 and 90, then a clean, normal-
# looking distribution from 90 upward. The 90+ group is ALL genuine matches
# (confirmed by reading every one at exactly 90.0) - including the Medipol
# test case (90.0 exactly) and several more real truncated-name/address-
# merged-into-name cases (Healers Lab, Stallion Laboratories, Zest Pharma,
# Syncom Healthcare). Raising WRatio's bar to 90 removes the artifact
# cluster entirely while keeping everything of value.

# Shown to the user as a frequency report before any addition was made.
# Additions below (HEALTH, CARE, PARENTERALS, SURGICAL, INDUSTRIES,
# INTERNATIONAL, MEDICAL, CHEMICALS, BIO, MEDITECH, BHARAT) were picked by
# the user from that report, 2026-10-31. ANG and BRILLIANT were reviewed and
# deliberately kept OUT - they are real distinctive name fragments, not
# generic words. INDORE, MUMBAI, ROAD, NAGAR were also reviewed and kept
# OUT, but for a different reason: their presence means address text leaked
# into the firm name (a bug, not a generic-word question) - see
# strip_leaked_address() in common.py, which fixes that at the source.
GENERIC_WORDS = {"PHARMACEUTICALS", "PHARMACEUTICAL", "PHARMA", "FORMULATIONS", "FORMULATION",
                 "LABORATORIES", "LABS", "LAB", "HEALTHCARE", "DRUGS", "REMEDIES", "LIFE", "SCIENCES",
                 "SCIENCE", "LIFESCIENCES", "LIFESCIENCE", "BIOTECH", "INDIA", "PVT", "LTD", "PRIVATE",
                 "LIMITED", "LLP", "AND", "CO",
                 "HEALTH", "CARE", "PARENTERALS", "SURGICAL", "INDUSTRIES", "INTERNATIONAL",
                 "MEDICAL", "CHEMICALS", "BIO", "MEDITECH"}
# BHARAT removed 2026-10-31 per the user's review - too common a real
# distinctive fragment of actual firm names to treat as generic filler.


def content_words(standard_name_text):
    """Words left after stripping generic terms. 4+ letter words are kept
    for FUZZY comparison (a typo like MERCURY/MERCUREY should still
    match); 3-letter words are ALSO kept (added 2026-10-31 - ANG, ZEE,
    BDR, SAM etc. are real distinctive fragments, not noise) but only
    count as shared on an EXACT match - fuzzing a 3-letter word is too
    unreliable to trust (almost anything scores close to anything else
    that short)."""
    return [w for w in (standard_name_text or "").split() if w not in GENERIC_WORDS and len(w) >= 3]


def has_fuzzy_shared_word(words_a, words_b):
    for wa in words_a:
        for wb in words_b:
            if len(wa) == 3 or len(wb) == 3:
                if wa == wb:
                    return True
            elif fuzz.ratio(wa, wb) >= THRESHOLD:
                return True
    # Also compare each side's words fused together with no spaces, so a
    # name split across several words on one side matches a compound
    # word on the other (e.g. "BDR PHARMA" vs some name containing
    # "BDRPHARMA" as one token) - added 2026-10-31 per the user's review,
    # same technique used for drug names in build_rc_vs_bans.py.
    fused_a, fused_b = "".join(words_a), "".join(words_b)
    if len(fused_a) >= 6 and len(fused_b) >= 6 and (fused_a in fused_b or fused_b in fused_a):
        return True
    return False


def report_word_frequency(unique_names):
    counter = defaultdict(int)
    for name in unique_names:
        for w in set(name.split()):
            counter[w] += 1
    print(f"\nTop 40 words across {len(unique_names)} unique standard names "
          f"(for you to review - nothing added automatically):")
    for word, count in sorted(counter.items(), key=lambda kv: -kv[1])[:40]:
        flag = " (already generic)" if word in GENERIC_WORDS else ""
        print(f"  {count:5d}  {word}{flag}")


def load_decisions():
    """Reads output/match_decisions.xlsx (keyed by BOTH standard names,
    order-independent) if it exists - the user's own saved decisions,
    kept across runs even as master_list.xlsx changes and pairs get
    re-ordered or re-scored. Returns {frozenset({name_a, name_b}): decision}.
    Added 2026-10-31 per the user's review - decisions must never be
    lost just because this script was re-run."""
    if not DECISIONS_PATH.exists():
        print(f"\nNo {DECISIONS_PATH.name} found yet - every row's 'My decision' will be blank "
              f"this run (not an error - it gets created the first time you save decisions there).")
        return {}
    wb = load_workbook(DECISIONS_PATH, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}
    decisions = {}
    for r in rows[1:]:
        a, b, decision = r[idx["Standard name A"]], r[idx["Standard name B"]], r[idx["My decision"]]
        if a and b and decision:
            decisions[frozenset((a, b))] = decision
    print(f"\nLoaded {len(decisions)} saved decision(s) from {DECISIONS_PATH.name}.")
    return decisions


def build_groups():
    wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    idx = {name: i for i, name in enumerate(header)}
    id_i, name_i, std_i, buyer_i = idx["Entry ID"], idx["Firm name (as written)"], \
        idx["Firm name (standard)"], idx["Buyer"]

    groups = {}
    for r in data:
        std = r[std_i]
        if not std:
            continue
        g = groups.setdefault(std, {"entry_ids": [], "buyers": set(), "as_written": set()})
        g["entry_ids"].append(r[id_i])
        g["buyers"].add(r[buyer_i])
        g["as_written"].add(r[name_i])
    return groups


def main():
    decisions = load_decisions()
    groups = build_groups()
    unique_names = sorted(groups)
    report_word_frequency(unique_names)

    main_rows, generic_rows = [], []
    checked, restored = 0, 0
    for name_a, name_b in combinations(unique_names, 2):
        ga, gb = groups[name_a], groups[name_b]
        if ga["buyers"] == gb["buyers"]:
            continue                                        # not a cross-buyer pair at all
        checked += 1
        ts_score = fuzz.token_sort_ratio(name_a, name_b)
        wr_score = fuzz.WRatio(name_a, name_b)
        ts_hit, wr_hit = ts_score >= THRESHOLD, wr_score >= WRATIO_THRESHOLD
        if not (ts_hit or wr_hit):
            continue
        caught_by = "both" if (ts_hit and wr_hit) else ("WRatio" if wr_hit else "token_sort")

        words_a, words_b = content_words(name_a), content_words(name_b)
        kept = has_fuzzy_shared_word(words_a, words_b)

        decision = decisions.get(frozenset((name_a, name_b)), "")
        if decision:
            restored += 1
        row = [name_a, "; ".join(sorted(ga["as_written"])), ", ".join(sorted(ga["buyers"])),
               ", ".join(ga["entry_ids"]),
               name_b, "; ".join(sorted(gb["as_written"])), ", ".join(sorted(gb["buyers"])),
               ", ".join(gb["entry_ids"]),
               round(ts_score, 1), round(wr_score, 1), caught_by, decision]
        (main_rows if kept else generic_rows).append((row, caught_by, max(ts_score, wr_score)))

    catch_rank = {"both": 0, "WRatio": 1, "token_sort": 2}
    for group in (main_rows, generic_rows):
        group.sort(key=lambda t: (catch_rank[t[1]], -t[2]))

    kept_wr_scores = [r[9] for r, _, _ in main_rows]  # WRatio is column index 9 in `row`
    at_85 = sum(1 for s in kept_wr_scores if s >= 85)
    at_90 = sum(1 for s in kept_wr_scores if s >= 90)
    at_95 = sum(1 for s in kept_wr_scores if s >= 95)

    print(f"\nUnique standard names: {len(unique_names)}")
    print(f"Unique-name pairs checked (different buyers): {checked}")
    print(f"Kept (Match candidates - share a real word): {len(main_rows)}")
    print(f"Generic words only (separate sheet): {len(generic_rows)}")
    print(f"Saved decisions restored into this run: {restored} of {len(decisions)} loaded")
    print(f"Within kept group, WRatio >= 85: {at_85}  >= 90: {at_90}  >= 95: {at_95}")

    header_out = ["Standard name A", "As written A", "Buyers A", "Entry IDs A",
                  "Standard name B", "As written B", "Buyers B", "Entry IDs B",
                  "token_sort_ratio", "WRatio", "Caught by", "My decision"]

    wb = Workbook()
    write_sheet(wb.active, header_out, [r for r, _, _ in main_rows],
                [30, 34, 18, 30, 30, 34, 18, 30, 14, 10, 12, 14])
    wb.active.title = "Match candidates"
    write_sheet(wb.create_sheet("Generic words only"), header_out, [r for r, _, _ in generic_rows],
                [30, 34, 18, 30, 30, 34, 18, 30, 14, 10, 12, 14])
    out = OUT / "match_candidates.xlsx"
    wb.save(out)
    print(f"Saved: {out}")

    def find(rows, needle_a, needle_b):
        return [r for r, _, _ in rows if needle_a in r[0].lower() and needle_b in r[4].lower()
                or needle_a in r[4].lower() and needle_b in r[0].lower()]

    medipol = find(main_rows, "medipol", "medipol") + find(generic_rows, "medipol", "medipol")
    print("\nMedipol test:")
    if medipol:
        for r in medipol:
            where = "Match candidates" if any(r is x for x, _, _ in main_rows) else "Generic words only"
            print(f"  [{where}] {r[0]!r} vs {r[4]!r} - token_sort={r[8]} WRatio={r[9]} caught_by={r[10]}")
    else:
        print("  WARNING - not found above threshold at all - check manually.")

    super_form = find(main_rows, "super formulation", "") + find(generic_rows, "super formulation", "")
    print("\nSuper Formulations test:")
    if super_form:
        for r in super_form:
            where = "Match candidates" if any(r is x for x, _, _ in main_rows) else "Generic words only"
            print(f"  [{where}] {r[0]!r} vs {r[4]!r} - token_sort={r[8]} WRatio={r[9]} caught_by={r[10]}")
    else:
        print("  (no pair found above threshold for this name)")


if __name__ == "__main__":
    main()
