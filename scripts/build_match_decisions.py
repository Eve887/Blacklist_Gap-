"""
build_match_decisions.py - one-off conversion: copies the user's "My
decision" values from their review file (output/first_pass_review_2026-10-31.xlsx,
sheets "Match decisions (kept)" and "Match decisions (generic)") into
output/match_decisions.xlsx, keyed by the CORRECT standardized names -
14_match_firms.py reads that file on every run (see its own
load_decisions()), so this is what actually makes the decisions persist.

NOT run automatically, and not something 14_match_firms.py itself does -
a one-off, by-hand step each time the user hands over a fresh reviewed
file, per Rule B backed up like any other script.

Why a join by Entry ID, not by name text: the review file's own "As
written A"/"Entry IDs A" columns hold only ONE representative spelling
and (for side A) a single Entry ID, not the full standardized-name
group match_candidates.xlsx itself uses - matching by re-deriving
standard_name() from free text risks a subtle mismatch (plural/suffix
differences) that would silently fail to restore a decision. Entry IDs
are exact, stable identifiers - this script finds which
match_candidates.xlsx row contains ANY of the review row's own Entry
IDs on either side, and uses THAT row's "Standard name A"/"B" (correctly
oriented, since a review row might have A/B swapped relative to
match_candidates.xlsx) as the real key.

Writes:
  output/match_decisions.xlsx
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

from common import OUT, write_sheet

REVIEW_PATH = OUT / "first_pass_review_2026-10-31.xlsx"
CANDIDATES_PATH = OUT / "match_candidates.xlsx"
DECISIONS_PATH = OUT / "match_decisions.xlsx"


def load_candidate_lookup():
    """Returns a list of (entry_ids_a_set, entry_ids_b_set, std_a, std_b)
    for every row across BOTH match_candidates.xlsx sheets."""
    wb = load_workbook(CANDIDATES_PATH, read_only=True)
    out = []
    for sheet in ("Match candidates", "Generic words only"):
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        idx = {h: i for i, h in enumerate(header)}
        for r in rows[1:]:
            ids_a = {e.strip() for e in (r[idx["Entry IDs A"]] or "").split(",") if e.strip()}
            ids_b = {e.strip() for e in (r[idx["Entry IDs B"]] or "").split(",") if e.strip()}
            out.append((ids_a, ids_b, r[idx["Standard name A"]], r[idx["Standard name B"]]))
    return out


def find_standard_names(review_ids_a, review_ids_b, candidate_lookup):
    """Returns (std_a, std_b, matched) by finding a match_candidates.xlsx
    row sharing at least one Entry ID with the review row on EACH side -
    checked both ways in case A/B got swapped. None/None/False if no
    confident match was found (never guessed)."""
    for ids_a, ids_b, std_a, std_b in candidate_lookup:
        if review_ids_a & ids_a and review_ids_b & ids_b:
            return std_a, std_b, True
        if review_ids_a & ids_b and review_ids_b & ids_a:
            return std_b, std_a, True
    return None, None, False


def main():
    if not REVIEW_PATH.exists():
        print(f"Review file not found at {REVIEW_PATH} - nothing to do.")
        return

    candidate_lookup = load_candidate_lookup()
    print(f"Loaded {len(candidate_lookup)} candidate rows from {CANDIDATES_PATH.name} for joining.")

    review_wb = load_workbook(REVIEW_PATH, read_only=True)
    decision_rows, unmatched = [], []
    for sheet in ("Match decisions (kept)", "Match decisions (generic)"):
        ws = review_wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        idx = {h: i for i, h in enumerate(header)}
        for r in rows[1:]:
            decision = r[idx["My decision"]]
            if not decision:
                continue
            # "Entry IDs A" is NOT reliably a single ID - some review rows have a
            # comma-joined list there too (found by testing, not assumed) - split
            # it the same way as B, consistently.
            ids_a = {e.strip() for e in (r[idx["Entry IDs A"]] or "").split(",") if e.strip()}
            ids_b = {e.strip() for e in (r[idx["Entry IDs B"]] or "").split(",") if e.strip()}
            std_a, std_b, matched = find_standard_names(ids_a, ids_b, candidate_lookup)
            if matched:
                decision_rows.append([std_a, std_b, decision])
            else:
                unmatched.append((r[idx["As written A"]], r[idx["As written B"]], decision))

    if unmatched:
        print(f"\nWARNING - {len(unmatched)} review row(s) could not be matched to any current "
              f"match_candidates.xlsx row by Entry ID (not written to match_decisions.xlsx - "
              f"the underlying data may have changed since the review was done):")
        for a, b, d in unmatched:
            print(f"   {a!r} vs {b!r} -> {d!r}")

    wb = Workbook()
    write_sheet(wb.active, ["Standard name A", "Standard name B", "My decision"],
               decision_rows, [34, 34, 14])
    wb.active.title = "Decisions"
    wb.save(DECISIONS_PATH)

    print(f"\nSaved: {DECISIONS_PATH}")
    print(f"  Decisions written: {len(decision_rows)}")
    print(f"  Unmatched (not written): {len(unmatched)}")


if __name__ == "__main__":
    main()
