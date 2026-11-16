"""
11_cross_buyer_pairs.py - break every cross-buyer Match into buyer-pair rows.

Reads:
  output/master_list.xlsx   (made by 10_merge_master.py - "Master list" and
                              "Matches" sheets)
Writes:
  output/cross_buyer_cases.xlsx  - adds/replaces a "Cross-buyer pairs" sheet
                                    (keeps any other sheet already there,
                                    e.g. the hand-curated "Cross-buyer
                                    cases" sheet, untouched)

A Matches-sheet group can span 2 or more buyers (e.g. a firm appearing
under GMSCL, MPPHSCL AND RMSC). This turns every such group into ONE ROW
PER PAIR of buyers (3 buyers -> 3 pairs), each summarising both sides:
Level(s), Reason category(ies), From/To dates, and whether the address is
known to match. Pairs where BOTH sides include a "Quality" reason are
sorted first - those are the most interesting ones to check by hand (two
buyers both found a quality problem with the same firm).

Run this by hand, or just run 10_merge_master.py, which calls this
automatically after it finishes (so it always reflects the latest buyers).
"""

from itertools import combinations
from collections import defaultdict

from openpyxl import load_workbook

from common import OUT, write_sheet


def build_pairs():
    wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    by_id = {r[0]: r for r in data}

    addr_idx = header.index("Firm address")
    buyer_idx = header.index("Buyer")
    level_idx = header.index("Level (firm / product)")
    cat_idx = header.index("Reason category (Quality / Non-supply / Other / To check)")
    from_idx = header.index("Blacklisted from")
    to_idx = header.index("Blacklisted to")

    match_rows = list(wb["Matches"].iter_rows(values_only=True))[1:]

    pair_rows = []
    for firm, times_listed, entry_ids_text, buyers_text, meaning in match_rows:
        if not meaning or not str(meaning).startswith("DIFFERENT"):
            continue
        entry_ids = [e.strip() for e in entry_ids_text.split(",")]
        by_buyer = defaultdict(list)
        for eid in entry_ids:
            r = by_id.get(eid)
            if r is not None:
                by_buyer[r[buyer_idx]].append(r)

        for b1, b2 in combinations(sorted(by_buyer), 2):
            side1, side2 = by_buyer[b1], by_buyer[b2]

            def side_vals(side, idx):
                return ", ".join(sorted({(r[idx] or "").strip() for r in side}))

            def side_dates(side):
                return "; ".join(sorted({f"{(r[from_idx] or '').strip()} to {(r[to_idx] or '').strip()}"
                                           for r in side}))

            def side_addrs(side):
                return {(r[addr_idx] or "").strip() for r in side if (r[addr_idx] or "").strip()
                         and (r[addr_idx] or "").strip().lower() != "not stated"}

            cats1 = {(r[cat_idx] or "").strip() for r in side1}
            cats2 = {(r[cat_idx] or "").strip() for r in side2}
            quality_pair = "Quality" in cats1 and "Quality" in cats2

            a1, a2 = side_addrs(side1), side_addrs(side2)
            same_address = "unknown" if (not a1 or not a2) else ("yes" if a1 & a2 else "no")

            pair_rows.append([
                quality_pair, firm, b1, b2,
                side_vals(side1, level_idx), side_vals(side2, level_idx),
                side_vals(side1, cat_idx), side_vals(side2, cat_idx),
                side_dates(side1), side_dates(side2),
                same_address,
            ])

    pair_rows.sort(key=lambda r: (not r[0], r[1], r[2], r[3]))
    return [r[1:] for r in pair_rows], sum(1 for r in pair_rows if r[0])


def main():
    final_rows, quality_count = build_pairs()
    print(f"Total pairwise cross-buyer rows: {len(final_rows)}")
    print(f"Quality+Quality pairs: {quality_count}")

    pair_header = ["Firm", "Buyer A", "Buyer B", "Level A", "Level B", "Reason category A",
                   "Reason category B", "Dates A (from to)", "Dates B (from to)", "Same address?"]

    out_path = OUT / "cross_buyer_cases.xlsx"
    wb2 = load_workbook(out_path)
    if "Cross-buyer pairs" in wb2.sheetnames:
        del wb2["Cross-buyer pairs"]
    ws2 = wb2.create_sheet("Cross-buyer pairs")
    write_sheet(ws2, pair_header, final_rows, [40, 12, 12, 30, 30, 20, 20, 30, 30, 14])
    wb2.save(out_path)
    print(f"Saved: {out_path}  ('Cross-buyer pairs' sheet, {len(final_rows)} rows)")


if __name__ == "__main__":
    main()
