"""
12_candidate_gaps.py - Step 4 first pass: candidate buyer-coordination gaps.

Reads:
  output/master_list.xlsx   (made by 10_merge_master.py)
Writes:
  output/candidate_gaps.xlsx  with two sheets:
    - Candidate gaps : one row per candidate - buyer B found a Quality
                        problem on a date that falls inside buyer A's
                        already-active ban period for the same (name-
                        matched) firm. B could only have found that if it
                        was still buying from the firm - a candidate for
                        "A had already banned them, B kept buying".
    - By firm        : one row per DISTINCT firm with at least one
                        candidate, so you can see which firms have the
                        most/strongest evidence without re-reading every
                        row.

This is a REPORT ONLY - it never merges, bans, or changes any buyer's data.
Every row says explicitly why it is only a candidate, not a conclusion
(name-match only, proxy dates, Level mismatches, not checked against the
original order PDFs).

Run this by hand, or just run 10_merge_master.py, which calls this
automatically (after 11_cross_buyer_pairs.py) so it always reflects the
latest buyers.
"""

import re
from itertools import permutations
from collections import defaultdict
from datetime import datetime

from openpyxl import load_workbook, Workbook

from common import OUT, write_sheet

DATE_FMT_CANDIDATES = ["%d-%m-%Y", "%d-%b-%Y"]
ORDER_DATE_RE = re.compile(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})")


def parse_date(text):
    text = (text or "").strip()
    for fmt in DATE_FMT_CANDIDATES:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def order_date_fallback(where_text):
    """Pull a date out of the 'Page / row in source' order-reference text
    (e.g. 'Order 554/5.6.2026 (...)' -> 2026-06-05) - used ONLY when the
    real Blacklisted-from field is 'Not stated', as a rougher proxy."""
    m = ORDER_DATE_RE.search(where_text or "")
    if not m:
        return None
    d, mo, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    try:
        return datetime(int(y), int(mo), int(d))
    except ValueError:
        return None


def find_candidates():
    wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    by_id = {r[0]: r for r in data}

    id_idx = header.index("Entry ID")
    firm_idx = header.index("Firm name (as written)")
    buyer_idx = header.index("Buyer")
    level_idx = header.index("Level (firm / product)")
    prod_idx = header.index("Product(s)")
    from_idx = header.index("Blacklisted from")
    to_idx = header.index("Blacklisted to")
    cat_idx = header.index("Reason category (Quality / Non-supply / Other / To check)")
    reason_idx = header.index("Reason (as written)")
    where_idx = header.index("Page / row in source")

    def effective_from(r):
        d = parse_date(r[from_idx])
        if d:
            return d, False
        d = order_date_fallback(r[where_idx])
        return d, True

    def effective_to(r):
        d = parse_date(r[to_idx])
        return d

    match_rows = list(wb["Matches"].iter_rows(values_only=True))[1:]

    candidates = []
    for firm, times_listed, entry_ids_text, buyers_text, meaning in match_rows:
        if not meaning or not str(meaning).startswith("DIFFERENT"):
            continue
        entry_ids = [e.strip() for e in entry_ids_text.split(",")]
        entries = [by_id[e] for e in entry_ids if e in by_id]

        for a, b in permutations(entries, 2):
            if a[buyer_idx] == b[buyer_idx]:
                continue
            if b[cat_idx] != "Quality":
                continue
            b_date, b_is_proxy = effective_from(b)
            a_from, a_from_is_proxy = effective_from(a)
            a_to = effective_to(a)
            if b_date is None or a_to is None:
                continue
            lower_ok = (a_from is None) or (a_from <= b_date)
            if not (lower_ok and b_date <= a_to):
                continue

            caveats = ["Firms matched by standardized NAME only, not a confirmed identity "
                       "(address mostly unknown - see 'Same address?' in cross_buyer_cases.xlsx)."]
            if a_from is None:
                caveats.append(f"{a[buyer_idx]}'s own From-date is 'Not stated' - its ban start is "
                                "unbounded in this check (only the To-date was used as the upper bound).")
            elif a_from_is_proxy:
                caveats.append(f"{a[buyer_idx]}'s From-date was recovered from its order-reference "
                                "text, not a dedicated From-date column - a rougher proxy.")
            if b_is_proxy:
                caveats.append(f"{b[buyer_idx]}'s ban date was recovered from its order-reference "
                                "text (no dedicated From-date column on that site), not the actual "
                                "incident/testing date - a rougher proxy, could be off by some days.")
            if a[level_idx].split()[0] != b[level_idx].split()[0]:
                caveats.append(f"Level differs ({a[level_idx]} vs {b[level_idx]}) - a firm-wide ban "
                                "and a single-product ban are not the same scope.")
            caveats.append("Not independently verified against the original order PDFs for either side.")

            b_date_display = b[from_idx] if not b_is_proxy else f"(from order ref) {b_date.strftime('%d-%m-%Y')}"
            candidates.append({
                "firm": firm, "a_buyer": a[buyer_idx], "a_id": a[id_idx], "a_from": a[from_idx],
                "a_to": a[to_idx], "a_cat": a[cat_idx], "b_buyer": b[buyer_idx], "b_id": b[id_idx],
                "b_date_display": b_date_display, "b_date_dt": b_date, "b_reason": b[reason_idx],
                "b_product": b[prod_idx], "caveats": caveats, "num_caveats": len(caveats),
            })
    return candidates


def build_rows(candidates):
    return [[c["firm"], c["a_buyer"], c["a_id"], c["a_from"], c["a_to"], c["a_cat"],
             c["b_buyer"], c["b_id"], c["b_date_display"], c["b_reason"], c["b_product"],
             " | ".join(c["caveats"])] for c in candidates]


def build_by_firm(candidates):
    by_firm = defaultdict(list)
    for c in candidates:
        by_firm[c["firm"]].append(c)

    rows = []
    for firm, items in sorted(by_firm.items(), key=lambda kv: -len(kv[1])):
        buyers = sorted({c["a_buyer"] for c in items} | {c["b_buyer"] for c in items})
        dated = [c for c in items if c["b_date_dt"] is not None]
        earliest = min((c["b_date_dt"] for c in dated), default=None)
        latest = max((c["b_date_dt"] for c in dated), default=None)
        strongest = min(items, key=lambda c: (c["num_caveats"], c["b_date_dt"] or datetime.max))
        rows.append([
            firm, len(items), ", ".join(buyers),
            earliest.strftime("%d-%m-%Y") if earliest else "Not stated",
            latest.strftime("%d-%m-%Y") if latest else "Not stated",
            strongest["b_id"] + f" (vs {strongest['a_id']}, {strongest['num_caveats']} caveats)",
        ])
    return rows


def main():
    candidates = find_candidates()
    print(f"Candidate gap count: {len(candidates)}")

    by_firm_rows = build_by_firm(candidates)
    print(f"Distinct firms with a candidate gap: {len(by_firm_rows)}")

    cand_header = ["Firm", "Buyer A (already banned)", "A Entry ID", "A From", "A To", "A Reason category",
                   "Buyer B (found Quality issue)", "B Entry ID", "B ban/incident date", "B Reason (as written)",
                   "B Product", "Why this is only a candidate"]
    by_firm_header = ["Firm", "Number of candidate rows", "Buyers involved", "Earliest date", "Latest date",
                       "Strongest row's ID"]

    wb = Workbook()
    write_sheet(wb.active, cand_header, build_rows(candidates),
                [30, 16, 24, 13, 13, 16, 16, 24, 20, 30, 30, 90])
    wb.active.title = "Candidate gaps"
    write_sheet(wb.create_sheet("By firm"), by_firm_header, by_firm_rows, [30, 18, 30, 14, 14, 46])
    out = OUT / "candidate_gaps.xlsx"
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
