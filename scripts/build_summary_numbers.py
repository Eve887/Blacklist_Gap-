"""
build_summary_numbers.py - Part 1 (data science), FINAL STAGE. READ-ONLY
REPORT, using ONLY output/frozen_2026-11-07/ (master_list.xlsx and
concealment_cases.xlsx), per the user's instruction.

Writes output/analysis/summary_numbers.xlsx with one sheet per number
asked for:
  - Firms by buyer count   - firms with an active ban (see
    analysis_common.is_active_ban) from 1, 2, or 3+ DISTINCT buyers.
  - Q1 vs Q2               - rows and distinct firms, each level, ALL
    rows (not just active - a Q1/Q2 finding is a fact about the order
    letter regardless of whether the ban was later stayed/revoked).
  - Stayed or revoked by buyer - rows EXCLUDED by is_active_ban() (a
    genuine court stay/revocation/supersession - NOT a plain "period
    completed" note, see analysis_common's own comment), grouped by
    buyer.
  - First-to-second ban gap - for firms with 2+ DISTINCT ban dates
    (any row, active or not - a later-stayed ban was still a real ban
    EVENT at the time), the gap in days between the 1st and 2nd date,
    median + firm count.
  - Reason categories by buyer - a crosstab, ALL rows (Level != Lab).
  - Concealment cases - the 5 rows from concealment_cases.xlsx, as is.

Every sheet's own scope (active-only vs all rows) is stated in its own
first row, not left for the reader to guess.
"""

import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook

from common import write_sheet
from analysis_common import (ANALYSIS_DIR, load_master_rows, load_concealment_cases,
                              is_active_ban, parse_date, HEADLINE_FIRMS)


def main():
    idx, rows = load_master_rows()
    FIRM = idx["Firm name (standard)"]
    BUYER = idx["Buyer"]
    LEVEL = idx["Level (firm / product)"]
    CAT = idx["Reason category (Quality / Non-supply / Other / To check)"]
    COURT = idx["Court stay or revoked?"]
    QLEVEL = idx["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"]
    FROM = idx["Blacklisted from"]

    firm_rows = [r for r in rows if r[LEVEL] != "Lab"]

    wb = Workbook()
    wb.remove(wb.active)

    # --- 1. Firms by buyer count (active bans only) ---
    buyers_by_firm = defaultdict(set)
    for r in firm_rows:
        if is_active_ban(r[COURT], r[LEVEL]):
            buyers_by_firm[r[FIRM]].add(r[BUYER])
    bucket_counts = Counter()
    bucket_rows = []
    for firm, buyers in sorted(buyers_by_firm.items()):
        n = len(buyers)
        bucket = "1 buyer" if n == 1 else ("2 buyers" if n == 2 else "3+ buyers")
        bucket_counts[bucket] += 1
        bucket_rows.append([firm, n, ", ".join(sorted(buyers)),
                            "YES - headline case" if firm in HEADLINE_FIRMS else ""])
    summary1 = [["1 buyer", bucket_counts["1 buyer"]],
                ["2 buyers", bucket_counts["2 buyers"]],
                ["3+ buyers", bucket_counts["3+ buyers"]],
                ["TOTAL distinct firms with an active ban", len(buyers_by_firm)]]
    ws1 = wb.create_sheet("Firms by buyer count")
    write_sheet(ws1, ["Bucket (active bans only - see module docstring)", "Firms"],
                summary1, [44, 12])
    ws1b = wb.create_sheet("Firms by buyer count (detail)")
    write_sheet(ws1b, ["Firm name (standard)", "Number of distinct buyers (active bans)",
                       "Buyers", "Headline case?"], bucket_rows, [36, 16, 50, 20])

    # --- 2. Q1 vs Q2 (rows and firms), ALL rows ---
    q_rows = Counter(r[QLEVEL] for r in firm_rows)
    q_firms = defaultdict(set)
    for r in firm_rows:
        q_firms[r[QLEVEL]].add(r[FIRM])
    q2_rows_summary = [
        ["Q1 (named statutory lab)", q_rows.get("Q1", 0), len(q_firms.get("Q1", set()))],
        ["Q2 (buyer's own/empanelled lab, or no lab named)", q_rows.get("Q2", 0),
         len(q_firms.get("Q2", set()))],
        ["N/A (not a Quality-category row)", q_rows.get("N/A", 0), len(q_firms.get("N/A", set()))],
    ]
    write_sheet(wb.create_sheet("Q1 vs Q2"),
                ["Quality level (ALL rows, not just active)", "Rows", "Distinct firms"],
                q2_rows_summary, [50, 10, 16])

    # --- 3. Stayed or revoked, by buyer ---
    excluded = [r for r in firm_rows if not is_active_ban(r[COURT], r[LEVEL])]
    by_buyer = Counter(r[BUYER] for r in excluded)
    stay_rows = [[buyer, n] for buyer, n in sorted(by_buyer.items(), key=lambda x: -x[1])]
    stay_rows.append(["TOTAL", sum(by_buyer.values())])
    write_sheet(wb.create_sheet("Stayed or revoked by buyer"),
                ["Buyer", "Rows genuinely stayed/revoked/superseded (not a plain "
                 "'period completed' note - see module docstring)"], stay_rows, [16, 60])

    # --- 4. First-to-second ban gap (days) ---
    dates_by_firm = defaultdict(set)
    for r in firm_rows:
        d = parse_date(r[FROM])
        if d:
            dates_by_firm[r[FIRM]].add(d)
    gaps, gap_detail = [], []
    for firm, dates in sorted(dates_by_firm.items()):
        ds = sorted(dates)
        if len(ds) < 2:
            continue
        gap = (ds[1] - ds[0]).days
        gaps.append(gap)
        gap_detail.append([firm, ds[0].strftime("%d-%m-%Y"), ds[1].strftime("%d-%m-%Y"), gap])
    median_gap = round(statistics.median(gaps), 1) if gaps else None
    write_sheet(wb.create_sheet("First-to-second ban gap"),
                ["Median days between a firm's first and second distinct ban date",
                 "Number of firms with 2+ distinct ban dates (any row, active or not)"],
                [[median_gap, len(gaps)]], [50, 50])
    write_sheet(wb.create_sheet("Ban gap (detail)"),
                ["Firm name (standard)", "1st ban date", "2nd ban date", "Gap (days)"],
                sorted(gap_detail, key=lambda r: r[3]), [36, 16, 16, 12])

    # --- 5. Reason categories by buyer ---
    buyers_sorted = sorted({r[BUYER] for r in firm_rows})
    cats_sorted = ["Quality", "Non-supply", "Other", "To check"]
    cross_rows = []
    for cat in cats_sorted:
        row = [cat] + [sum(1 for r in firm_rows if r[BUYER] == b and r[CAT] == cat)
                       for b in buyers_sorted]
        row.append(sum(row[1:]))
        cross_rows.append(row)
    total_row = ["TOTAL"] + [sum(1 for r in firm_rows if r[BUYER] == b) for b in buyers_sorted]
    total_row.append(sum(total_row[1:]))
    cross_rows.append(total_row)
    write_sheet(wb.create_sheet("Reason categories by buyer"),
                ["Reason category"] + buyers_sorted + ["TOTAL"], cross_rows,
                [14] + [12] * len(buyers_sorted) + [10])

    # --- 6. Concealment cases (as-is, from the frozen folder) ---
    c_idx, c_rows = load_concealment_cases()
    c_header = list(c_idx.keys())
    write_sheet(wb.create_sheet("Concealment cases"), c_header, [list(r) for r in c_rows],
                [26, 16, 50, 22, 24, 16, 60, 40][:len(c_header)])

    out = ANALYSIS_DIR / "summary_numbers.xlsx"
    wb.save(out)
    print(f"Saved: {out}")
    print(f"  Firms by buyer count: {dict(bucket_counts)} (total {len(buyers_by_firm)})")
    print(f"  Q1 rows={q_rows.get('Q1',0)} firms={len(q_firms.get('Q1',set()))}; "
          f"Q2 rows={q_rows.get('Q2',0)} firms={len(q_firms.get('Q2',set()))}")
    print(f"  Stayed/revoked rows: {sum(by_buyer.values())} across {len(by_buyer)} buyers")
    print(f"  First-to-second ban gap: median {median_gap} days, {len(gaps)} firms")


if __name__ == "__main__":
    main()
